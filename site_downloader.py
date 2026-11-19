#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Download all PDF files linked from the college timetable page and parse them."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import sys
from datetime import datetime
from email.message import Message
from html import unescape
from pathlib import Path
from urllib.parse import unquote, urljoin, urlparse

import requests
from bs4 import BeautifulSoup

from reference_match import select_reference_xlsx

DEFAULT_URL = (
    "https://spt-11.gosuslugi.ru/svedeniya-ob-obrazovatelnoy-organizatsii/"
    "dokumenty/raspisanie-zanyatiy-s-01092026-po-04092026.html"
)
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:132.0) Gecko/20100101 Firefox/132.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
]
MAX_PDF_BYTES = 60 * 1024 * 1024


def safe_filename(value: str, fallback: str) -> str:
    value = unquote(value or "").strip().strip("\"'")
    value = re.sub(r"[\\/:*?\"<>|]+", "_", value)
    value = re.sub(r"\s+", " ", value).strip(" .")
    if not value:
        value = fallback
    if not value.casefold().endswith(".pdf"):
        value += ".pdf"
    return value[:180]


def content_disposition_filename(header: str | None) -> str | None:
    if not header:
        return None
    msg = Message()
    msg["content-disposition"] = header
    value = msg.get_filename()
    return unquote(value) if value else None


def fetch_page(session: requests.Session, url: str) -> tuple[str, str]:
    last_error: Exception | None = None
    for ua in USER_AGENTS:
        headers = {
            "User-Agent": ua,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.7,en;q=0.6",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        }
        try:
            response = session.get(url, headers=headers, timeout=(15, 45), allow_redirects=True)
            if response.status_code == 403:
                last_error = requests.HTTPError("HTTP 403 Forbidden: сайт отклонил автоматический запрос", response=response)
                continue
            response.raise_for_status()
            if not response.encoding or response.encoding.lower() in {"iso-8859-1", "ascii"}:
                response.encoding = response.apparent_encoding or "utf-8"
            return response.text, response.url
        except Exception as exc:  # show the real network/HTTP issue to the user
            last_error = exc
    if last_error:
        raise last_error
    raise RuntimeError("Не удалось получить страницу сайта")


def _is_pdf_candidate(url: str, tag_text: str = "", parent_text: str = "") -> bool:
    path = urlparse(url).path.casefold()
    combined = (tag_text + " " + parent_text).casefold()
    return path.endswith(".pdf") or ("pdf" in combined and any(word in combined for word in ("скачать", "смотреть", "расписание", "замена")))


def collect_pdf_links(html: str, page_url: str) -> list[dict[str, str]]:
    soup = BeautifulSoup(html, "html.parser")
    candidates: dict[str, dict[str, str]] = {}

    for tag in soup.find_all(["a", "iframe", "embed", "object", "source"]):
        attrs = tag.attrs
        raw_url = attrs.get("href") or attrs.get("src") or attrs.get("data") or attrs.get("data-href") or attrs.get("data-src")
        if isinstance(raw_url, list):
            raw_url = raw_url[0] if raw_url else ""
        if not raw_url:
            onclick = str(attrs.get("onclick") or "")
            match = re.search(r"(?:https?://[^'\"]+|/[^'\"]+\.pdf(?:\?[^'\"]*)?)", onclick, re.I)
            raw_url = match.group(0) if match else ""
        if not raw_url or str(raw_url).startswith(("javascript:", "#", "mailto:")):
            continue
        absolute = urljoin(page_url, unescape(str(raw_url).strip()))
        tag_text = tag.get_text(" ", strip=True)
        parent_text = ""
        parent = tag
        for _ in range(4):
            if parent.parent is None:
                break
            parent = parent.parent
            parent_text = parent.get_text(" ", strip=True)
            if len(parent_text) > 1200:
                parent_text = parent_text[:1200]
                break
        if not _is_pdf_candidate(absolute, tag_text, parent_text):
            continue
        # Avoid collecting unrelated off-site links from a broad page layout.
        host = (urlparse(absolute).hostname or "").lower()
        page_host = (urlparse(page_url).hostname or "").lower()
        if host != page_host and not host.endswith(".gosuslugi.ru"):
            continue
        candidates.setdefault(absolute, {
            "url": absolute,
            "link_text": tag_text,
            "context": parent_text,
        })

    # Some NetCat layouts expose the attachment URL in inline scripts or data attributes.
    for tag in soup.find_all(True):
        for attr, raw in tag.attrs.items():
            values = raw if isinstance(raw, list) else [raw]
            for value in values:
                if not isinstance(value, str):
                    continue
                for match in re.findall(r"(?:https?://[^\s'\"<>]+|/[^\s'\"<>]+)\.pdf(?:\?[^\s'\"<>]*)?", value, re.I):
                    absolute = urljoin(page_url, unescape(match))
                    host = (urlparse(absolute).hostname or "").lower()
                    if host == (urlparse(page_url).hostname or "").lower() or host.endswith(".gosuslugi.ru"):
                        candidates.setdefault(absolute, {"url": absolute, "link_text": "", "context": "Ссылка найдена в атрибуте HTML"})

    return list(candidates.values())


def choose_filename(item: dict[str, str], response: requests.Response, index: int) -> str:
    name = content_disposition_filename(response.headers.get("Content-Disposition"))
    if not name:
        path_name = Path(unquote(urlparse(response.url).path)).name
        name = path_name if path_name and path_name.casefold().endswith(".pdf") else None
    if not name:
        context = item.get("context", "")
        match = re.search(r"([^\n]{5,120}?)\s*(?:PDF|Скачать|Смотреть)", context, re.I)
        name = match.group(1).strip() if match else f"document_{index:02d}"
    return safe_filename(name, f"document_{index:02d}")


def download_pdf(session: requests.Session, item: dict[str, str], out_dir: Path, index: int) -> dict[str, object]:
    response = session.get(
        item["url"],
        headers={"User-Agent": USER_AGENTS[0], "Accept": "application/pdf,*/*;q=0.8", "Referer": DEFAULT_URL},
        timeout=(20, 90),
        stream=True,
        allow_redirects=True,
    )
    response.raise_for_status()
    content_type = response.headers.get("Content-Type", "").lower()
    declared_size = response.headers.get("Content-Length")
    if declared_size and int(declared_size) > MAX_PDF_BYTES:
        raise ValueError(f"Файл больше лимита {MAX_PDF_BYTES // (1024 * 1024)} МБ")

    chunks: list[bytes] = []
    total = 0
    digest = hashlib.sha256()
    prefix = b""
    for chunk in response.iter_content(chunk_size=64 * 1024):
        if not chunk:
            continue
        total += len(chunk)
        if total > MAX_PDF_BYTES:
            raise ValueError(f"Файл превысил лимит {MAX_PDF_BYTES // (1024 * 1024)} МБ")
        if len(prefix) < 16:
            prefix += chunk[:16 - len(prefix)]
        digest.update(chunk)
        chunks.append(chunk)
    data_prefix = prefix.lstrip()
    if not data_prefix.startswith(b"%PDF-"):
        preview = prefix.decode("utf-8", errors="replace").replace("\n", " ")[:100]
        raise ValueError(f"Ссылка не вернула PDF (Content-Type: {content_type or 'не указан'}; начало ответа: {preview!r})")

    filename = choose_filename(item, response, index)
    sha = digest.hexdigest()
    stem = Path(filename).stem
    destination = out_dir / filename
    if destination.exists():
        # Keep multiple versions without overwriting an earlier download.
        destination = out_dir / f"{stem}_{sha[:8]}.pdf"
    with destination.open("wb") as file:
        for chunk in chunks:
            file.write(chunk)

    return {
        "filename": destination.name,
        "url": item["url"],
        "final_url": response.url,
        "size_bytes": total,
        "sha256": sha,
        "content_type": content_type,
        "parsed_text": None,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Скачать PDF расписания с официальной страницы техникума и распознать таблицы")
    parser.add_argument("--url", default=DEFAULT_URL, help="URL страницы с расписанием")
    parser.add_argument("--output", default="downloads", help="Каталог для скачанных PDF")
    parser.add_argument("--no-download", action="store_true", help="Только показать найденные ссылки")
    parser.add_argument("--skip-parse", action="store_true", help="Не запускать локальный PDF-парсер после загрузки")
    parser.add_argument("--reference-dir", help="Папка с эталонными XLSX; выбирается файл по диапазону дат PDF")
    parser.add_argument("--reference-xlsx", help="Один XLSX для сопоставления со всеми найденными PDF")
    args = parser.parse_args()

    session = requests.Session()
    print(f"Страница: {args.url}")
    try:
        html, final_page_url = fetch_page(session, args.url)
    except requests.HTTPError as exc:
        response = getattr(exc, "response", None)
        if response is not None and response.status_code == 403:
            print("ОШИБКА: сайт ответил HTTP 403 Forbidden.")
            print("Это означает, что сервер не разрешил автоматическое чтение страницы с этого компьютера.")
            print("Проверьте, открывается ли страница в браузере; не вводите пароли или токены в этот скрипт.")
        else:
            print(f"ОШИБКА HTTP: {exc}")
        return 2
    except requests.RequestException as exc:
        print(f"ОШИБКА СЕТИ: {exc}")
        print("Проверьте подключение к интернету, DNS и доступность сайта.")
        return 2
    except Exception as exc:
        print(f"ОШИБКА: {exc}")
        return 2

    soup = BeautifulSoup(html, "html.parser")
    title = soup.title.get_text(" ", strip=True) if soup.title else "Без заголовка"
    links = collect_pdf_links(html, final_page_url)
    print(f"Страница открыта: HTTP 200 (после возможного перенаправления)")
    print(f"Заголовок: {title}")
    print(f"PDF-ссылок найдено: {len(links)}")
    for idx, item in enumerate(links, 1):
        print(f"  {idx}. {item['url']}")

    if args.no_download:
        return 0 if links else 3
    if not links:
        print("PDF-ссылки не найдены. Возможно, сайт изменил HTML-разметку или файлы подгружаются JavaScript.")
        return 3

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = Path(args.output).expanduser().resolve() / stamp
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest: dict[str, object] = {
        "page_url": final_page_url,
        "page_title": title,
        "checked_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "pdf_count_found": len(links),
        "files": [],
    }
    downloaded: list[dict[str, object]] = []
    failures: list[dict[str, str]] = []
    for idx, item in enumerate(links, 1):
        try:
            info = download_pdf(session, item, out_dir, idx)
            print(f"[OK] {info['filename']} ({info['size_bytes']} байт)")
            downloaded.append(info)
        except Exception as exc:
            print(f"[FAIL] {item['url']} — {exc}")
            failures.append({"url": item["url"], "error": str(exc)})

    if not args.skip_parse:
        parser_script = Path(__file__).with_name("pdf_parser.py")
        for item in downloaded:
            pdf_path = out_dir / str(item["filename"])
            try:
                parser_command = [sys.executable, str(parser_script), str(pdf_path)]
                reference_path = Path(args.reference_xlsx).expanduser().resolve() if args.reference_xlsx else None
                if reference_path is None and args.reference_dir:
                    reference_path = select_reference_xlsx(pdf_path, Path(args.reference_dir).expanduser().resolve())
                if reference_path is not None:
                    parser_command.extend(["--reference-xlsx", str(reference_path)])
                    print(f"[REFERENCE] {pdf_path.name} ← {reference_path.name}")
                elif args.reference_dir:
                    print(f"[REFERENCE] {pdf_path.name}: XLSX для этой даты не найден")
                result = subprocess.run(
                    parser_command,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    timeout=120,
                )
                item["parsed_text"] = result.stdout.strip()
                item["parse_return_code"] = result.returncode
                if reference_path is not None:
                    item["reference_xlsx"] = reference_path.name
                if result.returncode:
                    item["parse_error"] = result.stderr.strip()[-2000:]
                    print(f"[PARSE-FAIL] {pdf_path.name}: {item['parse_error']}")
                else:
                    summaries = [
                        line.strip() for line in result.stdout.splitlines()
                        if line.strip().startswith(("Pages:", "Recovered subjects:"))
                    ]
                    detail = " | ".join(summaries)
                    print(f"[PARSED] {pdf_path.name}" + (f" | {detail}" if detail else ""))
            except Exception as exc:
                item["parse_error"] = str(exc)
                print(f"[PARSE-FAIL] {pdf_path.name}: {exc}")

    manifest["files"] = downloaded
    manifest["failures"] = failures
    manifest_path = out_dir / "manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print("\nИТОГ")
    print(f"Скачано PDF: {len(downloaded)} из {len(links)}")
    print(f"Ошибок скачивания: {len(failures)}")
    print(f"Папка результатов: {out_dir}")
    print(f"Манифест: {manifest_path}")
    return 0 if downloaded and not failures else 4


if __name__ == "__main__":
    raise SystemExit(main())
