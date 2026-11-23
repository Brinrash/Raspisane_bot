"""Periodic downloader for timetable PDFs published on the college website.

The network/parser part is synchronous on purpose; the MAX bot runs it through
asyncio.to_thread so PDF parsing never blocks message handling.
"""
from __future__ import annotations

import hashlib
import json
import re
from datetime import datetime, timedelta
from email.message import Message
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse
from zoneinfo import ZoneInfo

import requests

from pdf_parser import VERSION as PDF_PARSER_VERSION, load_subject_dictionary, parse_pdf
from reference_match import apply_reference, select_reference_xlsx
from site_downloader import DEFAULT_URL, USER_AGENTS, collect_pdf_links, fetch_page, safe_filename

MAX_PDF_BYTES = 60 * 1024 * 1024


def next_check_at(
    now: datetime,
    *,
    interval_minutes: int = 15,
    start_hour: int = 12,
    end_hour: int = 24,
) -> datetime:
    """Return the next local quarter-hour check in [12:00, 24:00).

    When called exactly on a scheduled boundary, return that boundary so a
    freshly started process can perform the check immediately. Subsequent
    calls happen a few seconds later and naturally select the next boundary.
    """
    if interval_minutes <= 0 or 60 % interval_minutes:
        raise ValueError("interval_minutes must be a positive divisor of 60")
    if not (0 <= start_hour < 24 and 1 <= end_hour <= 24 and start_hour < end_hour):
        raise ValueError("Invalid monitoring hours")

    midnight = now.replace(hour=0, minute=0, second=0, microsecond=0)
    start = midnight + timedelta(hours=start_hour)
    end = midnight + timedelta(hours=end_hour)
    if now < start:
        return start
    if now >= end:
        return start + timedelta(days=1)
    if now.minute % interval_minutes == 0 and now.second == 0 and now.microsecond == 0:
        return now

    elapsed_minutes = now.hour * 60 + now.minute
    next_minutes = ((elapsed_minutes // interval_minutes) + 1) * interval_minutes
    candidate = midnight + timedelta(minutes=next_minutes)
    if candidate >= end:
        return start + timedelta(days=1)
    return candidate


def load_site_state(path: str | Path) -> dict[str, Any]:
    file_path = Path(path)
    if not file_path.exists():
        return {"files": {}}
    try:
        data = json.loads(file_path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {"files": {}}
        data.setdefault("files", {})
        return data
    except (OSError, json.JSONDecodeError):
        return {"files": {}}


def save_site_state(path: str | Path, state: dict[str, Any]) -> None:
    file_path = Path(path)
    file_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = file_path.with_suffix(file_path.suffix + ".tmp")
    temporary.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(file_path)


def invalidate_site_cache(state_path: str | Path) -> int:
    """Force a re-parse of previously downloaded PDFs after the XLSX reference changes."""
    state = load_site_state(state_path)
    changed = 0
    for record in (state.get("files") or {}).values():
        if record.get("sha256") and record.get("status") in {"processed", "failed"}:
            record["status"] = "reference_changed"
            changed += 1
    if changed:
        state["reference_changed_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
        save_site_state(state_path, state)
    return changed


def mark_site_file(
    state_path: str | Path,
    url: str,
    record: dict[str, Any],
    *,
    status: str,
    error: str | None = None,
) -> None:
    state = load_site_state(state_path)
    files = state.setdefault("files", {})
    old = files.get(url, {})
    files[url] = {
        **old,
        **record,
        "status": status,
        "processed_at": datetime.now().astimezone().isoformat(timespec="seconds"),
        "error": error,
    }
    # The page can accumulate old attachments. Keep a bounded audit index.
    if len(files) > 250:
        ordered = sorted(
            files.items(),
            key=lambda item: str(item[1].get("processed_at") or item[1].get("checked_at") or ""),
            reverse=True,
        )
        state["files"] = dict(ordered[:250])
    save_site_state(state_path, state)


def prune_site_archive(directory: str | Path, *, max_files: int = 120) -> list[str]:
    """Bound the PDF archive so repeated weekly publications do not fill the host disk."""
    folder = Path(directory)
    if not folder.exists():
        return []
    files = sorted(
        (path for path in folder.glob("*.pdf") if path.is_file()),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    removed: list[str] = []
    for path in files[max(1, max_files):]:
        try:
            path.unlink(missing_ok=True)
            removed.append(path.name)
        except OSError:
            pass
    return removed


def _filename_from_response(response: requests.Response, index: int) -> str:
    name = None
    disposition = response.headers.get("Content-Disposition")
    if disposition:
        msg = Message()
        msg["content-disposition"] = disposition
        name = msg.get_filename()
    if not name:
        url_name = Path(unquote(urlparse(response.url).path)).name
        name = url_name if url_name.casefold().endswith(".pdf") else None
    return safe_filename(name or f"document_{index:02d}", f"document_{index:02d}")


def _read_pdf_response(response: requests.Response) -> bytes:
    declared = response.headers.get("Content-Length")
    if declared and int(declared) > MAX_PDF_BYTES:
        raise ValueError(f"PDF exceeds {MAX_PDF_BYTES // (1024 * 1024)} MB limit")
    chunks: list[bytes] = []
    total = 0
    for chunk in response.iter_content(chunk_size=64 * 1024):
        if not chunk:
            continue
        total += len(chunk)
        if total > MAX_PDF_BYTES:
            raise ValueError(f"PDF exceeds {MAX_PDF_BYTES // (1024 * 1024)} MB limit")
        chunks.append(chunk)
    content = b"".join(chunks)
    if not content.lstrip().startswith(b"%PDF-"):
        preview = content[:100].decode("utf-8", errors="replace").replace("\n", " ")
        raise ValueError(f"URL did not return a PDF (response starts with {preview!r})")
    return content


def scan_site_once(
    *,
    page_url: str = DEFAULT_URL,
    archive_dir: str | Path,
    reference_dir: str | Path,
    state_path: str | Path,
    dictionary_path: str | Path | None = None,
    force: bool = False,
) -> dict[str, Any]:
    """Find new/changed PDFs, download, parse and enrich them with XLSX reference.

    The caller must apply each returned document to the live schedule and then
    call mark_site_file(..., status='processed'). This makes the state update
    transactional with schedule persistence.
    """
    archive = Path(archive_dir)
    archive.mkdir(parents=True, exist_ok=True)
    reference = Path(reference_dir)
    state = load_site_state(state_path)
    session = requests.Session()
    try:
        html, final_page_url = fetch_page(session, page_url)
    except Exception:
        session.close()
        raise
    links = collect_pdf_links(html, final_page_url)
    docs: list[dict[str, Any]] = []
    failures: list[dict[str, Any]] = []
    unchanged = 0
    title_match = re.search(r"<title[^>]*>(.*?)</title>", html, re.I | re.S)
    page_title = re.sub(r"\s+", " ", title_match.group(1)).strip() if title_match else ""
    dictionary_file = Path(dictionary_path) if dictionary_path else Path(__file__).with_name("subjects_dictionary.json")
    dictionary = load_subject_dictionary(dictionary_file)

    for index, item in enumerate(links, start=1):
        url = item["url"]
        previous = (state.get("files") or {}).get(url, {})
        headers = {
            "User-Agent": USER_AGENTS[0],
            "Accept": "application/pdf,*/*;q=0.8",
            "Referer": final_page_url,
        }
        # A new XLSX reference or a new parser version invalidates the old
        # parse even if the PDF bytes are unchanged. Do not send HTTP validators
        # for one scan so the source is downloaded and reparsed.
        reference_was_updated = previous.get("status") == "reference_changed"
        parser_version_matches = previous.get("parser_version") == PDF_PARSER_VERSION
        parser_was_updated = not parser_version_matches
        if not force and not reference_was_updated and not parser_was_updated:
            if previous.get("etag"):
                headers["If-None-Match"] = str(previous["etag"])
            if previous.get("last_modified"):
                headers["If-Modified-Since"] = str(previous["last_modified"])

        response = None
        current_record: dict[str, Any] | None = None
        digest: str | None = None
        try:
            response = session.get(url, headers=headers, timeout=(20, 90), stream=True, allow_redirects=True)
            if response.status_code == 304:
                unchanged += 1
                continue
            response.raise_for_status()
            content = _read_pdf_response(response)
            digest = hashlib.sha256(content).hexdigest()
            record = {
                "sha256": digest,
                "etag": response.headers.get("ETag"),
                "last_modified": response.headers.get("Last-Modified"),
                "filename": _filename_from_response(response, index),
                "final_url": response.url,
                "parser_version": PDF_PARSER_VERSION,
                "checked_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            }
            current_record = record
            if not force and not reference_was_updated and parser_version_matches and previous.get("sha256") == digest and previous.get("status") == "processed":
                unchanged += 1
                # Refresh HTTP validators even if the body is identical.
                mark_site_file(state_path, url, record, status="processed")
                continue
            if not force and not reference_was_updated and parser_version_matches and previous.get("sha256") == digest and previous.get("status") == "failed":
                unchanged += 1
                continue

            stem = Path(record["filename"]).stem
            local_path = archive / f"{stem}_{digest[:8]}.pdf"
            if not local_path.exists():
                local_path.write_bytes(content)

            parsed = parse_pdf(local_path, dictionary)
            reference_path = select_reference_xlsx(local_path, reference)
            if reference_path:
                apply_reference(parsed, reference_path)

            docs.append({
                "url": url,
                "filename": record["filename"],
                "path": str(local_path),
                "sha256": digest,
                "record": record,
                "parsed": parsed,
                "reference_path": str(reference_path) if reference_path else None,
            })
        except Exception as exc:
            failure = {
                "url": url,
                "filename": (current_record or {}).get("filename") or previous.get("filename") or url,
                "error": str(exc),
            }
            if current_record:
                failure["record"] = current_record
            failures.append(failure)
        finally:
            if response is not None:
                response.close()

    # Update page-level audit fields without marking pending PDFs as processed.
    latest = load_site_state(state_path)
    latest["page_url"] = final_page_url
    latest["page_title"] = page_title
    latest["last_checked_at"] = datetime.now().astimezone().isoformat(timespec="seconds")
    latest["pdf_links_found"] = len(links)
    save_site_state(state_path, latest)
    pruned_files = prune_site_archive(archive)
    if pruned_files:
        logger.info("Очистка архива сайта: удалено старых PDF — %d", len(pruned_files))
    session.close()

    return {
        "page_url": final_page_url,
        "page_title": page_title,
        "pdf_links_found": len(links),
        "documents": docs,
        "failures": failures,
        "unchanged": unchanged,
        "archive_pruned": len(pruned_files),
    }
