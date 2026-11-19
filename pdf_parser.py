#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Coordinate-aware parser for SPT weekly timetable and daily replacement PDFs.

The parser uses the PDF text layer and vector table borders. It does not run OCR.
It emits one structured record per group/date/pair and retains source fragments
so a human can audit uncertain or clipped cells.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from statistics import median
from typing import Any

from reference_match import apply_reference, select_reference_xlsx

try:
    # Current PyMuPDF import name. This avoids the deprecated legacy `fitz`
    # compatibility package and its warning on newer Python installations.
    import pymupdf as fitz
except ImportError:
    print("PyMuPDF is missing. Install it with: python -m pip install -r requirements.txt", file=sys.stderr)
    raise SystemExit(2)

VERSION = "0.5"
DATE_RE = re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b")
GROUP_RE = re.compile(r"^\d{3}$")
NUM_RE = re.compile(r"^\d{1,2}$")
WEEKDAYS = ["ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА", "ВОСКРЕСЕНЬЕ"]
ROOM_RE = re.compile(r"^(?:\d{1,3}(?:/\d{1,3})?(?:[-–]\d+)?|с\s*/\s*р|дистант)$", re.I)
SUBGROUP_RE = re.compile(r"(?:\(\s*)?([12])\s*п\s*/\s*(?:г)?\s*\)?", re.I)
PARTIAL_SUBGROUP_RE = re.compile(r"\(\s*([12])\s*(?:п\s*/\s*(?:г)?)?\s*$", re.I)
SELF_STUDY_RE = re.compile(r"(?:с\s*/\s*р|самостоятельн\w*\s+работ\w*)", re.I)
SUBGROUP_SUFFIX_RE = re.compile(
    r"(?:\(\s*|\s+)([12])\s*п\s*/\s*(?:г)?\s*\)?\s*$", re.I
)
PARTIAL_SUBGROUP_SUFFIX_RE = re.compile(r"\(\s*([12])\s*$", re.I)
PARTIAL_SUBJECT_ENDINGS = (
    "культур", "гидравлик", "финансов", "предприни", "электроник",
    "жизн", "строитель", "маш", "средс",
)


def clean_text(value: str) -> str:
    value = value.replace("\u00a0", " ").replace("\u2009", " ").replace("\u202f", " ")
    value = re.sub(r"\s+", " ", value).strip()
    # Repair split punctuation/space artifacts without guessing missing letters.
    value = re.sub(r"\(\s+", "(", value)
    value = re.sub(r"\s+\)", ")", value)
    value = re.sub(r"\s+([,.;:])", r"\1", value)
    return value


def iso_date(value: str) -> str:
    return datetime.strptime(value, "%d.%m.%Y").date().isoformat()


def extract_words(page: fitz.Page) -> list[dict[str, Any]]:
    words = []
    for item in page.get_text("words", sort=True):
        x0, y0, x1, y1, token = item[:5]
        token = str(token).strip()
        if token:
            words.append({"x0": float(x0), "y0": float(y0), "x1": float(x1), "y1": float(y1), "text": token})
    return words


def vertical_rules(page: fitz.Page) -> list[float]:
    """Read vertical table rules drawn as very thin filled rectangles."""
    values: list[float] = []
    for drawing in page.get_drawings():
        rect = drawing.get("rect")
        if rect and rect.height > 5 and rect.width <= 2.0:
            values.append((float(rect.x0) + float(rect.x1)) / 2)
    values.sort()
    merged: list[float] = []
    for value in values:
        if not merged or abs(value - merged[-1]) > 1.4:
            merged.append(value)
        else:
            merged[-1] = (merged[-1] + value) / 2
    return merged


def nearest_left_rule(x: float, rules: list[float], max_gap: float = 12.0) -> float | None:
    candidates = [r for r in rules if 0 <= x - r <= max_gap]
    return max(candidates) if candidates else None


def cluster_by_y(items: list[dict[str, Any]], tolerance: float = 2.5) -> list[list[dict[str, Any]]]:
    clusters: list[list[dict[str, Any]]] = []
    for item in sorted(items, key=lambda z: z["y0"]):
        if not clusters or abs(item["y0"] - median(x["y0"] for x in clusters[-1])) > tolerance:
            clusters.append([item])
        else:
            clusters[-1].append(item)
    return clusters


def page_group_headers(page: fitz.Page, words: list[dict[str, Any]], kind: str, rules: list[float]) -> list[dict[str, Any]]:
    """Find group-number headings and their table x-ranges on a page."""
    if kind == "weekly":
        candidates = [w for w in words if 3 <= w["y0"] <= 102 and w["x0"] > 60 and GROUP_RE.fullmatch(w["text"])]
        # Weekly group numbers are just to the right of their left table rule.
        candidates = [w for w in candidates if (nearest_left_rule(w["x0"], rules) or 0) and w["x0"] - (nearest_left_rule(w["x0"], rules) or 0) < 12]
        clusters = cluster_by_y(candidates, 3.0)
        if not clusters:
            return []
        # Select the main group row (not a stray 3-digit room elsewhere on the page).
        main = max(clusters, key=lambda c: len(c))
        headers = sorted(main, key=lambda w: w["x0"])
        return _build_spans(headers, rules, width_hint=67.45)

    # Daily replacements have 3-digit group labels in several horizontal blocks.
    candidates = [w for w in words if 112 <= w["y0"] <= 420 and 60 <= w["x0"] <= 760 and GROUP_RE.fullmatch(w["text"])]
    candidates = [w for w in candidates if (nearest_left_rule(w["x0"], rules) is not None and 13 <= w["x0"] - float(nearest_left_rule(w["x0"], rules)) <= 35)]
    clusters = [c for c in cluster_by_y(candidates, 3.0) if len(c) >= 2]
    result: list[dict[str, Any]] = []
    for row in clusters:
        result.extend(_build_spans(sorted(row, key=lambda w: w["x0"]), rules, width_hint=87.14))
    # De-duplicate header rows, preserving top-to-bottom order.
    dedup: dict[tuple[str, int], dict[str, Any]] = {}
    for span in result:
        key = (span["group"], round(span["header_y"] / 3))
        dedup[key] = span
    return sorted(dedup.values(), key=lambda s: (s["header_y"], s["x0"]))


def _build_spans(headers: list[dict[str, Any]], rules: list[float], width_hint: float) -> list[dict[str, Any]]:
    starts: list[tuple[dict[str, Any], float]] = []
    for header in headers:
        left = nearest_left_rule(header["x0"], rules, max_gap=18)
        if left is None:
            left = header["x0"] - 2
        starts.append((header, float(left)))
    starts.sort(key=lambda x: x[1])
    # Group columns have a fixed width in each PDF type. Use the median observed
    # step when available, otherwise the known geometric spacing as a fallback.
    diffs = [starts[i + 1][1] - starts[i][1] for i in range(len(starts) - 1) if 30 < starts[i + 1][1] - starts[i][1] < 130]
    width = median(diffs) if diffs else width_hint
    spans: list[dict[str, Any]] = []
    for index, (header, left) in enumerate(starts):
        right = starts[index + 1][1] if index + 1 < len(starts) else left + width
        if right - left < 30 or right - left > 130:
            right = left + width
        spans.append({"group": header["text"], "x0": left, "x1": right, "header_y": float(header["y0"])})
    return spans


def find_room_split(words: list[dict[str, Any]], span: dict[str, Any], kind: str, rules: list[float]) -> float:
    """Find the start of the auditorium subcolumn from the 'Ауд.' header."""
    header_limit = 105 if kind == "weekly" else 110
    aud_words = [
        w for w in words
        if w["text"].casefold().startswith("ауд")
        and span["x0"] - 1 <= w["x0"] <= span["x1"] + 1
        and w["y0"] <= header_limit
    ]
    if aud_words:
        aud = min(aud_words, key=lambda w: abs(((w["x0"] + w["x1"]) / 2) - (span["x0"] + span["x1"]) / 2))
        border = nearest_left_rule(aud["x0"], rules, max_gap=10)
        if border is not None and span["x0"] + 25 < border < span["x1"] - 4:
            return border
        return max(span["x0"] + 30, aud["x0"] - 4.5)
    # Fallback: the room column is a narrow cell at the right of each group.
    return span["x1"] - (16 if kind == "weekly" else 21)


def words_in_band(words: list[dict[str, Any]], x0: float, x1: float, y0: float, y1: float) -> list[dict[str, Any]]:
    return [w for w in words if x0 - 0.4 <= (w["x0"] + w["x1"]) / 2 < x1 + 0.4 and y0 <= (w["y0"] + w["y1"]) / 2 < y1]


def join_words(words: list[dict[str, Any]]) -> str:
    return clean_text(" ".join(w["text"] for w in sorted(words, key=lambda w: (w["y0"], w["x0"]))))


def normalize_subject(value: str) -> str:
    """Normalize a subject name for conservative dictionary matching."""
    value = value.casefold().replace("ё", "е")
    value = re.sub(r"(?:\(\s*|\s+)[12]\s*п\s*/\s*(?:г)?\s*\)?\s*$", "", value, flags=re.I)
    value = re.sub(r"\s*(?:с\s*/\s*р|самостоятельная\s+работа)\s*$", "", value, flags=re.I)
    value = clean_text(value).strip(" .,;:-")
    return value


def load_subject_dictionary(path: Path | None) -> list[str]:
    """Load canonical subject names from the adjacent JSON file."""
    if not path or not path.exists():
        return []
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        values = payload if isinstance(payload, list) else payload.get("subjects", [])
        return sorted({clean_text(str(x)) for x in values if clean_text(str(x))}, key=lambda x: (len(x), x.casefold()))
    except (OSError, json.JSONDecodeError, AttributeError, TypeError):
        return []


def _ends_inside_word(fragment: str, candidate: str) -> bool:
    """True only if the final fragment token is a strict prefix of a word."""
    left = normalize_subject(fragment).split()
    right = normalize_subject(candidate).split()
    if not left or len(right) < len(left):
        return False
    if right[:len(left) - 1] != left[:-1]:
        return False
    fragment_word = left[-1].strip(".,;:")
    candidate_word = right[len(left) - 1].strip(".,;:")
    return (
        bool(fragment_word)
        and len(candidate_word) > len(fragment_word)
        and candidate_word.startswith(fragment_word)
    )


def recover_subject(subject: str, dictionary: list[str]) -> tuple[str, dict[str, Any]]:
    """Restore a clipped tail only if a single canonical title completes a word.

    A phrase that ends after a complete word (for example "Основы экономики")
    is not expanded merely because a longer title begins with that phrase.
    This avoids silently inventing the rest of an ambiguous title.
    """
    normalized = normalize_subject(subject)
    normalized_dictionary = [(item, normalize_subject(item)) for item in dictionary]
    exact = [item for item, key in normalized_dictionary if key == normalized]
    if exact:
        return subject, {"recovered": False, "candidates": [], "ambiguous": False}

    prefix_candidates = [
        item for item, key in normalized_dictionary
        if len(key) > len(normalized) and key.startswith(normalized)
    ]
    candidates = [item for item in prefix_candidates if _ends_inside_word(subject, item)]
    if len(candidates) == 1:
        return candidates[0], {
            "recovered": True, "candidates": candidates, "ambiguous": False,
        }
    if candidates:
        return subject, {"recovered": False, "candidates": candidates, "ambiguous": True}
    if prefix_candidates:
        return subject, {"recovered": False, "candidates": prefix_candidates, "ambiguous": True}

    likely_partial = any(normalized.endswith(ending) for ending in PARTIAL_SUBJECT_ENDINGS)
    return subject, {"recovered": False, "candidates": [], "ambiguous": False, "likely_partial": likely_partial}


def extract_cell_fields(
    words: list[dict[str, Any]], span: dict[str, Any], room_split: float,
    subject_y: float, teacher_y: float, subject_height: float, teacher_height: float,
    subject_dictionary: list[str] | None = None,
) -> dict[str, Any]:
    # Use word centers for the split: glyph boxes sometimes cross the visual
    # divider by a fraction of a point even when the PDF cell is correct.
    body = words_in_band(words, span["x0"] + 0.4, span["x1"] - 0.2, subject_y, subject_y + subject_height)
    teacher_words = words_in_band(words, span["x0"] + 0.4, span["x1"] - 0.2, teacher_y, teacher_y + teacher_height)
    subject_words = [w for w in body if (w["x0"] + w["x1"]) / 2 < room_split]
    room_words = [w for w in body if (w["x0"] + w["x1"]) / 2 >= room_split]
    subject_raw = join_words(subject_words)
    room_raw = join_words(room_words)
    teacher = join_words(teacher_words) or None

    combined = f"{subject_raw} {room_raw}".strip()
    subgroup_match = SUBGROUP_RE.search(combined) or PARTIAL_SUBGROUP_RE.search(subject_raw)
    subgroup = int(subgroup_match.group(1)) if subgroup_match else None
    subgroup_marker_repaired = bool(PARTIAL_SUBGROUP_RE.search(subject_raw))
    self_study = bool(SELF_STUDY_RE.search(combined))

    # In this table the subgroup and self-study markers are metadata, not part
    # of the subject title. Remove them without altering the raw source text.
    subject = subject_raw
    if self_study:
        subject = re.sub(r"\s*(?:с\s*/?\s*р?|самостоятельн\w*\s+работ\w*)\s*$", "", subject, flags=re.I).strip()
        if not subject and room_raw:
            subject = re.sub(r"\s*(?:с\s*/?\s*р?|самостоятельн\w*\s+работ\w*)\s*$", "", room_raw, flags=re.I).strip()
        room = None
    else:
        room = room_raw or None

    # Remove both complete and clipped subgroup suffixes. The subgroup number
    # is preserved separately, and the raw source remains available for audit.
    if SUBGROUP_SUFFIX_RE.search(subject):
        subject = SUBGROUP_SUFFIX_RE.sub("", subject).strip()
    else:
        # Only remove an incomplete marker that actually begins with "(".
        # Do not strip the final digit from module codes such as "МДК 01.01".
        subject = PARTIAL_SUBGROUP_SUFFIX_RE.sub("", subject).strip()

    subject_original = subject
    recovery = {"recovered": False, "candidates": [], "ambiguous": False}
    if subject and subject_dictionary:
        subject, recovery = recover_subject(subject, subject_dictionary)

    normalized = normalize_subject(subject_original)
    likely_partial = bool(recovery.get("likely_partial")) or any(
        normalized.endswith(ending) for ending in PARTIAL_SUBJECT_ENDINGS
    )
    source_text_truncated = bool(recovery.get("recovered") or recovery.get("ambiguous") or likely_partial)
    quality_flags: list[str] = []
    if recovery.get("recovered"):
        quality_flags.append("subject_recovered_from_dictionary")
    if recovery.get("ambiguous"):
        quality_flags.append("ambiguous_subject_prefix")
    if subgroup_marker_repaired:
        quality_flags.append("subgroup_marker_repaired")
    if likely_partial and not recovery.get("recovered"):
        quality_flags.append("possible_subject_truncation")
    if not subject and (room or teacher):
        quality_flags.append("missing_subject")

    needs_review = bool(
        recovery.get("ambiguous")
        or (likely_partial and not recovery.get("recovered"))
        or (not subject and (room or teacher))
    )
    if recovery.get("recovered"):
        quality_status = "recovered"
    elif needs_review:
        quality_status = "review"
    else:
        quality_status = "ok"

    return {
        "subject": subject or None,
        "subject_raw": subject_raw or None,
        "subject_original": subject_original or None,
        "subject_recovered": bool(recovery.get("recovered")),
        "subject_dictionary_match": subject if recovery.get("recovered") else None,
        "subject_dictionary_candidates": recovery.get("candidates", []) if recovery.get("ambiguous") else [],
        "room": room,
        "teacher": teacher,
        "subgroup": subgroup,
        "subgroup_marker_repaired": subgroup_marker_repaired,
        "self_study": self_study,
        "possibly_truncated": source_text_truncated,
        "source_text_truncated": source_text_truncated,
        "needs_review": needs_review,
        "quality_status": quality_status,
        "quality_flags": quality_flags,
        "raw_lines": [x for x in (subject_raw, room_raw, teacher) if x],
    }


def get_page_dates(words: list[dict[str, Any]]) -> list[str]:
    dates = set()
    for w in words:
        if DATE_RE.fullmatch(w["text"]):
            try:
                dates.add(iso_date(w["text"]))
            except ValueError:
                pass
    return sorted(dates)


def parse_weekly_page(page: fitz.Page, page_number: int, subject_dictionary: list[str] | None = None) -> dict[str, Any]:
    words = extract_words(page)
    rules = vertical_rules(page)
    spans = page_group_headers(page, words, "weekly", rules)
    dates = get_page_dates(words)
    page_text = page.get_text("text")
    course_match = re.search(r"([1-5](?:\s*[,и]\s*[1-5])*)\s*курсы", page_text, re.I)
    course = f"{re.sub(r'\s+', '', course_match.group(1))} курсы" if course_match else None
    if not spans:
        return {"page": page_number, "course": course, "groups": [], "entries": [], "warning": "Не найдена строка групп"}

    # Weekly timetable row labels are pair numbers, each followed by subject and teacher lines.
    markers = []
    for w in words:
        if 58 <= w["x0"] <= 67 and 100 <= w["y0"] <= page.rect.height - 20 and NUM_RE.fullmatch(w["text"]):
            value = int(w["text"])
            if 1 <= value <= 8:
                markers.append({"pair": value, "y": w["y0"]})
    markers.sort(key=lambda x: x["y"])
    unique_markers = []
    for marker in markers:
        if not unique_markers or abs(marker["y"] - unique_markers[-1]["y"]) > 3:
            unique_markers.append(marker)

    # A day begins when the numbered rows reset to 1; each daily table has at least 25pt gap.
    day_starts = []
    for marker in unique_markers:
        if marker["pair"] == 1 and (not day_starts or marker["y"] - day_starts[-1]["y"] > 25):
            day_starts.append(marker)
    day_blocks = []
    for index, start in enumerate(day_starts):
        end = day_starts[index + 1]["y"] if index + 1 < len(day_starts) else page.rect.height - 4
        block_markers = [m for m in unique_markers if start["y"] - 1 <= m["y"] < end]
        date_value = dates[index] if index < len(dates) else None
        weekday = WEEKDAYS[datetime.fromisoformat(date_value).weekday()] if date_value else None
        day_blocks.append({"date": date_value, "weekday": weekday, "markers": block_markers})

    entries = []
    groups: list[dict[str, Any]] = []
    first_date = dates[0] if dates else None
    for span in spans:
        split = find_room_split(words, span, "weekly", rules)
        group_days = []
        # Class-hour row is between group header and first numbered lesson.
        if first_date:
            class_hour = extract_cell_fields(words, span, split, 102.0, 108.0, 7.2, 6.0, subject_dictionary)
            if class_hour["subject"] and "классный час" in class_hour["subject"].casefold():
                entry = {
                    "group": span["group"], "date": first_date, "pair": 0,
                    "lesson_numbers": [], "document_type": "weekly", **class_hour,
                    "source_page": page_number,
                }
                entries.append(entry)

        for block in day_blocks:
            day_entries = []
            for marker in block["markers"]:
                fields = extract_cell_fields(
                    words, span, split,
                    subject_y=marker["y"] - 1.6,
                    teacher_y=marker["y"] + 6.4,
                    subject_height=7.8,
                    teacher_height=7.5,
                    subject_dictionary=subject_dictionary,
                )
                if not any((fields.get("subject"), fields.get("room"), fields.get("teacher"))):
                    continue
                entry = {
                    "group": span["group"],
                    "date": block["date"],
                    "pair": marker["pair"],
                    "lesson_numbers": [marker["pair"] * 2 - 1, marker["pair"] * 2],
                    "document_type": "weekly",
                    **fields,
                    "source_page": page_number,
                }
                entries.append(entry)
                day_entries.append(entry)
            group_days.append({"date": block["date"], "weekday": block["weekday"], "entries": day_entries})
        groups.append({"group": span["group"], "course": course, "days": group_days})

    return {"page": page_number, "course": course, "groups": groups, "entries": entries}


def replacement_header_rows(words: list[dict[str, Any]], rules: list[float]) -> list[tuple[float, list[dict[str, Any]]]]:
    candidates = [w for w in words if 112 <= w["y0"] <= 420 and 60 <= w["x0"] <= 760 and GROUP_RE.fullmatch(w["text"])]
    candidates = [
        w for w in candidates
        if (lambda left: left is not None and 0 <= w["x0"] - left <= 4.5)(nearest_left_rule(w["x0"], rules))
    ]
    rows = []
    for cluster in cluster_by_y(candidates, 3.0):
        if len(cluster) >= 2:
            y = median(w["y0"] for w in cluster)
            rows.append((y, sorted(cluster, key=lambda w: w["x0"])))
    return sorted(rows, key=lambda item: item[0])


def parse_replacements_page(page: fitz.Page, page_number: int, subject_dictionary: list[str] | None = None) -> dict[str, Any]:
    words = extract_words(page)
    rules = vertical_rules(page)
    text = page.get_text("text")
    date_match = DATE_RE.search(text)
    date_value = iso_date(date_match.group(0)) if date_match else None
    headers = replacement_header_rows(words, rules)
    if not headers:
        raise ValueError(f"Страница {page_number}: не найдены строки с группами замен")

    group_blocks = []
    entries = []
    for block_index, (header_y, header_words) in enumerate(headers):
        spans = _build_spans(header_words, rules, width_hint=87.14)
        next_header_y = headers[block_index + 1][0] if block_index + 1 < len(headers) else page.rect.height - 2
        pair_markers = []
        for w in words:
            if 40 <= w["x0"] <= 51.5 and header_y + 3 < w["y0"] < next_header_y - 2 and NUM_RE.fullmatch(w["text"]):
                value = int(w["text"])
                if 1 <= value <= 8:
                    pair_markers.append({"pair": value, "y": w["y0"]})
        pair_markers.sort(key=lambda x: x["y"])
        unique = []
        for m in pair_markers:
            if not unique or abs(m["y"] - unique[-1]["y"]) > 3:
                unique.append(m)

        block_groups = []
        for span in spans:
            split = find_room_split(words, span, "replacements", rules)
            group_pairs = []
            for marker in unique:
                fields = extract_cell_fields(
                    words, span, split,
                    subject_y=marker["y"] - 1.7,
                    teacher_y=marker["y"] + 7.5,
                    subject_height=8.8,
                    teacher_height=8.8,
                    subject_dictionary=subject_dictionary,
                )
                if not any((fields.get("subject"), fields.get("room"), fields.get("teacher"))):
                    continue
                entry = {
                    "group": span["group"],
                    "date": date_value,
                    "pair": marker["pair"],
                    "lesson_numbers": [marker["pair"] * 2 - 1, marker["pair"] * 2],
                    "document_type": "replacements",
                    **fields,
                    "source_page": page_number,
                    "group_block": block_index + 1,
                }
                entries.append(entry)
                group_pairs.append(entry)
            block_groups.append({"group": span["group"], "pairs": group_pairs})
        group_blocks.append({"header_y": round(header_y, 2), "groups": block_groups})

    return {"page": page_number, "date": date_value, "group_blocks": group_blocks, "entries": entries}


def detect_document_type(doc: fitz.Document, path: Path) -> str:
    sample = "\n".join(page.get_text("text") for page in doc)
    named = path.name.casefold()
    if "замен" in named or re.search(r"замен[аы]\s+заняти", sample, re.I):
        return "replacements"
    if "распис" in named or re.search(r"расписание\s+занятий", sample, re.I):
        return "weekly"
    raise ValueError("Не удалось определить PDF: не найдено основное расписание или замены")


def parse_pdf(path: Path, subject_dictionary: list[str] | None = None) -> dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(path)
    if path.suffix.lower() != ".pdf":
        raise ValueError("Входной файл должен иметь расширение .pdf")
    doc = fitz.open(path)
    if len(doc) == 0:
        raise ValueError("PDF пустой")
    kind = detect_document_type(doc, path)
    pages = []
    all_entries = []
    all_groups: set[str] = set()
    all_dates: set[str] = set()
    for page_no, page in enumerate(doc, start=1):
        parsed = parse_weekly_page(page, page_no, subject_dictionary) if kind == "weekly" else parse_replacements_page(page, page_no, subject_dictionary)
        pages.append(parsed)
        all_entries.extend(parsed.get("entries", []))
        for group in parsed.get("groups", []):
            all_groups.add(str(group["group"]))
        for block in parsed.get("group_blocks", []):
            all_groups.update(str(g["group"]) for g in block.get("groups", []))
        for entry in parsed.get("entries", []):
            if entry.get("date"):
                all_dates.add(entry["date"])
        if kind == "weekly":
            for group in parsed.get("groups", []):
                for day in group.get("days", []):
                    if day.get("date"):
                        all_dates.add(day["date"])
    replacement_mode = "supplement" if re.search(r"дополн|дополнитель", path.name, re.I) else "replace"
    review_entries = [entry for entry in all_entries if entry.get("needs_review")]
    recovered_entries = [entry for entry in all_entries if entry.get("subject_recovered")]
    repaired_subgroup_entries = [entry for entry in all_entries if entry.get("subgroup_marker_repaired")]
    return {
        "parser_version": VERSION,
        "schema_version": 2,
        "source_file": path.name,
        "document_type": kind,
        "replacement_mode": replacement_mode if kind == "replacements" else None,
        "page_count": len(doc),
        "dates": sorted(all_dates),
        "group_count": len(all_groups),
        "groups_found": sorted(all_groups, key=lambda x: (len(x), x)),
        "entry_count": len(all_entries),
        "quality_summary": {
            "needs_review": len(review_entries),
            "subject_recovered": len(recovered_entries),
            "subgroup_markers_repaired": len(repaired_subgroup_entries),
        },
        "entries": all_entries,
        "pages": pages,
        "limitations": [
            "Парсер не использует OCR: PDF должен содержать текстовый слой.",
            "Для восстановления очевидно обрезанных названий используется словарь дисциплин из предыдущего расписания. Исходный фрагмент сохраняется в subject_raw.",
            "Если короткий фрагмент неоднозначен или не найден в словаре, запись помечается needs_review и не должна автоматически применяться без проверки.",
            "Все записи сохраняют source_page и raw_lines для ручной проверки.",
        ],
    }


def report_text(data: dict[str, Any], group_filter: str | None = None) -> str:
    lines = [
        f"Файл: {data['source_file']}",
        f"Тип: {data['document_type']}",
        f"Страниц: {data['page_count']}",
        f"Групп: {data['group_count']}",
        f"Записей занятий: {data['entry_count']}",
        f"Восстановлено названий: {data.get('quality_summary', {}).get('subject_recovered', 0)}",
        f"Требуют проверки: {data.get('quality_summary', {}).get('needs_review', 0)}",
        f"Группы: {', '.join(data['groups_found'])}",
        f"Даты: {', '.join(data['dates']) or 'не определены'}",
        "",
        *( [
            f"Эталон XLSX: {data['reference_xlsx']['source_file']}",
            f"Групп в XLSX: {data['reference_xlsx']['group_count']}; дисциплин: {data['reference_xlsx']['subject_count']}",
            f"Восстановлено по XLSX: {data['reference_xlsx']['recovered_from_reference']}; неоднозначных совпадений: {data['reference_xlsx']['ambiguous_reference_matches']}",
            f"Совпало с базовым слотом: {data['reference_xlsx']['same_as_reference_slot']}; отличается: {data['reference_xlsx']['differs_from_reference_slot']}; отсутствует в эталоне: {data['reference_xlsx']['not_in_reference_slot']}",
            "",
        ] if data.get('reference_xlsx') else [] ),
        "ПОЛЯ: пара | предмет | подгруппа | аудитория | преподаватель | отметки",
        "=" * 100,
    ]
    entries = data["entries"]
    if group_filter:
        entries = [e for e in entries if e["group"] == group_filter]
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for entry in entries:
        grouped[(entry["group"], entry.get("date") or "?")].append(entry)
    for (group, day), items in sorted(grouped.items(), key=lambda x: (x[0][0], x[0][1])):
        lines.append(f"\nГРУППА {group} · {day}")
        for e in sorted(items, key=lambda x: x["pair"]):
            flags = []
            if e.get("self_study"):
                flags.append("САМОСТОЯТЕЛЬНАЯ")
            if e.get("reference_recovered"):
                flags.append("НАЗВАНИЕ ВОССТАНОВЛЕНО ПО XLSX")
            elif e.get("subject_recovered"):
                flags.append("НАЗВАНИЕ ВОССТАНОВЛЕНО ПО СЛОВАРЮ")
            if e.get("reference_match") == "ambiguous":
                flags.append("НЕОДНОЗНАЧНЫЙ ЭТАЛОН XLSX")
            if e.get("reference_slot_status") == "differs_from_reference":
                flags.append("РАСХОЖДЕНИЕ С XLSX-ЭТАЛОНОМ")
            if e.get("needs_review"):
                flags.append("ТРЕБУЕТ ПРОВЕРКИ")
            if e.get("subgroup_marker_repaired"):
                flags.append("МЕТКА П/Г ВОССТАНОВЛЕНА")
            marker = ", ".join(flags) if flags else "ok"
            lines.append(
                f"  Пара {e['pair']}: {e.get('subject') or '[без предмета]'}"
                f" | п/г={e.get('subgroup') or '-'}"
                f" | ауд.={e.get('room') or '-'}"
                f" | преп.={e.get('teacher') or '-'}"
                f" | {marker}"
            )
    if not entries:
        lines.append("\nДля выбранной группы занятий не найдено.")
    lines += ["", "ПРИМЕЧАНИЯ"] + [f"- {x}" for x in data["limitations"]]
    return "\n".join(lines) + "\n"


def review_text(data: dict[str, Any]) -> str:
    """Create a compact audit file containing only unresolved parser records."""
    entries = [entry for entry in data.get("entries", []) if entry.get("needs_review")]
    lines = [
        f"Файл: {data['source_file']}",
        f"Тип: {data['document_type']}",
        f"Записей, требующих проверки: {len(entries)}",
        "",
    ]
    if not entries:
        lines.append("Непроверенных записей нет.")
        return "\n".join(lines) + "\n"
    for entry in sorted(entries, key=lambda item: (str(item.get("group")), str(item.get("date")), int(item.get("pair") or 0))):
        lines.append(
            f"Группа {entry.get('group')} · {entry.get('date')} · {entry.get('pair')} пара"
        )
        lines.append(f"  Исходный предмет: {entry.get('subject_raw') or '[пусто]'}")
        lines.append(f"  Полученный предмет: {entry.get('subject') or '[пусто]'}")
        if entry.get("subject_dictionary_candidates"):
            lines.append("  Возможные варианты: " + " / ".join(entry["subject_dictionary_candidates"]))
        if entry.get("reference_candidates"):
            lines.append("  Варианты из XLSX: " + " / ".join(entry["reference_candidates"]))
        if entry.get("reference_candidate_scope"):
            lines.append("  Область поиска XLSX: " + str(entry["reference_candidate_scope"]))
        if entry.get("reference_slot_entries"):
            baselines = [f"{x.get('subject')} | {x.get('room') or '-'} | {x.get('teacher') or '-'}" for x in entry["reference_slot_entries"]]
            lines.append("  Основное расписание: " + " ; ".join(baselines))
        lines.append(f"  Подгруппа: {entry.get('subgroup') or '-'} · аудитория: {entry.get('room') or '-'} · преподаватель: {entry.get('teacher') or '-'}")
        lines.append("  Причины: " + ", ".join(entry.get("quality_flags", [])))
        if entry.get("reference_slot_status"):
            lines.append("  Сопоставление с эталонным слотом: " + str(entry["reference_slot_status"]))
        lines.append(f"  Страница PDF: {entry.get('source_page')}")
        lines.append("")
    return "\n".join(lines) + "\n"


def main() -> int:
    ap = argparse.ArgumentParser(description="Разбор PDF-расписания СПТ по координатам таблицы")
    ap.add_argument("pdf", help="Путь к PDF")
    ap.add_argument("--group", help="Показать в TXT только выбранную группу")
    ap.add_argument("--dictionary", help="JSON со справочником полных названий дисциплин")
    ap.add_argument("--reference-xlsx", help="Недельный XLSX для сопоставления и восстановления обрезанных названий")
    ap.add_argument("--reference-dir", help="Папка с эталонными XLSX; выбирается файл, покрывающий даты PDF")
    ap.add_argument("-o", "--output", help="Путь к JSON; TXT создаётся рядом")
    args = ap.parse_args()
    path = Path(args.pdf).expanduser().resolve()
    try:
        default_dictionary = Path(__file__).with_name("subjects_dictionary.json")
        dictionary_path = Path(args.dictionary).expanduser().resolve() if args.dictionary else default_dictionary
        subject_dictionary = load_subject_dictionary(dictionary_path)
        data = parse_pdf(path, subject_dictionary)
        reference_path = Path(args.reference_xlsx).expanduser().resolve() if args.reference_xlsx else None
        if reference_path is None and args.reference_dir:
            reference_path = select_reference_xlsx(path, Path(args.reference_dir).expanduser().resolve())
        if reference_path is not None:
            reference_summary = apply_reference(data, reference_path)
            print(f"Reference XLSX: {reference_summary['source_file']} ({reference_summary['group_count']} groups; {reference_summary['record_count']} reference rows)")
            print(f"Reference recovered: {reference_summary['recovered_from_reference']}; ambiguous: {reference_summary['ambiguous_reference_matches']}")
        elif args.reference_dir:
            print("Reference XLSX: not found for this PDF date range")
        json_path = Path(args.output).expanduser().resolve() if args.output else path.with_suffix(".parsed.json")
        txt_path = json_path.with_suffix(".txt")
        json_path.parent.mkdir(parents=True, exist_ok=True)
        json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        txt_path.write_text(report_text(data, args.group), encoding="utf-8")
        review_path = json_path.with_name(json_path.stem + ".review.txt")
        review_path.write_text(review_text(data), encoding="utf-8")
        print(f"OK: {data['document_type']}")
        print(f"Pages: {data['page_count']}; groups: {data['group_count']}; entries: {data['entry_count']}")
        print(f"Dates: {', '.join(data['dates'])}")
        quality = data.get("quality_summary", {})
        print(f"Recovered subjects: {quality.get('subject_recovered', 0)}; needs review: {quality.get('needs_review', 0)}; subgroup markers repaired: {quality.get('subgroup_markers_repaired', 0)}")
        print(f"Dictionary entries: {len(subject_dictionary)} ({dictionary_path.name if dictionary_path.exists() else 'not found'})")
        if data.get("reference_xlsx"):
            ref = data["reference_xlsx"]
            print(f"Reference slots: same={ref['same_as_reference_slot']}; changed={ref['differs_from_reference_slot']}; no slot={ref['not_in_reference_slot']}")
        print(f"JSON: {json_path}")
        print(f"TXT: {txt_path}")
        print(f"Review TXT: {review_path}")
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
