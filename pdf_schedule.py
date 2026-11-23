"""Convert coordinate-parsed PDF entries into the bot's schedule JSON schema."""
from __future__ import annotations

from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any

from parser import parse_xlsx


def pdf_to_schedule(
    parsed: dict[str, Any],
    *,
    reference_path: str | Path | None = None,
) -> tuple[dict[str, Any], list[dict[str, str]]]:
    """Convert safe PDF group/date blocks into the bot's schedule JSON schema.

    For a weekly PDF only, an XLSX with the exact same date range may fill a
    missing or uncertain group/date block. This is safe because the file is an
    explicitly selected same-week reference. Daily replacement PDFs are never
    filled from the baseline XLSX: a replacement must remain authoritative.
    """
    entries = parsed.get("entries") or []
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for item in entries:
        group = str(item.get("group") or "").strip()
        day = str(item.get("date") or "").strip()
        if not group or not day:
            continue
        grouped.setdefault((group, day), []).append(item)

    groups: dict[str, dict[str, Any]] = {}
    skipped: list[dict[str, str]] = []
    doc_type = str(parsed.get("document_type") or "")

    # Weekly PDF pages explicitly describe every group/date cell, including
    # empty days. Preserve those empty blocks so an old cached Thursday/Friday
    # is cleared instead of silently surviving when the new PDF has no lessons.
    weekly_blocks: dict[tuple[str, str], list[dict[str, Any]]] = {}
    if doc_type == "weekly":
        for page in parsed.get("pages", []):
            for page_group in page.get("groups", []):
                group_code = str(page_group.get("group") or "").strip()
                if not group_code:
                    continue
                for day_block in page_group.get("days", []):
                    day_value = str(day_block.get("date") or "").strip()
                    if day_value:
                        weekly_blocks[(group_code, day_value)] = list(day_block.get("entries") or [])

    # Use page-level weekly blocks where available. For older parser payloads
    # and daily replacement PDFs, the flat entries list remains authoritative.
    source_blocks = dict(weekly_blocks)
    for key, rows in grouped.items():
        source_blocks.setdefault(key, rows)

    for (group, day), rows in source_blocks.items():
        uncertain = [
            row for row in rows
            if row.get("needs_review") or not str(row.get("subject") or "").strip()
        ]
        if uncertain:
            skipped.append({
                "group": group,
                "date": day,
                "reason": f"сомнительных/пустых записей: {len(uncertain)}",
            })
            continue

        lessons = _pdf_rows_to_lessons(rows)
        target = groups.setdefault(group, {"sheet": "PDF", "days": {}})
        # An explicitly empty weekly block is meaningful: no lessons that day.
        target["days"][day] = lessons

    recovered_reference_blocks: list[dict[str, str]] = []
    dates = sorted(str(day) for day in (parsed.get("dates") or []))

    # A same-range XLSX can safely complete the weekly timetable where PDF text
    # extraction lost a whole column/day or where a block was deliberately held.
    if doc_type == "weekly" and reference_path and dates:
        reference = parse_xlsx(reference_path)
        if (
            str(reference.get("date_from") or "") == str(parsed.get("date_from") or dates[0])
            and str(reference.get("date_to") or "") == str(parsed.get("date_to") or dates[-1])
        ):
            blocked_keys = {(item["group"], item["date"]) for item in skipped}
            # A present-but-empty PDF block is not a missing block and must not
            # be repopulated from the XLSX baseline. Fallback is only for truly
            # absent or rejected blocks.
            existing_keys = set(source_blocks)
            for group_code, group_data in reference.get("groups", {}).items():
                for day in dates:
                    ref_lessons = group_data.get("days", {}).get(day)
                    key = (str(group_code), day)
                    # Only fill a genuinely absent PDF block or a block already
                    # rejected as uncertain; never overwrite a parsed PDF block.
                    if not ref_lessons or key in existing_keys and key not in blocked_keys:
                        continue
                    target = groups.setdefault(str(group_code), {"sheet": "PDF + XLSX reference", "days": {}})
                    target["days"][day] = _xlsx_rows_to_lessons(ref_lessons)
                    recovered_reference_blocks.append({"group": str(group_code), "date": day})
                    skipped = [item for item in skipped if (item["group"], item["date"]) != key]

    dates_in_output = sorted({
        day
        for group_data in groups.values()
        for day in group_data.get("days", {})
    })
    if not dates_in_output:
        raise ValueError(
            "В PDF не осталось безопасных для импорта блоков. "
            f"Блоков с сомнительными строками: {len(skipped)}. Проверьте XLSX-эталон."
        )

    mode = parsed.get("replacement_mode") if doc_type == "replacements" else None
    return {
        "source_file": parsed.get("source_file") or "schedule.pdf",
        "parsed_at": datetime.now().isoformat(timespec="seconds"),
        "date_from": dates_in_output[0],
        "date_to": dates_in_output[-1],
        "dates": dates_in_output,
        "groups": groups,
        "group_count": len(groups),
        "table_count": parsed.get("page_count", 0),
        "replacement_mode": mode or "replace",
        "format": "pdf_weekly" if doc_type == "weekly" else "pdf_replacement",
        "reference_fallback_blocks": recovered_reference_blocks,
    }, skipped


def _pdf_rows_to_lessons(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    lessons: list[dict[str, Any]] = []
    seen: set[tuple[Any, ...]] = set()
    for row in rows:
        try:
            pair = int(row.get("pair") or 0)
        except (TypeError, ValueError):
            continue
        if not 0 <= pair <= 12:
            continue
        lesson_numbers = row.get("lesson_numbers") or []
        lesson_number = None
        if lesson_numbers:
            try:
                lesson_number = int(lesson_numbers[0])
            except (TypeError, ValueError):
                lesson_number = None
        lesson = {
            "pair": pair,
            "subject": str(row.get("subject") or "").strip(),
            "room": str(row.get("room") or "").strip() or None,
            "teacher": str(row.get("teacher") or "").strip() or None,
            "subgroup": row.get("subgroup"),
            "self_study": bool(row.get("self_study")),
            "source_page": row.get("source_page"),
            "raw_lines": row.get("raw_lines") or [],
            "quality_status": row.get("quality_status") or "ok",
            "quality_flags": row.get("quality_flags") or [],
        }
        if lesson_number is not None:
            lesson["lesson_number"] = lesson_number
        slot_key = (pair, lesson.get("subgroup"), lesson["subject"], lesson.get("teacher"), lesson.get("room"))
        if slot_key in seen:
            continue
        seen.add(slot_key)
        lessons.append(lesson)
    lessons.sort(key=lambda item: (
        int(item.get("pair") or 0),
        int(item.get("lesson_number") or 0),
        int(item.get("subgroup") or 0),
    ))
    return lessons


def _xlsx_rows_to_lessons(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    lessons: list[dict[str, Any]] = []
    for row in rows:
        lesson = deepcopy(row)
        lesson["self_study"] = bool(lesson.get("self_study"))
        lesson["source_page"] = None
        lesson["raw_lines"] = ["Восстановлено из XLSX-эталона того же диапазона дат"]
        lesson["quality_status"] = "same_week_xlsx_fallback"
        lesson["quality_flags"] = ["weekly_block_filled_from_matching_xlsx"]
        lessons.append(lesson)
    lessons.sort(key=lambda item: (int(item.get("pair") or 0), int(item.get("subgroup") or 0)))
    return lessons


def reference_xlsx_to_schedule(path: str | Path) -> dict[str, Any]:
    """Parse a weekly XLSX file into the bot's native schedule schema."""
    parsed = parse_xlsx(path)
    parsed["replacement_mode"] = "replace"
    parsed["format"] = "xlsx_reference"
    return parsed
