"""Persistent schedule loading, three-week retention and safe merging of updates."""
from __future__ import annotations

import json
from copy import deepcopy
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from parser import parse_replacement_xlsx, parse_xlsx, save_json


def retention_bounds(today: date) -> tuple[date, date]:
    """Inclusive range: previous, current and next Monday-Sunday week."""
    current_monday = today - timedelta(days=today.weekday())
    return current_monday - timedelta(days=7), current_monday + timedelta(days=13)


def prune_schedule_data(data: dict[str, Any], today: date) -> tuple[dict[str, Any], dict[str, Any]]:
    """Drop only lesson dates outside the rolling three-calendar-week window."""
    start, end = retention_bounds(today)
    result = deepcopy(data) if data else {"groups": {}}
    result.setdefault("groups", {})
    removed_dates: set[str] = set()
    for group_code in list(result["groups"]):
        days = result["groups"][group_code].get("days", {})
        for date_text in list(days):
            try:
                parsed_date = date.fromisoformat(date_text)
            except (TypeError, ValueError):
                del days[date_text]
                continue
            if not start <= parsed_date <= end:
                removed_dates.add(date_text)
                del days[date_text]
        if not days:
            del result["groups"][group_code]
    dates = sorted({
        date_text
        for group in result["groups"].values()
        for date_text in group.get("days", {})
    })
    result["dates"] = dates
    result["date_from"] = dates[0] if dates else None
    result["date_to"] = dates[-1] if dates else None
    result["group_count"] = len(result["groups"])
    result["retention_start"] = start.isoformat()
    result["retention_end"] = end.isoformat()
    return result, {"removed_dates": sorted(removed_dates), "retention_start": start.isoformat(), "retention_end": end.isoformat()}


def prune_incoming_files(directory: str | Path, today: date, max_files: int = 30) -> list[str]:
    """Keep recent source files only; old imports should not fill the small host disk."""
    folder = Path(directory)
    if not folder.exists():
        return []
    start, _ = retention_bounds(today)
    cutoff = datetime.combine(start, datetime.min.time()).timestamp()
    files = sorted((p for p in folder.iterdir() if p.is_file()), key=lambda p: p.stat().st_mtime, reverse=True)
    removed: list[str] = []
    for index, file_path in enumerate(files):
        try:
            too_old = file_path.stat().st_mtime < cutoff
        except OSError:
            continue
        if too_old or index >= max_files:
            try:
                file_path.unlink(missing_ok=True)
                removed.append(file_path.name)
            except OSError:
                pass
    return removed


def load_schedule(
    xlsx_path: str | Path,
    json_path: str | Path,
    *,
    today: date | None = None,
) -> dict[str, Any]:
    """Load cached schedule, prune stale dates, or parse the bundled XLSX once."""
    today = today or date.today()
    json_file = Path(json_path)
    if json_file.exists():
        try:
            data = json.loads(json_file.read_text(encoding="utf-8"))
            if isinstance(data, dict) and isinstance(data.get("groups"), dict):
                data, prune_info = prune_schedule_data(data, today)
                if prune_info["removed_dates"]:
                    data["retention_pruned_at"] = datetime.now().isoformat(timespec="seconds")
                data.setdefault("update_history", [])
                if data.get("groups"):
                    save_json(data, json_file)
                    return data
        except (OSError, json.JSONDecodeError):
            pass

    data = parse_xlsx(xlsx_path)
    data.setdefault("update_history", [])
    data["schedule_updated_at"] = data.get("parsed_at")
    data, _ = prune_schedule_data(data, today)
    save_json(data, json_file)
    return data


def _is_valid_iso_date(value: str) -> bool:
    try:
        date.fromisoformat(value)
        return True
    except (TypeError, ValueError):
        return False


def merge_schedule_data(
    current: dict[str, Any],
    incoming: dict[str, Any],
    *,
    today: date | None = None,
    source_name: str | None = None,
    source: str = "upload",
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Incoming data wins per (group,date), then data is trimmed to three weeks."""
    today = today or date.today()
    start, end = retention_bounds(today)
    merged = deepcopy(current) if current else {"groups": {}}
    merged.setdefault("groups", {})
    changed_groups: set[str] = set()
    changed_dates: set[str] = set()
    accepted_incoming_dates: set[str] = set()

    for group_code, incoming_group in incoming.get("groups", {}).items():
        target = merged["groups"].setdefault(group_code, {"sheet": incoming_group.get("sheet"), "days": {}})
        target.setdefault("days", {})
        if incoming_group.get("sheet"):
            target["sheet"] = incoming_group["sheet"]
        for date_text, lessons in incoming_group.get("days", {}).items():
            try:
                incoming_day = date.fromisoformat(date_text)
            except (TypeError, ValueError):
                continue
            if not start <= incoming_day <= end:
                continue
            accepted_incoming_dates.add(date_text)
            old_lessons = target["days"].get(date_text)
            new_lessons = deepcopy(lessons)
            if incoming.get("replacement_mode") == "supplement":
                # "Дополнительно" files supplement a day's timetable. Where
                # both files contain the same individual lesson slot and
                # subgroup, the newer row wins; all other lessons are kept.
                combined = deepcopy(old_lessons or [])
                for incoming_lesson in new_lessons:
                    slot_key = (
                        incoming_lesson.get("pair"),
                        incoming_lesson.get("lesson_number"),
                        incoming_lesson.get("subgroup"),
                    )
                    matched_index = next((
                        index for index, existing_lesson in enumerate(combined)
                        if existing_lesson.get("lesson_number") is not None
                        and (
                            existing_lesson.get("pair"),
                            existing_lesson.get("lesson_number"),
                            existing_lesson.get("subgroup"),
                        ) == slot_key
                    ), None)
                    if matched_index is not None:
                        combined[matched_index] = incoming_lesson
                    elif incoming_lesson not in combined:
                        combined.append(incoming_lesson)
                combined.sort(key=lambda item: (
                    int(item.get("pair") or 0),
                    int(item.get("lesson_number") or 0),
                    int(item.get("subgroup") or 0),
                ))
                final_lessons = combined
            else:
                final_lessons = new_lessons

            if old_lessons != final_lessons:
                changed_groups.add(group_code)
                changed_dates.add(date_text)
            target["days"][date_text] = final_lessons

    if incoming.get("dates") and not accepted_incoming_dates:
        has_current_window_data = any(
            start <= date.fromisoformat(date_text) <= end
            for group in current.get("groups", {}).values()
            for date_text in group.get("days", {})
            if _is_valid_iso_date(date_text)
        )
        if not has_current_window_data:
            raise ValueError(
                "В файле нет дат в допустимом окне хранения: предыдущая, текущая и следующая неделя "
                f"({start.isoformat()} — {end.isoformat()})."
            )

    merged, prune_info = prune_schedule_data(merged, today)
    if not merged.get("dates"):
        raise ValueError(
            "В файле нет дат в допустимом окне хранения: предыдущая, текущая и следующая неделя "
            f"({prune_info['retention_start']} — {prune_info['retention_end']})."
        )

    merged["table_count"] = max(int(merged.get("table_count", 0)), int(incoming.get("table_count", 0)))
    merged["source_file"] = source_name or incoming.get("source_file") or "schedule.xlsx"
    merged["schedule_updated_at"] = datetime.now().isoformat(timespec="seconds")
    history = list(merged.get("update_history", []))
    history.append({
        "at": merged["schedule_updated_at"],
        "source": source,
        "file": merged["source_file"],
        "date_from": incoming.get("date_from"),
        "date_to": incoming.get("date_to"),
        "changed_groups": len(changed_groups),
        "changed_dates": sorted(changed_dates),
    })
    merged["update_history"] = history[-12:]

    diff = {
        "changed_groups": sorted(changed_groups, key=lambda x: (len(x), x)),
        "changed_dates": sorted(changed_dates),
        "incoming_groups": len(incoming.get("groups", {})),
        "incoming_dates": sorted(incoming.get("dates", [])),
        "replacement_mode": incoming.get("replacement_mode", "replace"),
        "removed_dates": prune_info["removed_dates"],
        "retention_start": prune_info["retention_start"],
        "retention_end": prune_info["retention_end"],
    }
    return merged, diff


def import_schedule_file(
    file_path: str | Path,
    current: dict[str, Any],
    json_path: str | Path,
    *,
    today: date | None = None,
    source: str = "upload",
    source_name: str | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    path = Path(file_path)
    # Daily replacement workbooks have a different layout from the weekly file.
    # Prefer their dedicated parser when the file's name/title identifies them;
    # this also avoids the misleading weekly-parser traceback seen in older builds.
    is_replacement_name = any(
        marker in path.name.casefold()
        for marker in ("замена", "дополнительно")
    )
    if is_replacement_name:
        try:
            parsed = parse_replacement_xlsx(path)
        except ValueError as replacement_error:
            try:
                parsed = parse_xlsx(path)
            except ValueError as weekly_error:
                raise ValueError(
                    "Не удалось распознать XLSX ни как таблицу замен, ни как недельное расписание. "
                    f"Формат замен: {replacement_error}; недельный формат: {weekly_error}"
                ) from replacement_error
    else:
        try:
            parsed = parse_xlsx(path)
        except ValueError as weekly_error:
            try:
                parsed = parse_replacement_xlsx(path)
            except ValueError as replacement_error:
                raise ValueError(
                    "Не удалось распознать XLSX ни как недельное расписание, ни как таблицу замен. "
                    f"Недельный формат: {weekly_error}; формат замен: {replacement_error}"
                ) from weekly_error
    if not parsed.get("groups"):
        raise ValueError("В файле не найдено ни одной группы.")
    merged, diff = merge_schedule_data(
        current,
        parsed,
        today=today,
        source_name=source_name or path.name,
        source=source,
    )
    save_json(merged, json_path)
    return merged, diff
