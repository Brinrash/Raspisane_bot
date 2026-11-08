"""Institution bell timetable, pair-to-lesson mapping and dated overrides."""
from __future__ import annotations

import json
import re
import zipfile
from copy import deepcopy
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


def _lesson(start: str, end: str, pair: int) -> dict[str, Any]:
    return {"start": start, "end": end, "pair": pair}


# The institution's documents number individual lessons, not pairs.
# Two consecutive lessons form one pair: lessons 1+2 = pair 1, 3+4 = pair 2, etc.
DEFAULT_BELLS: dict[str, dict[str, Any]] = {
    "monday": {
        "label": "Понедельник",
        "class_hour": {"start": "08:15", "end": "09:00"},
        "lessons": {
            "1": _lesson("09:20", "10:05", 1),
            "2": _lesson("10:15", "11:00", 1),
            "3": _lesson("11:10", "11:55", 2),
            "4": _lesson("12:35", "13:20", 2),
            "5": _lesson("13:30", "14:15", 3),
            "6": _lesson("14:25", "15:10", 3),
            "7": _lesson("15:30", "16:15", 4),
            "8": _lesson("16:25", "17:10", 4),
        },
        "breaks": [
            {"after_lesson": 0, "label": "Завтрак", "duration": "20 мин."},
            {"after_lesson": 3, "label": "Обед", "duration": "40 мин."},
            {"after_lesson": 6, "label": "Полдник", "duration": "20 мин."},
        ],
        "practice_shifts": [
            {"label": "1 смена", "start": "09:10", "end": "14:00"},
            {"label": "2 смена", "start": "14:10", "end": "19:10"},
        ],
    },
    "weekday": {
        "label": "Со вторника по пятницу",
        "prep": {"start": "08:00", "end": "08:10", "label": "Десятиминутка для руководителей групп"},
        "lessons": {
            "1": _lesson("08:10", "08:55", 1),
            "2": _lesson("09:15", "10:00", 1),
            "3": _lesson("10:10", "10:55", 2),
            "4": _lesson("11:05", "11:50", 2),
            "5": _lesson("12:30", "13:15", 3),
            "6": _lesson("13:25", "14:10", 3),
            "7": _lesson("14:20", "15:05", 4),
            "8": _lesson("15:25", "16:10", 4),
            "9": _lesson("16:20", "17:05", 5),
            "10": _lesson("17:15", "18:00", 5),
            "11": _lesson("18:10", "18:55", 6),
            "12": _lesson("19:05", "19:50", 6),
        },
        "breaks": [
            {"after_lesson": 1, "label": "Завтрак", "duration": "20 мин."},
            {"after_lesson": 4, "label": "Обед", "duration": "40 мин."},
            {"after_lesson": 7, "label": "Полдник", "duration": "20 мин."},
        ],
        "practice_shifts": [
            {"label": "1 смена", "start": "08:00", "end": "13:00"},
            {"label": "2 смена", "start": "13:00", "end": "18:00"},
        ],
    },
}

TIME_RANGE_RE = re.compile(r"(?<!\d)(\d{1,2})[.:](\d{2})\s*[-–—]\s*(\d{1,2})[.:](\d{2})(?!\d)")
DATE_RE = re.compile(r"(?<!\d)(\d{2})[.](\d{2})[.](\d{4})(?!\d)")
MEAL_WORDS = ("завтрак", "обед", "полдник")


def _normalize_text(value: str) -> str:
    value = str(value or "")
    value = re.sub(r"(?<=\d)\s+(?=[\d.:])", "", value)
    value = re.sub(r"(?<=[.:])\s+(?=\d)", "", value)
    value = re.sub(r"\s*[-–—]\s*", "-", value)
    return value


def _range_from_text(text: str) -> dict[str, str] | None:
    match = TIME_RANGE_RE.search(_normalize_text(text))
    if not match:
        return None
    return {
        "start": f"{int(match.group(1)):02d}:{match.group(2)}",
        "end": f"{int(match.group(3)):02d}:{match.group(4)}",
    }


def _read_docx_tables(path: Path) -> tuple[list[list[list[str]]], str]:
    """Read table cell text directly from DOCX XML (no python-docx dependency)."""
    with zipfile.ZipFile(path) as archive:
        xml = ET.fromstring(archive.read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    body = xml.find("w:body", ns)
    tables: list[list[list[str]]] = []
    full_text: list[str] = []
    if body is None:
        return tables, ""
    for child in body:
        if child.tag.endswith("}tbl"):
            table: list[list[str]] = []
            for row in child.findall("w:tr", ns):
                cells = [" ".join(t.text or "" for t in cell.findall(".//w:t", ns)).strip()
                         for cell in row.findall("w:tc", ns)]
                table.append(cells)
                full_text.extend(cells)
            tables.append(table)
        elif child.tag.endswith("}p"):
            full_text.extend(t.text or "" for t in child.findall(".//w:t", ns))
    return tables, " ".join(full_text)


def _date_from_text(text: str) -> date | None:
    match = DATE_RE.search(text)
    if not match:
        return None
    try:
        return date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
    except ValueError:
        return None


def _parse_docx(path: Path) -> tuple[dict[str, Any], str]:
    tables, full_text = _read_docx_tables(path)
    result: dict[str, Any] = {"lessons": {}, "breaks": [], "practice_shifts": []}
    lesson_table = tables[0] if tables else []
    highest_lesson_seen = 0

    for row in lesson_table:
        joined = " ".join(row).strip()
        lower = joined.casefold()
        # Parse the time from its own cell. Joining a leading lesson number and
        # the time cell can collapse e.g. "2  10.15" into "210.15".
        time_range = next((value for cell in row if (value := _range_from_text(cell))), None)

        if any(word in lower for word in MEAL_WORDS):
            meal = next(word for word in MEAL_WORDS if word in lower)
            duration = re.search(r"(\d+)\s*мин", _normalize_text(joined), re.IGNORECASE)
            label = {"завтрак": "Завтрак", "обед": "Обед", "полдник": "Полдник"}[meal]
            result["breaks"].append({
                "after_lesson": highest_lesson_seen,
                "label": label,
                "duration": f"{duration.group(1)} мин." if duration else "",
            })
            continue

        if "классный час" in lower and time_range:
            result["class_hour"] = time_range
            continue
        if "десятиминутка" in lower and time_range:
            result["prep"] = {**time_range, "label": "Десятиминутка для руководителей групп"}
            continue

        index_text = _normalize_text(row[0] if row else "")
        # Some short-day documents combine two lesson numbers in one row
        # (e.g. "1,2" / "3,4") and give one shared time interval for both.
        lesson_numbers = [int(value) for value in re.findall(r"(?<!\d)\d{1,2}(?!\d)", index_text)]
        if lesson_numbers and time_range:
            for lesson_number in lesson_numbers:
                if 1 <= lesson_number <= 12:
                    result["lessons"][str(lesson_number)] = time_range
                    highest_lesson_seen = max(highest_lesson_seen, lesson_number)

    # Explicitly shortened schedules replace the normal day's entire bell
    # timetable, not just the lesson slots listed in the document.
    if re.search(r"сокращ[её]нн", full_text, re.IGNORECASE):
        result["complete_timetable"] = True

    # Practice shifts are in a separate table in the source documents.
    for table in tables[1:]:
        for row in table:
            joined = " ".join(row).strip()
            match = re.search(r"([12])\s*смена", joined, re.IGNORECASE)
            time_range = _range_from_text(joined)
            if match and time_range:
                result["practice_shifts"].append({
                    "label": f"{match.group(1)} смена",
                    **time_range,
                })

    if not result["lessons"]:
        raise ValueError("В DOCX не найдена таблица времени уроков с номерами.")
    return result, full_text


def _parse_pdf(path: Path) -> tuple[dict[str, Any], str]:
    try:
        from pypdf import PdfReader  # type: ignore
    except ImportError as exc:
        raise RuntimeError("Для разбора PDF установи зависимость pypdf из requirements.txt") from exc
    text = "\n".join(page.extract_text() or "" for page in PdfReader(str(path)).pages)
    # Normalize each line separately. Normalizing the entire PDF text can
    # accidentally join the last digit of one row to the first digit of the
    # next row (e.g. "09.10\n3,4" -> "09.103,4").
    flat = "\n".join(_normalize_text(line) for line in text.splitlines())
    result: dict[str, Any] = {"lessons": {}, "breaks": [], "practice_shifts": []}

    # Rows may identify one lesson ("1 I ...") or a merged pair
    # ("1,2 I ...") with a single shortened interval for both lessons.
    number_group = r"(\d{1,2}(?:[ \t]*[,;][ \t]*\d{1,2})?)"
    row_re = re.compile(
        rf"(?<!\d){number_group}[ \t]+(?:I{{1,3}}|IV|V|VI)[ \t]+"
        rf"(\d{{1,2}}[.:]\d{{2}})[ \t]*-[ \t]*(\d{{1,2}}[.:]\d{{2}})(?!\d)",
        re.IGNORECASE,
    )
    row_matches = list(row_re.finditer(flat))
    for match in row_matches:
        numbers = [int(value) for value in re.findall(r"\d{1,2}", match.group(1))]
        value = _range_from_text(f"{match.group(2)}-{match.group(3)}")
        if value:
            for n in numbers:
                if 1 <= n <= 12:
                    result["lessons"][str(n)] = value

    for label in MEAL_WORDS:
        for match in re.finditer(rf"{label}[^\n]*?(\d+)\s*мин", flat, re.IGNORECASE):
            # Estimate placement from the lesson rows appearing before this break.
            seen_numbers: list[int] = []
            for row_match in row_matches:
                if row_match.start() >= match.start():
                    break
                seen_numbers.extend(int(value) for value in re.findall(r"\d{1,2}", row_match.group(1)))
            result["breaks"].append({
                "after_lesson": max(seen_numbers, default=0),
                "label": label.capitalize(),
                "duration": f"{match.group(1)} мин.",
            })

    class_match = re.search(r"Классный\s+час[^\n]*?(\d{1,2}[.:]\d{2}\s*-\s*\d{1,2}[.:]\d{2})", flat, re.IGNORECASE)
    if class_match:
        result["class_hour"] = _range_from_text(class_match.group(1))
    prep_match = re.search(r"Десятиминутка[^\n]*?(\d{1,2}[.:]\d{2}\s*-\s*\d{1,2}[.:]\d{2})", flat, re.IGNORECASE)
    if prep_match:
        result["prep"] = {**(_range_from_text(prep_match.group(1)) or {}), "label": "Десятиминутка для руководителей групп"}

    for shift in (1, 2):
        shift_match = re.search(rf"{shift}\s*смена[^\n]*?(\d{{1,2}}[.:]\d{{2}}\s*-\s*\d{{1,2}}[.:]\d{{2}})", flat, re.IGNORECASE)
        if shift_match:
            value = _range_from_text(shift_match.group(1))
            if value:
                result["practice_shifts"].append({"label": f"{shift} смена", **value})

    if re.search(r"сокращ[её]нн", text, re.IGNORECASE):
        result["complete_timetable"] = True
    if not result["lessons"]:
        raise ValueError("В PDF не удалось извлечь уроки. Возможно, это скан без текстового слоя.")
    return result, text


def parse_bell_document_dates(path: str | Path) -> tuple[list[date], dict[str, Any]]:
    """Parse a DOCX/PDF bell timetable and all dates to which it applies.

    If the document names a date range ("с 03.11.2026 по 05.11.2026"),
    the same override is applied to every date in that inclusive range.
    Otherwise all explicit dates are used (duplicates are removed).
    """
    file_path = Path(path)
    if file_path.suffix.lower() == ".docx":
        parsed, text = _parse_docx(file_path)
    elif file_path.suffix.lower() == ".pdf":
        parsed, text = _parse_pdf(file_path)
    else:
        raise ValueError("Распорядок звонков должен быть в DOCX или PDF.")

    combined = f"{file_path.name} {text}"
    date_matches = list(DATE_RE.finditer(combined))
    dates_found: list[date] = []
    for match in date_matches:
        try:
            value = date(int(match.group(3)), int(match.group(2)), int(match.group(1)))
        except ValueError:
            continue
        if value not in dates_found:
            dates_found.append(value)

    range_match = re.search(
        r"с\s*(\d{2}\.\d{2}\.\d{4})\s*по\s*(\d{2}\.\d{2}\.\d{4})",
        combined,
        re.IGNORECASE,
    )
    if range_match:
        start_day = _date_from_text(range_match.group(1))
        end_day = _date_from_text(range_match.group(2))
        if start_day and end_day and end_day >= start_day and (end_day - start_day).days <= 31:
            dates_found = [start_day + timedelta(days=offset) for offset in range((end_day - start_day).days + 1)]

    if not dates_found:
        raise ValueError("Не найдена дата изменения в имени файла или тексте (например, 09.11.2026).")
    if not 1 <= len(parsed.get("lessons", {})) <= 12:
        raise ValueError("В документе найдено необычное количество уроков; изменения не сохранены.")
    return dates_found, parsed


def parse_bell_document(path: str | Path) -> tuple[date, dict[str, Any]]:
    """Backward-compatible single-date parser; use parse_bell_document_dates for new code."""
    dates, parsed = parse_bell_document_dates(path)
    return dates[0], parsed


def load_bell_data(path: str | Path) -> dict[str, Any]:
    file_path = Path(path)
    if not file_path.exists():
        return {"overrides": {}, "updated_at": None}
    try:
        data = json.loads(file_path.read_text(encoding="utf-8"))
        if isinstance(data, dict) and isinstance(data.get("overrides"), dict):
            return data
    except (OSError, json.JSONDecodeError):
        pass
    return {"overrides": {}, "updated_at": None}


def save_bell_data(data: dict[str, Any], path: str | Path) -> None:
    file_path = Path(path)
    file_path.parent.mkdir(parents=True, exist_ok=True)
    file_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def merge_bell_override(current: dict[str, Any], day: date, incoming: dict[str, Any], source_name: str) -> dict[str, Any]:
    result = deepcopy(current or {"overrides": {}})
    result.setdefault("overrides", {})
    result["overrides"][day.isoformat()] = {
        **deepcopy(incoming),
        "source_file": source_name,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    result["updated_at"] = datetime.now().isoformat(timespec="seconds")
    return result


def retention_bounds(today: date) -> tuple[date, date]:
    """Keep previous, current and next Monday-Sunday calendar weeks inclusive."""
    current_monday = today - timedelta(days=today.weekday())
    return current_monday - timedelta(days=7), current_monday + timedelta(days=13)


def prune_bell_data(data: dict[str, Any], today: date) -> dict[str, Any]:
    start, end = retention_bounds(today)
    result = deepcopy(data or {"overrides": {}})
    result.setdefault("overrides", {})
    result["overrides"] = {
        key: value for key, value in result["overrides"].items()
        if _is_date_in_range(key, start, end)
    }
    return result


def _is_date_in_range(value: str, start: date, end: date) -> bool:
    try:
        return start <= date.fromisoformat(value) <= end
    except (TypeError, ValueError):
        return False


def bell_for_date(day: date, data: dict[str, Any] | None = None) -> dict[str, Any]:
    result = deepcopy(DEFAULT_BELLS["monday" if day.weekday() == 0 else "weekday"])
    override = (data or {}).get("overrides", {}).get(day.isoformat())
    if not override:
        return result
    if override.get("complete_timetable"):
        # A shortened/full replacement timetable is authoritative for this date.
        result["lessons"] = deepcopy(override.get("lessons") or {})
        for key in ("class_hour", "prep", "breaks", "practice_shifts"):
            if key in override:
                result[key] = deepcopy(override[key])
            elif key in {"class_hour", "prep"}:
                result.pop(key, None)
            else:
                result[key] = []
    else:
        result.setdefault("lessons", {}).update(override.get("lessons") or {})
        for key in ("class_hour", "prep", "breaks", "practice_shifts"):
            if override.get(key):
                result[key] = deepcopy(override[key])
    result["complete_timetable"] = bool(override.get("complete_timetable"))
    result["source_file"] = override.get("source_file")
    return result


def pair_time(day: date, pair: int, data: dict[str, Any] | None = None) -> str | None:
    """Return both individual lesson intervals belonging to the requested pair."""
    bells = bell_for_date(day, data)
    if pair == 0:
        value = bells.get("class_hour")
        return f"{value['start']}–{value['end']}" if value else None
    lesson_numbers = (pair * 2 - 1, pair * 2)
    intervals: list[str] = []
    lessons = bells.get("lessons", {})
    for number in lesson_numbers:
        value = lessons.get(str(number))
        if value and value.get("start") and value.get("end"):
            label = f"{value['start']}–{value['end']}"
            if label not in intervals:
                intervals.append(label)
    return " / ".join(intervals) if intervals else None


def format_bells(day: date, data: dict[str, Any] | None = None) -> str:
    bells = bell_for_date(day, data)
    weekday_name = "понедельник" if day.weekday() == 0 else "вторник–пятница"
    lines = [f"🔔 РАСПИСАНИЕ ЗВОНКОВ · {day:%d.%m.%Y}", f"Распорядок: {weekday_name}"]
    if bells.get("complete_timetable"):
        lines.append("⚠️ Действует специальное сокращённое расписание")
    lines.append("━━━━━━━━━━━━━━━━━━")

    class_hour = bells.get("class_hour")
    if class_hour:
        lines.append(f"🧑‍🏫 Классный час · {class_hour['start']}–{class_hour['end']}")
    prep = bells.get("prep")
    if prep:
        lines.append(f"🕗 {prep.get('label', 'Десятиминутка')} · {prep['start']}–{prep['end']}")

    breaks_by_position: dict[int, list[dict[str, Any]]] = {}
    for item in bells.get("breaks", []):
        breaks_by_position.setdefault(int(item.get("after_lesson", 0)), []).append(item)

    lessons = bells.get("lessons", {})
    for item in breaks_by_position.get(0, []):
        lines.append(f"🍽 {item.get('label', 'Перерыв')} · {item.get('duration') or 'перерыв'}")

    # A shortened-day document can give one shared time interval for a row
    # containing lesson numbers "1,2". Display that interval once as
    # "1–2-й урок" instead of duplicating the same time for both lessons.
    number = 1
    while number <= 12:
        value = lessons.get(str(number))
        if value and value.get("start") and value.get("end"):
            last_number = number
            while last_number < 12:
                following = lessons.get(str(last_number + 1))
                if not following or (following.get("start"), following.get("end")) != (value.get("start"), value.get("end")):
                    break
                last_number += 1
            label = f"{number}–{last_number}-й урок" if last_number > number else f"{number}-й урок"
            lines.append(f"{label} — {value['start']}–{value['end']}")
            number = last_number
        for item in breaks_by_position.get(number, []):
            lines.append(f"🍽 {item.get('label', 'Перерыв')} · {item.get('duration') or 'перерыв'}")
        number += 1

    practice = bells.get("practice_shifts", [])
    if practice:
        lines.extend(["", "📚 УЧЕБНАЯ ПРАКТИКА"])
        for shift in practice:
            lines.append(f"{shift['label']}: {shift['start']}–{shift['end']}")

    if day.isoformat() in (data or {}).get("overrides", {}):
        lines.extend(["", "⚠️ На эту дату действует опубликованное изменение звонков."])
    return "\n".join(lines)
