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

        index_text = row[0] if row else ""
        number_match = re.fullmatch(r"\s*(\d{1,2})\s*", index_text)
        if number_match and time_range:
            lesson_number = int(number_match.group(1))
            if 1 <= lesson_number <= 12:
                result["lessons"][str(lesson_number)] = time_range
                highest_lesson_seen = max(highest_lesson_seen, lesson_number)

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
    flat = _normalize_text(text)
    result: dict[str, Any] = {"lessons": {}, "breaks": [], "practice_shifts": []}

    # Lesson rows normally appear as: lesson number, Roman pair number, time range.
    row_re = re.compile(r"(?<!\d)(\d{1,2})\s+(?:I{1,3}|IV|V|VI)\s+(\d{1,2}[.:]\d{2})\s*-\s*(\d{1,2}[.:]\d{2})(?!\d)", re.IGNORECASE)
    for match in row_re.finditer(flat):
        n = int(match.group(1))
        if 1 <= n <= 12:
            value = _range_from_text(f"{match.group(2)}-{match.group(3)}")
            if value:
                result["lessons"][str(n)] = value

    for label in MEAL_WORDS:
        for match in re.finditer(rf"{label}\s*(\d+)\s*мин", flat, re.IGNORECASE):
            # Estimate placement from the number of lesson rows preceding the meal.
            seen = [int(m.group(1)) for m in row_re.finditer(flat[:match.start()])]
            result["breaks"].append({
                "after_lesson": max(seen, default=0),
                "label": label.capitalize(),
                "duration": f"{match.group(1)} мин.",
            })

    class_match = re.search(r"Классный\s+час\s*(\d{1,2}[.:]\d{2}\s*-\s*\d{1,2}[.:]\d{2})", flat, re.IGNORECASE)
    if class_match:
        result["class_hour"] = _range_from_text(class_match.group(1))
    prep_match = re.search(r"Десятиминутка.*?(\d{1,2}[.:]\d{2}\s*-\s*\d{1,2}[.:]\d{2})", flat, re.IGNORECASE)
    if prep_match:
        result["prep"] = {**(_range_from_text(prep_match.group(1)) or {}), "label": "Десятиминутка для руководителей групп"}

    for shift in (1, 2):
        shift_match = re.search(rf"{shift}\s*смена\s*(\d{{1,2}}[.:]\d{{2}}\s*-\s*\d{{1,2}}[.:]\d{{2}})", flat, re.IGNORECASE)
        if shift_match:
            value = _range_from_text(shift_match.group(1))
            if value:
                result["practice_shifts"].append({"label": f"{shift} смена", **value})

    if not result["lessons"]:
        raise ValueError("В PDF не удалось извлечь уроки. Возможно, это скан без текстового слоя.")
    return result, text


def parse_bell_document(path: str | Path) -> tuple[date, dict[str, Any]]:
    file_path = Path(path)
    if file_path.suffix.lower() == ".docx":
        parsed, text = _parse_docx(file_path)
    elif file_path.suffix.lower() == ".pdf":
        parsed, text = _parse_pdf(file_path)
    else:
        raise ValueError("Распорядок звонков должен быть в DOCX или PDF.")
    target_day = _date_from_text(file_path.name) or _date_from_text(text)
    if not target_day:
        raise ValueError("Не найдена дата изменения в имени файла или тексте (например, 09.11.2026).")
    if not 1 <= len(parsed.get("lessons", {})) <= 12:
        raise ValueError("В документе найдено необычное количество уроков; изменения не сохранены.")
    return target_day, parsed


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
    result.setdefault("lessons", {}).update(override.get("lessons") or {})
    for key in ("class_hour", "prep", "breaks", "practice_shifts"):
        if override.get(key):
            result[key] = deepcopy(override[key])
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
            intervals.append(f"{value['start']}–{value['end']}")
    return " / ".join(intervals) if intervals else None


def format_bells(day: date, data: dict[str, Any] | None = None) -> str:
    bells = bell_for_date(day, data)
    weekday_name = "понедельник" if day.weekday() == 0 else "вторник–пятница"
    lines = [f"🔔 РАСПИСАНИЕ ЗВОНКОВ · {day:%d.%m.%Y}", f"Распорядок: {weekday_name}", "━━━━━━━━━━━━━━━━━━"]

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
    for number in range(1, 13):
        value = lessons.get(str(number))
        if value and value.get("start") and value.get("end"):
            lines.append(f"{number}-й урок — {value['start']}–{value['end']}")
        for item in breaks_by_position.get(number, []):
            lines.append(f"🍽 {item.get('label', 'Перерыв')} · {item.get('duration') or 'перерыв'}")

    practice = bells.get("practice_shifts", [])
    if practice:
        lines.extend(["", "📚 УЧЕБНАЯ ПРАКТИКА"])
        for shift in practice:
            lines.append(f"{shift['label']}: {shift['start']}–{shift['end']}")

    if day.isoformat() in (data or {}).get("overrides", {}):
        lines.extend(["", "⚠️ На эту дату действует опубликованное изменение звонков."])
    return "\n".join(lines)
