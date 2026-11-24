"""Extract lesson schedules from the college XLSX file using only the stdlib.

The source workbook is an OOXML ZIP package. We read its worksheet XML directly
so the bot has no dependency on desktop Excel, LibreOffice or a heavy XLSX SDK.
"""
from __future__ import annotations

import json
import posixpath
import re
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"m": MAIN_NS, "r": REL_NS, "pr": PKG_REL_NS}
DATE_RE = re.compile(r"(?P<date>\d{2}\.\d{2}\.\d{4})")
GROUP_RE = re.compile(r"^\d{2,4}$")


def _col_number(cell_ref: str) -> int:
    """Convert an A1 column name to a 1-based integer."""
    letters = re.match(r"[A-Z]+", cell_ref.upper())
    if not letters:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    number = 0
    for char in letters.group(0):
        number = number * 26 + ord(char) - 64
    return number


def _parse_scalar(value: str | None, cell_type: str | None, shared: list[str]) -> Any:
    if value is None:
        return None
    if cell_type == "s":
        try:
            return shared[int(value)]
        except (ValueError, IndexError):
            return value
    if cell_type == "b":
        return value == "1"
    if cell_type in ("str", "e"):
        return value
    # Numeric values are kept as ints where possible. Date fields in this
    # workbook are text, not Excel serial dates.
    try:
        number = float(value)
        return int(number) if number.is_integer() else number
    except ValueError:
        return value


def _load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    strings: list[str] = []
    for item in root.findall("m:si", NS):
        # A shared string can consist of multiple rich-text runs.
        strings.append("".join(t.text or "" for t in item.findall(".//m:t", NS)))
    return strings


def _worksheet_paths(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall("pr:Relationship", NS)
    }
    result: list[tuple[str, str]] = []
    for sheet in workbook.findall("m:sheets/m:sheet", NS):
        name = sheet.attrib["name"]
        rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
        target = targets[rel_id]
        if target.startswith("/"):
            path = target.lstrip("/")
        else:
            path = posixpath.normpath(posixpath.join("xl", target))
        result.append((name, path))
    return result


def _read_sheet(archive: zipfile.ZipFile, path: str, shared: list[str]) -> dict[int, dict[int, Any]]:
    root = ET.fromstring(archive.read(path))
    rows: dict[int, dict[int, Any]] = {}
    for row in root.findall(".//m:sheetData/m:row", NS):
        row_no = int(row.attrib["r"])
        row_values: dict[int, Any] = {}
        for cell in row.findall("m:c", NS):
            ref = cell.attrib.get("r")
            if not ref:
                continue
            col = _col_number(ref)
            cell_type = cell.attrib.get("t")
            if cell_type == "inlineStr":
                value = "".join(t.text or "" for t in cell.findall(".//m:t", NS))
            else:
                value_node = cell.find("m:v", NS)
                value = _parse_scalar(value_node.text if value_node is not None else None, cell_type, shared)
            if value is not None and value != "":
                row_values[col] = value
        if row_values:
            rows[row_no] = row_values
    return rows


def _text(value: Any) -> str | None:
    if value is None:
        return None
    result = str(value).strip()
    return result if result else None


def _is_pair_number(value: Any) -> bool:
    try:
        n = int(value)
    except (TypeError, ValueError):
        return False
    return 1 <= n <= 12 and str(value).strip() in {str(n), f"{n}.0"}


def _is_self_study(subject: Any, room: Any = None) -> bool:
    """Recognize the college's self-study marker without displaying it as a lesson."""
    subject_text = (_text(subject) or "").casefold()
    room_text = re.sub(r"\s+", "", (_text(room) or "").casefold())
    return bool(
        re.search(r"(?<![а-яё])с\s*/\s*р(?![а-яё])", subject_text)
        or room_text in {"с/р", "ср", "самостоятельнаяработа"}
        or "самостоятельная работа" in subject_text
    )


def _day_name(date_value: datetime) -> str:
    names = ["понедельник", "вторник", "среда", "четверг", "пятница", "суббота", "воскресенье"]
    return names[date_value.weekday()]


def parse_xlsx(path: str | Path) -> dict[str, Any]:
    """Parse every group and day from the uploaded schedule workbook."""
    file_path = Path(path)
    if not file_path.is_file():
        raise FileNotFoundError(file_path)
    if file_path.suffix.lower() != ".xlsx":
        raise ValueError("Поддерживаются только файлы .xlsx")

    with zipfile.ZipFile(file_path) as archive:
        shared = _load_shared_strings(archive)
        sheets = _worksheet_paths(archive)
        groups: dict[str, dict[str, Any]] = {}
        all_dates: set[str] = set()
        parsed_tables = 0

        for sheet_name, sheet_path in sheets:
            rows = _read_sheet(archive, sheet_path, shared)
            header_rows = [
                row_no for row_no, row in rows.items()
                if any(_text(value) == "День недели" for value in row.values())
            ]
            for table_index, header_row in enumerate(header_rows):
                group_row_no = header_row + 2
                subgroup_row_no = header_row + 3
                group_row = rows.get(group_row_no, {})
                subgroup_row = rows.get(subgroup_row_no, {})
                group_starts: list[tuple[str, int]] = []
                for col, value in group_row.items():
                    if col < 6:
                        continue
                    code = _text(value)
                    if not code or not GROUP_RE.fullmatch(code):
                        continue
                    # A group number is placed in the same column as its
                    # "Дисциплина" heading. This excludes unrelated numbers.
                    if _text(rows.get(header_row, {}).get(col)) == "Дисциплина":
                        group_starts.append((code, col))
                if not group_starts:
                    continue
                parsed_tables += 1
                next_header = header_rows[table_index + 1] if table_index + 1 < len(header_rows) else max(rows, default=header_row) + 1
                class_hour_row = header_row + 4
                first_data_row = header_row + 5

                for code, base_col in group_starts:
                    group = groups.setdefault(code, {"sheet": sheet_name, "days": {}})
                    # Class hour is shown in a separate row immediately before
                    # the dated weekday blocks; attach it to the first day.
                    class_hour_items: list[dict[str, Any]] = []
                    class_row = rows.get(class_hour_row, {})
                    for subgroup, subject_col in ((1, base_col), (2, base_col + 2)):
                        subject = _text(class_row.get(subject_col))
                        if subject:
                            class_hour_items.append({
                                "pair": 0,
                                "subject": subject,
                                "room": _text(class_row.get(base_col + 3)) or _text(class_row.get(base_col + 1)),
                                "teacher": None,
                                "subgroup": subgroup,
                            })

                    current_date: str | None = None
                    # Some first-year layouts have a separate pair-number
                    # column for subgroup 2; later-year layouts share column B.
                    subgroup2_pair_col = 3 if _text(group_row.get(3)) == "№" else 2
                    for row_no in range(first_data_row, next_header):
                        row = rows.get(row_no, {})
                        date_cell = _text(row.get(1))
                        if date_cell:
                            match = DATE_RE.search(date_cell)
                            if match:
                                date_obj = datetime.strptime(match.group("date"), "%d.%m.%Y")
                                current_date = date_obj.strftime("%Y-%m-%d")
                                all_dates.add(current_date)
                                group["days"].setdefault(current_date, [])
                        if not current_date:
                            continue

                        teacher_row = rows.get(row_no + 1, {})
                        subgroup_columns = ((1, base_col, 2), (2, base_col + 2, subgroup2_pair_col))
                        for subgroup, subject_col, pair_col in subgroup_columns:
                            pair_value = row.get(pair_col)
                            subject = _text(row.get(subject_col))
                            if not subject or not _is_pair_number(pair_value):
                                continue
                            room = _text(row.get(base_col + 3)) or _text(row.get(base_col + 1))
                            lesson = {
                                "pair": int(pair_value),
                                "subject": subject,
                                "room": room,
                                "teacher": _text(teacher_row.get(subject_col)),
                                "subgroup": subgroup,
                                "self_study": _is_self_study(subject, room),
                            }
                            group["days"].setdefault(current_date, []).append(lesson)

                    if class_hour_items and group["days"]:
                        first_date = min(group["days"])
                        existing = group["days"].setdefault(first_date, [])
                        existing[:0] = class_hour_items

        # Normalize and validate the extracted dataset.
        for group in groups.values():
            for date, lessons in group["days"].items():
                lessons.sort(key=lambda item: (item["pair"], item.get("subgroup", 0)))
        groups = {code: data for code, data in groups.items() if data["days"]}
        if not groups:
            raise ValueError("Не удалось найти группы и занятия. Проверьте структуру XLSX.")
        if not all_dates:
            raise ValueError("В файле не найдены даты занятий.")

        ordered_dates = sorted(all_dates)
        return {
            "source_file": file_path.name,
            "parsed_at": datetime.now().isoformat(timespec="seconds"),
            "date_from": ordered_dates[0],
            "date_to": ordered_dates[-1],
            "dates": ordered_dates,
            "groups": groups,
            "group_count": len(groups),
            "table_count": parsed_tables,
        }



def parse_replacement_xlsx(path: str | Path) -> dict[str, Any]:
    """Parse a dated daily replacement workbook (the college's compact template).

    Unlike the weekly timetable, this layout has repeated blocks headed by
    ``Группа:``. Every lesson is represented by a subject row followed by a
    teacher row. A lesson's pair is derived from the individual lesson number.
    """
    file_path = Path(path)
    if not file_path.is_file():
        raise FileNotFoundError(file_path)
    if file_path.suffix.lower() != ".xlsx":
        raise ValueError("Поддерживаются только файлы .xlsx")

    def number(value: Any) -> int | None:
        text = _text(value)
        if not text:
            return None
        match = re.fullmatch(r"(\d+)(?:\.0+)?", text)
        return int(match.group(1)) if match else None

    with zipfile.ZipFile(file_path) as archive:
        shared = _load_shared_strings(archive)
        groups: dict[str, dict[str, Any]] = {}
        all_dates: set[str] = set()
        parsed_blocks = 0
        title_text = ""
        sheet_names = _worksheet_paths(archive)
        for sheet_name, sheet_path in sheet_names:
            rows = _read_sheet(archive, sheet_path, shared)
            if not rows:
                continue

            title_text = " ".join(
                _text(value) or ""
                for row_no, row in rows.items() if row_no <= 3
                for value in row.values()
            )
            date_match = DATE_RE.search(file_path.name) or DATE_RE.search(title_text) or DATE_RE.search(sheet_name)
            if not date_match:
                continue
            day_obj = datetime.strptime(date_match.group("date"), "%d.%m.%Y").date()
            day_text = day_obj.isoformat()
            all_dates.add(day_text)

            # The top header has one "Дисциплина" column per group column.
            header_candidates = [
                (row_no, row) for row_no, row in rows.items()
                if any((_text(v) or "").casefold() == "дисциплина" for v in row.values())
                and any((_text(v) or "").casefold() == "пара" for v in row.values())
            ]
            if not header_candidates:
                continue
            header_row_no, header_row = header_candidates[0]
            discipline_cols = {
                col for col, value in header_row.items()
                if (_text(value) or "").casefold() == "дисциплина"
            }
            group_rows = [
                row_no for row_no, row in rows.items()
                if any((_text(value) or "").casefold() == "группа:" for value in row.values())
            ]
            for block_index, group_row_no in enumerate(group_rows):
                group_row = rows.get(group_row_no, {})
                block_groups = [
                    (str(_text(group_row.get(col))), col)
                    for col in sorted(discipline_cols)
                    if _text(group_row.get(col)) and GROUP_RE.fullmatch(str(_text(group_row.get(col))))
                ]
                if not block_groups:
                    continue
                parsed_blocks += 1
                next_group_row = group_rows[block_index + 1] if block_index + 1 < len(group_rows) else max(rows, default=group_row_no) + 1
                first_subject_row = group_row_no + 1

                for group_code, base_col in block_groups:
                    group = groups.setdefault(group_code, {"sheet": sheet_name, "days": {}})
                    day_lessons = group.setdefault("days", {}).setdefault(day_text, [])

                    # Subject rows alternate with teacher rows. Only read those
                    # subject rows so teacher names cannot be mistaken for subjects.
                    for row_no in range(first_subject_row, next_group_row, 2):
                        row = rows.get(row_no, {})
                        teacher_row = rows.get(row_no + 1, {})
                        lesson_number = number(row.get(2))
                        pair_value = number(row.get(1))
                        if lesson_number is None:
                            if pair_value is None:
                                continue
                            lesson_number = pair_value * 2 - 1
                        if not 1 <= lesson_number <= 20:
                            continue
                        pair = pair_value or ((lesson_number + 1) // 2)
                        if not 1 <= pair <= 10:
                            continue

                        room = _text(row.get(base_col + 3)) or _text(row.get(base_col + 1))
                        for subgroup, subject_col in ((1, base_col), (2, base_col + 2)):
                            subject = _text(row.get(subject_col))
                            if not subject:
                                continue
                            # Header labels and numeric markers are not disciplines.
                            if subject.casefold() in {"дисциплина", "преподаватель", "ауд.", "ауд"}:
                                continue
                            teacher = _text(teacher_row.get(subject_col))
                            item = {
                                "pair": pair,
                                "lesson_number": lesson_number,
                                "subject": subject,
                                "room": room,
                                "teacher": teacher,
                                "subgroup": subgroup,
                                "self_study": _is_self_study(subject, room),
                            }
                            if item not in day_lessons:
                                day_lessons.append(item)

        if not groups:
            raise ValueError(
                "Не удалось распознать XLSX как недельное расписание или таблицу замен. "
                "Ожидались строки «Группа:», «Пара», «Урок» и «Дисциплина»."
            )

    for group in groups.values():
        for lessons in group["days"].values():
            lessons.sort(key=lambda item: (item.get("pair", 0), item.get("lesson_number", 0), item.get("subgroup", 0)))

    dates = sorted(all_dates or {day for group in groups.values() for day in group.get("days", {})})
    if not dates:
        raise ValueError("В файле замен не удалось определить дату.")
    title_all = file_path.name.casefold() + " " + title_text.casefold()
    replacement_mode = "supplement" if "дополнительно" in title_all else "replace"
    return {
        "source_file": file_path.name,
        "parsed_at": datetime.now().isoformat(timespec="seconds"),
        "date_from": dates[0],
        "date_to": dates[-1],
        "dates": dates,
        "groups": groups,
        "group_count": len(groups),
        "table_count": parsed_blocks,
        "replacement_mode": replacement_mode,
        "format": "daily_replacement",
    }

def save_json(data: dict[str, Any], path: str | Path) -> None:
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_json(path: str | Path) -> dict[str, Any]:
    return json.loads(Path(path).read_text(encoding="utf-8"))


if __name__ == "__main__":
    import argparse

    cli = argparse.ArgumentParser(description="Разобрать XLSX с расписанием")
    cli.add_argument("file", help="Путь к .xlsx")
    cli.add_argument("--group", help="Показать только одну группу")
    cli.add_argument("--json", help="Сохранить результат в JSON")
    args = cli.parse_args()
    result = parse_xlsx(args.file)
    if args.json:
        save_json(result, args.json)
    print(f"Групп: {result['group_count']}; таблиц: {result['table_count']}; даты: {result['date_from']} — {result['date_to']}")
    if args.group:
        selected = result["groups"].get(str(args.group))
        if not selected:
            raise SystemExit(f"Группа {args.group} не найдена. Доступно: {', '.join(sorted(result['groups']))}")
        for date, lessons in sorted(selected["days"].items()):
            date_obj = datetime.strptime(date, "%Y-%m-%d")
            print(f"\n{date_obj.strftime('%d.%m.%Y')} ({_day_name(date_obj)})")
            for lesson in lessons:
                prefix = "Классный час" if lesson["pair"] == 0 else f"{lesson['pair']} пара"
                print(f"  {prefix}: {lesson['subject']} | {lesson.get('room') or 'кабинет не указан'} | {lesson.get('teacher') or 'преподаватель не указан'}")
