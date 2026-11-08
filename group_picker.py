"""MAX inline keyboards for choosing a course and study group."""
from __future__ import annotations

from collections import defaultdict
from math import ceil

PAGE_SIZE = 30
COLUMNS = 5


def course_for_group(group_code: str) -> int | None:
    value = str(group_code).strip()
    if not value.isdigit() or len(value) < 3:
        return None
    course = int(value) // 100
    return course if course >= 1 else None


def groups_by_course(groups: dict) -> dict[int, list[str]]:
    result: dict[int, list[str]] = defaultdict(list)
    for code in groups:
        course = course_for_group(str(code))
        if course is not None:
            result[course].append(str(code))
    return {course: sorted(codes, key=lambda x: (int(x), x)) for course, codes in sorted(result.items())}


def _button(text: str, payload: str) -> dict[str, str]:
    return {"type": "callback", "text": text, "payload": payload}


def _keyboard(buttons: list[list[dict[str, str]]]) -> list[list[dict[str, str]]]:
    # MAX allows at most 30 rows and 7 buttons in a row.
    return [row[:7] for row in buttons[:30] if row]


def home_keyboard() -> list[list[dict[str, str]]]:
    return _keyboard([
        [_button("📅 Сегодня", "menu:today"), _button("➡️ Завтра", "menu:tomorrow")],
        [_button("⬅️ Прошлая", "menu:prevweek"), _button("🗓 Текущая", "menu:week"), _button("⏭ Следующая", "menu:nextweek")],
        [_button("🎓 Выбрать группу", "menu:groups"), _button("👥 Моя группа", "menu:mygroup")],
        [_button("⏰ Звонки", "menu:bells")],
        [_button("ℹ️ Помощь", "menu:help")],
    ])


def course_keyboard(groups: dict) -> list[list[dict[str, str]]]:
    grouped = groups_by_course(groups)
    rows: list[list[dict[str, str]]] = []
    for course, codes in grouped.items():
        title = f"{course} курс · {len(codes)} групп" if course <= 4 else f"Другие · {len(codes)} групп"
        rows.append([_button(title, f"group:course:{course}")])
    rows.append([_button("↩️ В меню", "menu:home")])
    return _keyboard(rows)


def group_page_keyboard(groups: dict, course: int, page: int = 0) -> list[list[dict[str, str]]]:
    grouped = groups_by_course(groups)
    codes = grouped.get(course, [])
    pages = max(1, ceil(len(codes) / PAGE_SIZE))
    page = max(0, min(page, pages - 1))
    selected = codes[page * PAGE_SIZE:(page + 1) * PAGE_SIZE]
    rows: list[list[dict[str, str]]] = []
    for start in range(0, len(selected), COLUMNS):
        rows.append([_button(code, f"group:pick:{code}") for code in selected[start:start + COLUMNS]])
    nav: list[dict[str, str]] = []
    if page > 0:
        nav.append(_button("⬅️ Назад", f"group:page:{course}:{page - 1}"))
    nav.append(_button(f"{page + 1}/{pages}", "noop"))
    if page + 1 < pages:
        nav.append(_button("Вперёд ➡️", f"group:page:{course}:{page + 1}"))
    rows.append(nav)
    rows.append([_button("↩️ К курсам", "menu:groups")])
    return _keyboard(rows)
