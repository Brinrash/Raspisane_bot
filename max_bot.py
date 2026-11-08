from __future__ import annotations

import asyncio
import hmac
import json
import logging
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import urlparse
from uuid import uuid4
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from aiohttp import web

from config import (
    BELL_JSON, DATA_DIR, INCOMING_DIR, MAX_ADMIN_IDS, MAX_BOT_TOKEN, MAX_CA_BUNDLE,
    MAX_POLL_TIMEOUT, MAX_SOURCE_CHAT_ID, MAX_STATE_PATH, MAX_UPLOAD_BYTES,
    MAX_WEBHOOK_URL, WEBHOOK_SECRET, WEBHOOK_PATH, PORT,
    MAX_USER_DB, NOTIFY_USERS_ON_UPDATE, SCHEDULE_JSON, SCHEDULE_XLSX, TIMEZONE,
)
from bells import (
    bell_for_date, format_bells, load_bell_data, merge_bell_override,
    pair_time, parse_bell_document_dates, prune_bell_data, save_bell_data,
)
from group_picker import course_keyboard, group_page_keyboard, home_keyboard
from max_api import MaxApi, MaxApiError
from max_files import download_max_file
from schedule_manager import import_schedule_file, load_schedule, prune_incoming_files, retention_bounds
from storage import UserStorage

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("schedule-bot.max")

storage = UserStorage(MAX_USER_DB)
schedule: dict[str, Any] = {}
bell_schedule: dict[str, Any] = {"overrides": {}}


def local_today() -> date:
    try:
        return datetime.now(ZoneInfo(TIMEZONE)).date()
    except ZoneInfoNotFoundError:
        logger.warning("Неизвестный часовой пояс %s; используется системная дата", TIMEZONE)
        return date.today()


def verify_webhook_secret(expected: str, supplied: str | None) -> bool:
    """Constant-time comparison for the MAX webhook secret header."""
    return bool(expected) and hmac.compare_digest(expected, supplied or "")


def split_text(text: str, limit: int = 3850) -> list[str]:
    if len(text) <= limit:
        return [text]
    result: list[str] = []
    current = ""
    for line in text.splitlines(keepends=True):
        if current and len(current) + len(line) > limit:
            result.append(current.rstrip())
            current = ""
        if len(line) > limit:
            for start in range(0, len(line), limit):
                if current:
                    result.append(current.rstrip())
                    current = ""
                result.append(line[start:start + limit].rstrip())
        else:
            current += line
    if current.strip():
        result.append(current.rstrip())
    return result


def _subject_parts(lesson: dict[str, Any]) -> tuple[str, None]:
    """Remove technical subgroup marks; do not expose subgroup labels to users."""
    subject = str(lesson.get("subject") or "Без названия")
    subject = re.sub(r"\s*\(\s*[12]\s*п/г\s*\)", "", subject, flags=re.IGNORECASE)
    subject = re.sub(r"\s+", " ", subject).strip()
    return subject or "Без названия", None


def lesson_line(lesson: dict[str, Any], day: date | None = None, bell_data: dict[str, Any] | None = None) -> str:
    """Render an attended lesson or a visually separate self-study item."""
    pair = int(lesson.get("pair") or 0)
    subject, _ = _subject_parts(lesson)
    room = str(lesson.get("room") or "").strip()
    self_study = bool(lesson.get("self_study")) or bool(
        re.search(r"(?<![а-яё])с\s*/\s*р(?![а-яё])", str(lesson.get("subject") or ""), re.IGNORECASE)
        or re.sub(r"\s+", "", room.casefold()) in {"с/р", "ср", "самостоятельнаяработа"}
        or "самостоятельная работа" in str(lesson.get("subject") or "").casefold()
    )

    teacher = str(lesson.get("teacher") or "").strip()
    if self_study:
        # A self-study assignment does not require attendance: intentionally
        # show neither a pair number nor a bell time/room.
        subject = re.sub(r"\s*(?:с\s*/\s*р|самостоятельная\s+работа)\s*$", "", subject, flags=re.IGNORECASE).strip()
        lines = ["📝 САМОСТОЯТЕЛЬНАЯ РАБОТА"]
        if subject:
            lines.append(subject)
        if teacher:
            lines.append(f"   👨‍🏫 {teacher}")
        return "\n".join(lines)

    lesson_time = pair_time(day, pair, bell_data) if day is not None else None
    if pair == 0:
        heading = f"🟡 Классный час · {lesson_time}" if lesson_time else "🟡 Классный час"
    else:
        heading = f"🔹 {pair} пара · {lesson_time}" if lesson_time else f"🔹 {pair} пара"

    lines = [heading]
    if not (pair == 0 and subject.casefold() == "классный час"):
        lines.append(subject)
    if room:
        lines.append(f"   📍 Каб.: {room}")
    if teacher:
        lines.append(f"   👨‍🏫 {teacher}")
    return "\n".join(lines)


def _weekday_name(day: date) -> str:
    return ["ПОНЕДЕЛЬНИК", "ВТОРНИК", "СРЕДА", "ЧЕТВЕРГ", "ПЯТНИЦА", "СУББОТА", "ВОСКРЕСЕНЬЕ"][day.weekday()]


def _updated_time_label() -> str | None:
    value = schedule.get("schedule_updated_at") or schedule.get("parsed_at")
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        return parsed.strftime("%H:%M")
    except (TypeError, ValueError):
        return None


def format_day(
    group_code: str,
    day: date,
    day_data: list[dict] | None,
    bell_data: dict[str, Any] | None = None,
    *,
    include_group_header: bool = True,
) -> str:
    lines: list[str] = []
    if include_group_header:
        lines.append(f"🎓 ГРУППА: {group_code}")
        updated = _updated_time_label()
        if updated:
            lines.append(f"🕒 Обновлено: {updated}")
        lines.append("━━━━━━━━━━━━━━━━━━")
    lines.extend([f"📅 {_weekday_name(day)} {day:%d.%m.%Y}", "━━━━━━━━━━━━━━━━━━"])

    if not day_data:
        lines.append("На эту дату занятий в опубликованном расписании нет.")
        return "\n".join(lines)

    for lesson in day_data:
        lines.extend(["", lesson_line(lesson, day, bell_data)])
    return "\n".join(lines)

def extract_file_context(update: dict[str, Any], message: dict[str, Any]) -> dict[str, Any]:
    """Return file attachments and trusted origin metadata for direct or forwarded messages.

    MAX represents a forwarded post in ``message.link``; the original post body and
    attachments are nested under ``link.message`` while ``link.chat_id`` identifies
    the source channel/chat. The outer ``message.body`` can be null for a pure forward.
    """
    body = message.get("body") or {}
    link = message.get("link") or {}
    is_forward = str(link.get("type") or "").lower() == "forward"
    original = link.get("message") or {} if is_forward else {}
    content = original if is_forward and original else body

    attachments = content.get("attachments") or []
    # Be tolerant of a client/API variant that puts file attachments on the outer body.
    if is_forward and not attachments:
        attachments = body.get("attachments") or []

    recipient = message.get("recipient") or {}
    current_chat_id = recipient.get("chat_id", update.get("chat_id"))
    origin_chat_id = link.get("chat_id") if is_forward else current_chat_id
    message_mid = content.get("mid") or body.get("mid") or update.get("timestamp") or "event"

    return {
        "body": body,
        "content": content,
        "attachments": [item for item in attachments if item.get("type") == "file"],
        "is_forward": is_forward,
        "origin_chat_id": origin_chat_id,
        "message_mid": str(message_mid),
        "original_text": content.get("text") or "",
    }


def is_allowed_file_source(
    sender_id: int | None,
    origin_chat_id: int | str | None,
    source_chat_id: str,
    admin_ids: set[int],
) -> bool:
    """Allow admins to upload directly and optionally allow forwards from one source."""
    if sender_id is not None and sender_id in admin_ids:
        return True
    return bool(source_chat_id and origin_chat_id is not None and str(origin_chat_id) == source_chat_id)


def user_group(user_id: int) -> str | None:
    return storage.get_group(user_id)


def group_data(group_code: str) -> dict[str, Any] | None:
    return schedule.get("groups", {}).get(str(group_code))


def selected_day_text(user_id: int, day: date) -> str:
    code = user_group(user_id)
    if not code:
        return "Сначала выбери курс и группу кнопками ниже."
    data = group_data(code)
    if not data:
        return f"Группа {code} не найдена в текущем расписании. Выбери группу заново."
    return format_day(code, day, data.get("days", {}).get(day.isoformat()), bell_schedule)


def week_text(user_id: int, offset_weeks: int = 0) -> str:
    code = user_group(user_id)
    if not code:
        return "Сначала выбери курс и группу кнопками ниже."
    data = group_data(code)
    if not data:
        return f"Группа {code} не найдена в текущем расписании. Выбери группу заново."
    days = data.get("days", {})
    if not days:
        return "Для этой группы расписание не найдено."

    today = local_today()
    current_monday = today - timedelta(days=today.weekday())
    target_monday = current_monday + timedelta(weeks=offset_weeks)
    target_sunday = target_monday + timedelta(days=6)
    # Deliberately include past days in the selected calendar week. A timetable
    # received on Friday/Saturday still needs to be viewable for the whole week.
    dates = [
        value for value in sorted(days)
        if target_monday.isoformat() <= value <= target_sunday.isoformat()
    ]

    titles = {-1: "ПРЕДЫДУЩАЯ НЕДЕЛЯ", 0: "ТЕКУЩАЯ НЕДЕЛЯ", 1: "СЛЕДУЮЩАЯ НЕДЕЛЯ"}
    title = titles.get(offset_weeks, "НЕДЕЛЯ")
    lines = [f"🎓 ГРУППА: {code}"]
    updated = _updated_time_label()
    if updated:
        lines.append(f"🕒 Обновлено: {updated}")
    lines.extend([
        "━━━━━━━━━━━━━━━━━━",
        f"🗓 {title}",
        f"{target_monday:%d.%m.%Y} — {target_sunday:%d.%m.%Y}",
        "━━━━━━━━━━━━━━━━━━",
    ])

    if not dates:
        available = sorted(days)
        if available:
            first = datetime.strptime(available[0], "%Y-%m-%d").date()
            last = datetime.strptime(available[-1], "%Y-%m-%d").date()
            lines.append(
                f"На эту неделю занятий не найдено. В базе есть расписание: {first:%d.%m.%Y} — {last:%d.%m.%Y}."
            )
        else:
            lines.append("Расписание на эту неделю пока не опубликовано.")
        return "\n".join(lines)

    for date_text in dates:
        day = datetime.strptime(date_text, "%Y-%m-%d").date()
        lines.extend([
            "",
            format_day(code, day, days[date_text], bell_schedule, include_group_header=False),
        ])
    return "\n".join(lines)


def bells_text(day: date) -> str:
    """On weekends, show the next school day's bells instead of a dead-end notice."""
    overrides = bell_schedule.get("overrides", {})
    if day.weekday() >= 5 and day.isoformat() not in overrides:
        next_school_day = day + timedelta(days=(7 - day.weekday()))
        return (
            f"🔔 {day:%d.%m.%Y} — выходной, занятий нет.\n"
            f"Ближайший учебный день: {_weekday_name(next_school_day).capitalize()}, "
            f"{next_school_day:%d.%m.%Y}.\n\n"
            f"{format_bells(next_school_day, bell_schedule)}"
        )
    return format_bells(day, bell_schedule)


class ScheduleMaxBot:
    def __init__(self, api: MaxApi) -> None:
        self.api = api
        self.processed_events: set[str] = set()
        self.marker: int | None = None
        self.state_file = MAX_STATE_PATH
        self.pending_webhook_events: set[str] = set()

    async def load_state(self) -> None:
        if not self.state_file.exists():
            return
        try:
            state = json.loads(self.state_file.read_text(encoding="utf-8"))
            self.marker = state.get("marker")
            self.processed_events = set(state.get("processed_events", []))
        except (OSError, json.JSONDecodeError, TypeError):
            logger.warning("Не удалось прочитать MAX state; начинаю с последнего события")

    def save_state(self) -> None:
        state = {
            "marker": self.marker,
            "processed_events": list(self.processed_events)[-500:],
        }
        self.state_file.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

    async def send_user(self, user_id: int, text: str, buttons: list[list[dict[str, str]]] | None = None) -> None:
        parts = split_text(text)
        for index, part in enumerate(parts):
            keyboard = buttons if index == len(parts) - 1 else None
            await self.api.send_user(user_id, part, keyboard)
            if len(parts) > 1:
                await asyncio.sleep(0.55)

    async def answer(self, callback_id: str, text: str, buttons: list[list[dict[str, str]]] | None = None) -> None:
        await self.api.answer_callback(callback_id, text[:4000], buttons)

    async def send_admins(self, text: str) -> None:
        for admin_id in sorted(MAX_ADMIN_IDS):
            try:
                await self.send_user(admin_id, text, home_keyboard())
            except Exception:
                logger.exception("Не удалось отправить сообщение MAX-администратору %s", admin_id)

    async def notify_update(self, diff: dict[str, Any], source: str, filename: str) -> None:
        groups = diff.get("changed_groups", [])
        dates = diff.get("changed_dates", [])
        if groups:
            date_labels = []
            for value in dates:
                try:
                    date_labels.append(datetime.strptime(value, "%Y-%m-%d").strftime("%d.%m"))
                except ValueError:
                    date_labels.append(value)
            admin_text = (
                f"✅ Расписание обновлено ({source}).\n"
                f"Файл: {filename}\n"
                f"Изменены группы: {len(groups)}\n"
                f"Даты: {', '.join(date_labels) or 'не определены'}"
            )
        else:
            admin_text = f"Файл «{filename}» получен из {source}, но изменений не обнаружено."
        await self.send_admins(admin_text)

        if NOTIFY_USERS_ON_UPDATE and groups:
            for user_id, group_code in storage.users_for_groups(groups):
                try:
                    await self.send_user(
                        user_id,
                        f"🔔 Расписание группы {group_code} обновлено. Открой меню «На неделю», чтобы посмотреть изменения.",
                        home_keyboard(),
                    )
                except Exception:
                    logger.exception("Не удалось уведомить MAX-пользователя %s", user_id)

    async def apply_xlsx(self, path: Path, original_name: str, source: str) -> dict[str, Any]:
        global schedule
        new_data, diff = import_schedule_file(
            path,
            schedule,
            SCHEDULE_JSON,
            today=local_today(),
            source=source,
            source_name=original_name,
        )
        schedule = new_data
        logger.info(
            "Импортировано из %s: %s, changed_groups=%d, dates=%s",
            source, original_name, len(diff.get("changed_groups", [])), diff.get("changed_dates", []),
        )
        await self.notify_update(diff, source, original_name)
        prune_incoming_files(INCOMING_DIR, local_today())
        return diff

    async def apply_bell_file(self, path: Path, original_name: str, source: str) -> dict[str, Any]:
        global bell_schedule
        target_days, parsed = parse_bell_document_dates(path)
        start, end = retention_bounds(local_today())
        outside = [day for day in target_days if not start <= day <= end]
        if outside:
            labels = ", ".join(day.strftime("%d.%m.%Y") for day in outside)
            raise ValueError(
                f"Даты {labels} вне окна хранения ({start:%d.%m.%Y} — {end:%d.%m.%Y})."
            )
        for target_day in target_days:
            bell_schedule = merge_bell_override(bell_schedule, target_day, parsed, original_name)
        bell_schedule = prune_bell_data(bell_schedule, local_today())
        save_bell_data(bell_schedule, BELL_JSON)
        date_labels = ", ".join(day.strftime("%d.%m.%Y") for day in target_days)
        logger.info("Импортированы изменения звонков: %s, dates=%s, source=%s", original_name, date_labels, source)
        await self.send_admins(
            f"🔔 Изменения звонков сохранены.\nФайл: {original_name}\nДаты: {date_labels}\n"
            f"Уроков с временем в файле: {len(parsed.get('lessons', {}))}\n"
            f"Специальное расписание: {'да' if parsed.get('complete_timetable') else 'нет'}\nИсточник: {source}"
        )
        if NOTIFY_USERS_ON_UPDATE:
            notified: set[int] = set()
            for user_id, _group_code in storage.all_users():
                if user_id in notified:
                    continue
                notified.add(user_id)
                try:
                    await self.send_user(
                        user_id,
                        f"🔔 Опубликовано изменение расписания звонков на даты: {date_labels}. "
                        "Время занятий обновлено в расписании.",
                        home_keyboard(),
                    )
                except Exception:
                    logger.exception("Не удалось уведомить MAX-пользователя %s об изменении звонков", user_id)
        prune_incoming_files(INCOMING_DIR, local_today())
        return {"dates": [day.isoformat() for day in target_days], "lessons": len(parsed.get("lessons", {}))}

    async def show_group_picker(self, user_id: int, callback_id: str | None = None) -> None:
        text = "🎓 Выбери курс:" if schedule.get("groups") else "В расписании пока нет групп. Администратор должен загрузить XLSX."
        keyboard = course_keyboard(schedule.get("groups", {})) if schedule.get("groups") else home_keyboard()
        if callback_id:
            await self.answer(callback_id, text, keyboard)
        else:
            await self.send_user(user_id, text, keyboard)

    async def on_callback(self, update: dict[str, Any]) -> None:
        callback = update.get("callback") or {}
        callback_id = str(callback.get("callback_id") or "")
        user = callback.get("user") or {}
        user_id = int(user.get("user_id") or 0)
        payload = str(callback.get("payload") or "")
        if not callback_id or not user_id:
            logger.warning("MAX callback без callback_id или user_id")
            return

        if payload == "noop":
            await self.answer(callback_id, "Страница выбора групп")
            return
        if payload in {"menu:home", "menu:help"}:
            text = (
                "Привет! Я бот расписания занятий.\n\n"
                "Выбери группу, затем смотри расписание на сегодня, завтра или неделю.\n"
                "Команды также можно отправлять текстом: /today, /tomorrow, /prevweek, /week, /nextweek, /bells, /group, /resetgroup, /myid."
                if payload == "menu:help" else "Главное меню расписания."
            )
            await self.answer(callback_id, text, home_keyboard())
            return
        if payload == "menu:groups":
            await self.show_group_picker(user_id, callback_id)
            return
        if payload == "menu:resetgroup":
            storage.clear_group(user_id)
            await self.answer(callback_id, "Сохранённая группа сброшена. Теперь выбери курс и группу:", course_keyboard(schedule.get("groups", {})))
            return
        if payload == "menu:today":
            await self.answer(callback_id, selected_day_text(user_id, local_today()), home_keyboard())
            return
        if payload == "menu:tomorrow":
            await self.answer(callback_id, selected_day_text(user_id, local_today() + timedelta(days=1)), home_keyboard())
            return
        if payload == "menu:prevweek":
            await self.answer(callback_id, week_text(user_id, -1), home_keyboard())
            return
        if payload == "menu:week":
            await self.answer(callback_id, week_text(user_id, 0), home_keyboard())
            return
        if payload == "menu:nextweek":
            await self.answer(callback_id, week_text(user_id, 1), home_keyboard())
            return
        if payload == "menu:bells":
            await self.answer(callback_id, bells_text(local_today()), home_keyboard())
            return
        if payload == "menu:mygroup":
            code = user_group(user_id)
            text = f"Твоя группа: {code}." if code else "Группа ещё не выбрана."
            buttons = home_keyboard() if code else course_keyboard(schedule.get("groups", {}))
            await self.answer(callback_id, text, buttons)
            return

        match = re.fullmatch(r"group:course:(\d+)", payload)
        if match:
            course = int(match.group(1))
            count = len([code for code in schedule.get("groups", {}) if str(code).isdigit() and int(code) // 100 == course])
            if not count:
                await self.answer(callback_id, "В этом курсе нет групп. Попробуй другой курс.", course_keyboard(schedule.get("groups", {})))
            else:
                await self.answer(callback_id, f"Выбери группу · {course if course <= 4 else 'другие'} курс:", group_page_keyboard(schedule.get("groups", {}), course, 0))
            return

        match = re.fullmatch(r"group:page:(\d+):(\d+)", payload)
        if match:
            course, page = map(int, match.groups())
            await self.answer(callback_id, f"Выбери группу · страница {page + 1}:", group_page_keyboard(schedule.get("groups", {}), course, page))
            return

        match = re.fullmatch(r"group:pick:([A-Za-z0-9_-]+)", payload)
        if match:
            code = match.group(1)
            if code not in schedule.get("groups", {}):
                await self.answer(callback_id, "Этой группы больше нет в расписании. Выбери заново.", course_keyboard(schedule.get("groups", {})))
                return
            storage.set_group(user_id, code)
            await self.answer(
                callback_id,
                f"✅ Группа {code} сохранена. Теперь можно смотреть расписание.",
                home_keyboard(),
            )
            return

        logger.warning("Неизвестный или устаревший callback MAX: payload=%r user_id=%s", payload, user_id)
        await self.answer(callback_id, "Эта кнопка устарела. Открой главное меню командой /start.", home_keyboard())

    async def handle_file_attachments(self, update: dict[str, Any], message: dict[str, Any], sender_id: int | None, chat_id: int | None) -> bool:
        context = extract_file_context(update, message)
        candidates = context["attachments"]
        if not candidates:
            return False

        origin_chat_id = context["origin_chat_id"]
        is_admin = sender_id in MAX_ADMIN_IDS if sender_id is not None else False
        allowed = is_allowed_file_source(sender_id, origin_chat_id, MAX_SOURCE_CHAT_ID, MAX_ADMIN_IDS)
        if not allowed:
            logger.warning(
                "Вложение отклонено: sender_id=%s, origin_chat_id=%s, expected_source=%s, forwarded=%s",
                sender_id, origin_chat_id, MAX_SOURCE_CHAT_ID or "(not set)", context["is_forward"],
            )
            if sender_id is not None:
                await self.send_user(
                    sender_id,
                    "⛔ Файл не обработан: отправь его из разрешённого канала или обратись к администратору бота.",
                    home_keyboard(),
                )
            return True

        if context["is_forward"]:
            source_label = f"пересланная публикация MAX (chat_id={origin_chat_id})"
        else:
            source_label = f"MAX chat_id={origin_chat_id if origin_chat_id is not None else chat_id}"

        for attachment in candidates:
            file_payload = attachment.get("payload") or {}
            filename = str(file_payload.get("filename") or "schedule.xlsx")
            suffix = Path(filename).suffix.lower()
            if suffix not in {".xlsx", ".pdf", ".docx"}:
                logger.info("Пропущено неподдерживаемое вложение: %s", filename)
                continue
            url = file_payload.get("url")
            if not url:
                notice = f"⚠️ MAX передал файл «{filename}», но не передал URL вложения."
                await self.send_admins(notice)
                if sender_id is not None:
                    await self.send_user(sender_id, notice, home_keyboard())
                continue
            safe_name = Path(filename).name.replace("..", "_")
            stamp = re.sub(r"[^A-Za-z0-9_-]", "_", context["message_mid"])[:40] or uuid4().hex[:12]
            path = INCOMING_DIR / f"max_{stamp}_{safe_name}"
            try:
                await download_max_file(str(url), path, MAX_UPLOAD_BYTES)
                if suffix == ".xlsx":
                    diff = await self.apply_xlsx(path, filename, source_label)
                    result = (
                        f"✅ Файл «{filename}» обработан.\n"
                        f"Источник: {source_label}\n"
                        f"Изменено групп: {len(diff.get('changed_groups', []))}\n"
                        f"Даты: {', '.join(diff.get('changed_dates', [])) or 'без изменений'}\n"
                        f"Хранение: {diff.get('retention_start')} — {diff.get('retention_end')}"
                    )
                    if sender_id is not None:
                        await self.send_user(sender_id, result, home_keyboard())
                elif any(
                    word in (filename + " " + str(context.get("original_text") or "")).casefold()
                    for word in ("звон", "распорядок", "время занятий")
                ):
                    bell_result = await self.apply_bell_file(path, filename, source_label)
                    result = (
                        f"✅ Изменения звонков приняты.\nДаты: {', '.join(bell_result['dates'])}\n"
                        f"Уроков с временем: {bell_result['lessons']}"
                    )
                    if sender_id is not None:
                        await self.send_user(sender_id, result, home_keyboard())
                else:
                    notice = (
                        f"📄 Файл «{filename}» получен, но не распознан как расписание занятий или звонков. "
                        "Для занятий нужен XLSX; для изменений звонков — DOCX или PDF с датой в имени файла/тексте "
                        "и названием, содержащим «звонки» или «распорядок»."
                    )
                    await self.send_admins(notice)
                    if sender_id is not None:
                        await self.send_user(sender_id, notice, home_keyboard())
                # Parsed source files are temporary; keep only current-window inputs on error.
                path.unlink(missing_ok=True)
            except Exception as exc:
                logger.exception("Не удалось обработать MAX-вложение %s", filename)
                notice = f"⚠️ Не удалось обработать файл «{filename}»: {exc}. Старое расписание сохранено."
                await self.send_admins(notice)
                if sender_id is not None:
                    await self.send_user(sender_id, notice, home_keyboard())
                prune_incoming_files(INCOMING_DIR, local_today())
        return True

    async def handle_message(self, update: dict[str, Any]) -> None:
        message = update.get("message") or {}
        sender = message.get("sender") or update.get("user") or {}
        sender_id_raw = sender.get("user_id")
        sender_id = int(sender_id_raw) if sender_id_raw is not None else None
        recipient = message.get("recipient") or {}
        chat_id_raw = recipient.get("chat_id", update.get("chat_id"))
        chat_id = int(chat_id_raw) if chat_id_raw is not None else None
        chat_type = str(recipient.get("chat_type") or "dialog").lower()
        body = message.get("body") or {}

        if await self.handle_file_attachments(update, message, sender_id, chat_id):
            return
        if sender_id is None:
            return

        text = str(body.get("text") or "").strip()
        normalized = text.lower()
        if not text:
            return

        # Reply to group/channel commands in that chat; private dialog replies go to user_id.
        async def reply(response: str, buttons: list[list[dict[str, str]]] | None = None) -> None:
            chunks = split_text(response)
            for index, chunk in enumerate(chunks):
                kb = buttons if index == len(chunks) - 1 else None
                if chat_type in {"chat", "channel"} and chat_id is not None:
                    await self.api.send_chat(chat_id, chunk, kb)
                else:
                    await self.api.send_user(sender_id, chunk, kb)

        if normalized in {"/start", "start", "начать", "привет"}:
            current = user_group(sender_id)
            text_out = "Привет! Я бот расписания занятий.\n\n"
            text_out += f"Твоя группа: {current}.\n\n" if current else "Сначала выбери курс и группу.\n\n"
            text_out += "Выбери действие кнопками ниже."
            await reply(text_out, home_keyboard())
            if not current:
                await reply("🎓 Выбери курс:", course_keyboard(schedule.get("groups", {})))
            return

        if normalized in {"/help", "help", "помощь"}:
            await reply("Команды: /today, /tomorrow, /prevweek, /week, /nextweek, /bells, /group, /resetgroup, /myid, /status. Также используй кнопки меню.", home_keyboard())
            return
        if normalized in {"/myid", "myid", "мой id", "/id"}:
            await reply(f"Твой MAX user_id: {sender_id}\nДля настройки администратора добавь это число в MAX_ADMIN_IDS файла .env.", home_keyboard())
            if chat_id is not None:
                logger.info("MAX chat_id seen: %s (type=%s)", chat_id, chat_type)
            return
        if normalized in {"/chatid", "chatid", "id чата"}:
            await reply(f"MAX chat_id: {chat_id if chat_id is not None else 'не передан'}\nТип чата: {chat_type}", home_keyboard())
            return
        if normalized in {"/group", "group", "группа", "выбрать группу", "сменить группу"}:
            await reply("🎓 Выбери курс:", course_keyboard(schedule.get("groups", {})))
            return
        if normalized in {"/resetgroup", "/clear_group", "сбросить группу", "сбросить мою группу"}:
            was_set = storage.clear_group(sender_id)
            message_text = (
                "Сохранённая группа сброшена. Выбери курс и группу заново."
                if was_set else "У тебя ещё не была сохранена группа. Выбери курс и группу."
            )
            await reply(message_text, course_keyboard(schedule.get("groups", {})))
            return
        if normalized in {"/today", "сегодня", "📅 сегодня"}:
            await reply(selected_day_text(sender_id, local_today()), home_keyboard())
            return
        if normalized in {"/tomorrow", "завтра", "➡️ завтра"}:
            await reply(selected_day_text(sender_id, local_today() + timedelta(days=1)), home_keyboard())
            return
        if normalized in {"/prevweek", "/previousweek", "прошлая неделя", "предыдущая неделя"}:
            await reply(week_text(sender_id, -1), home_keyboard())
            return
        if normalized in {"/week", "неделя", "на неделю", "🗓 на неделю"}:
            await reply(week_text(sender_id, 0), home_keyboard())
            return
        if normalized in {"/nextweek", "следующая неделя", "⏭ следующая"}:
            await reply(week_text(sender_id, 1), home_keyboard())
            return
        if normalized in {"/bells", "звонки", "расписание звонков", "⏰ звонки"}:
            await reply(bells_text(local_today()), home_keyboard())
            return
        if normalized.startswith("/bells ") or normalized.startswith("звонки "):
            raw_date = normalized.split(maxsplit=1)[1]
            try:
                target_day = datetime.strptime(raw_date, "%d.%m.%Y").date()
                await reply(bells_text(target_day), home_keyboard())
            except ValueError:
                await reply("Укажи дату в формате /bells 09.11.2026", home_keyboard())
            return
        if normalized in {"/mygroup", "моя группа"}:
            code = user_group(sender_id)
            await reply(f"Твоя группа: {code}." if code else "Группа ещё не выбрана.", home_keyboard() if code else course_keyboard(schedule.get("groups", {})))
            return
        if normalized == "/status":
            if sender_id not in MAX_ADMIN_IDS:
                await reply("Команда доступна только администратору.", home_keyboard())
                return
            last = schedule.get("update_history", [])[-1] if schedule.get("update_history") else {}
            await reply(
                "📊 Состояние расписания\n"
                f"Групп: {schedule.get('group_count', len(schedule.get('groups', {})))}\n"
                f"Период: {schedule.get('date_from', '—')} — {schedule.get('date_to', '—')}\n"
                f"Последнее обновление: {last.get('at', schedule.get('schedule_updated_at', 'неизвестно'))}\n"
                f"Источник: {last.get('source', 'исходный файл')}\n"
                f"Окно хранения: {schedule.get('retention_start', '—')} — {schedule.get('retention_end', '—')}\n"
                f"Изменения звонков: {len(bell_schedule.get('overrides', {}))} дат\n"
                f"MAX source chat: {MAX_SOURCE_CHAT_ID or 'не ограничен; доступны только файлы от админов и из чата-источника'}",
                home_keyboard(),
            )
            return
        if normalized in {"/menu", "меню"}:
            await reply("Главное меню расписания.", home_keyboard())
            return
        await reply("Не понял запрос. Используй кнопки меню или напиши /help.", home_keyboard())

    async def handle_update(self, update: dict[str, Any]) -> None:
        update_type = update.get("update_type")
        if update_type == "bot_started":
            user = update.get("user") or {}
            user_id = user.get("user_id")
            if user_id:
                await self.send_user(int(user_id), "Привет! Я бот расписания. Выбери курс и группу.", home_keyboard())
                if not user_group(int(user_id)):
                    await self.show_group_picker(int(user_id))
            return
        if update_type == "message_callback":
            await self.on_callback(update)
            return
        if update_type == "message_created":
            await self.handle_message(update)
            return
        # Edits are ignored to avoid processing the same uploaded schedule again.

    @staticmethod
    def event_key(update: dict[str, Any]) -> str:
        """Stable best-effort key for deduplicating polling and webhook deliveries."""
        event_message = update.get("message") or {}
        event_link = event_message.get("link") or {}
        event_original = event_link.get("message") or {}
        event_body = event_message.get("body") or {}
        event_mid = event_body.get("mid") or event_original.get("mid") or ""
        callback_id = (update.get("callback") or {}).get("callback_id", "")
        return (
            f"{update.get('update_type')}:{event_mid}:"
            f"{update.get('timestamp', '')}:{callback_id}"
        )

    async def webhook_forever(self) -> None:
        """Run an HTTP endpoint and subscribe it to MAX Bot API events."""
        if not MAX_WEBHOOK_URL:
            raise RuntimeError("Для Webhook необходимо задать MAX_WEBHOOK_URL или домен Bothost в DOMAIN.")
        if not re.fullmatch(r"[A-Za-z0-9_-]{5,256}", WEBHOOK_SECRET or ""):
            raise RuntimeError(
                "Для Webhook задай WEBHOOK_SECRET длиной 5–256 символов "
                "из латинских букв, цифр, дефиса и подчёркивания."
            )
        await self.load_state()

        async def health(_request: web.Request) -> web.Response:
            return web.json_response({"status": "ok", "service": "max-schedule-bot"})

        async def receive(request: web.Request) -> web.Response:
            supplied = request.headers.get("X-Max-Bot-Api-Secret", "")
            if not verify_webhook_secret(WEBHOOK_SECRET, supplied):
                logger.warning("Отклонён Webhook с неверным X-Max-Bot-Api-Secret")
                return web.json_response({"ok": False, "error": "unauthorized"}, status=401)
            try:
                update = await request.json()
            except (ValueError, json.JSONDecodeError):
                return web.json_response({"ok": False, "error": "invalid_json"}, status=400)
            if not isinstance(update, dict) or not update.get("update_type"):
                return web.json_response({"ok": False, "error": "invalid_update"}, status=400)

            event_key = self.event_key(update)
            if event_key in self.processed_events or event_key in self.pending_webhook_events:
                return web.json_response({"ok": True, "duplicate": True})

            self.pending_webhook_events.add(event_key)

            async def process_event() -> None:
                try:
                    await self.handle_update(update)
                    self.processed_events.add(event_key)
                    self.save_state()
                    logger.info("Обработан MAX Webhook event: %s", update.get("update_type"))
                except Exception:
                    logger.exception("Ошибка обработки MAX Webhook event type=%s", update.get("update_type"))
                    try:
                        await self.send_admins(
                            f"⚠️ Ошибка обработки события MAX типа {update.get('update_type')}. Подробности в логах хостинга."
                        )
                    except Exception:
                        logger.exception("Не удалось уведомить администраторов об ошибке Webhook")
                finally:
                    self.pending_webhook_events.discard(event_key)

            asyncio.create_task(process_event())
            # MAX expects a quick 200; heavy parsing and notifications run asynchronously.
            return web.json_response({"ok": True})

        app = web.Application(client_max_size=2 * 1024 * 1024)
        app.router.add_get("/health", health)
        app.router.add_post(WEBHOOK_PATH, receive)
        runner = web.AppRunner(app, access_log=None)
        await runner.setup()
        site = web.TCPSite(runner, host="0.0.0.0", port=PORT)
        await site.start()
        logger.info("MAX Webhook HTTP server запущен: port=%s path=%s", PORT, WEBHOOK_PATH)

        try:
            result = await self.api.request(
                "POST",
                "/subscriptions",
                payload={
                    "url": MAX_WEBHOOK_URL,
                    "update_types": ["message_created", "message_callback", "bot_started"],
                    "secret": WEBHOOK_SECRET,
                },
            )
            logger.info("Подписка MAX Webhook настроена: %s", result.get("success", True))
            await asyncio.Event().wait()
        finally:
            await runner.cleanup()

    async def poll_forever(self) -> None:
        await self.load_state()
        logger.info("MAX Long Polling запущен; source chat id=%s", MAX_SOURCE_CHAT_ID or "не задан")
        if self.marker is None:
            logger.warning("MAX marker отсутствует: первый запрос получит только последнее доступное событие; отправь боту /start для проверки")
        while True:
            try:
                payload = await self.api.get_updates(self.marker, MAX_POLL_TIMEOUT)
                updates = payload.get("updates", []) or []
                for update in updates:
                    update_key = self.event_key(update)
                    if update_key in self.processed_events:
                        continue
                    try:
                        await self.handle_update(update)
                    except Exception:
                        logger.exception("Ошибка обработки MAX event type=%s", update.get("update_type"))
                        await self.send_admins(f"⚠️ Ошибка обработки события MAX типа {update.get('update_type')} — подробности смотри в консоли.")
                    self.processed_events.add(update_key)
                if "marker" in payload:
                    self.marker = payload.get("marker")
                self.save_state()
            except asyncio.CancelledError:
                raise
            except MaxApiError as exc:
                logger.error("MAX API: %s", exc)
                if "401" in str(exc):
                    logger.error("Проверь MAX_BOT_TOKEN в .env")
                    await asyncio.sleep(60)
                else:
                    await asyncio.sleep(15)
            except Exception:
                logger.exception("Ошибка MAX Long Polling")
                await asyncio.sleep(15)


async def main() -> None:
    global schedule, bell_schedule
    if not MAX_BOT_TOKEN:
        raise RuntimeError("Не задан MAX_BOT_TOKEN. Открой .env и вставь API-токен MAX-бота.")
    schedule = load_schedule(SCHEDULE_XLSX, SCHEDULE_JSON, today=local_today())
    bell_schedule = prune_bell_data(load_bell_data(BELL_JSON), local_today())
    save_bell_data(bell_schedule, BELL_JSON)
    prune_incoming_files(INCOMING_DIR, local_today())
    logger.info(
        "Загружено расписание: %s групп, %s таблиц, %s — %s",
        schedule.get("group_count", len(schedule.get("groups", {}))),
        schedule.get("table_count", 0), schedule.get("date_from"), schedule.get("date_to"),
    )
    if not MAX_ADMIN_IDS:
        logger.warning("MAX_ADMIN_IDS не задан: загрузка XLSX в личном диалоге недоступна; можно использовать настроенный чат-источник")

    async with MaxApi(MAX_BOT_TOKEN, MAX_CA_BUNDLE) as api:
        me = await api.get_me()
        logger.info("Авторизация MAX успешна: %s", me.get("name") or me.get("username") or me.get("user_id") or "бот")
        bot = ScheduleMaxBot(api)
        if MAX_WEBHOOK_URL:
            logger.info("Режим работы: MAX Webhook (%s)", MAX_WEBHOOK_URL)
            await bot.webhook_forever()
        else:
            logger.info("Режим работы: MAX Long Polling")
            await bot.poll_forever()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Остановлено пользователем")
