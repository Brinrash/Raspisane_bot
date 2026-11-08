from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from group_picker import course_keyboard, group_page_keyboard, groups_by_course, home_keyboard
from storage import UserStorage
from max_api import MaxApi
from max_bot import split_text, lesson_line, extract_file_context, is_allowed_file_source, verify_webhook_secret
import max_bot


class GroupPickerTests(unittest.TestCase):
    def test_groups_are_split_by_hundreds(self):
        groups = {"101": {}, "201": {}, "315": {}, "403": {}, "521": {}}
        self.assertEqual(groups_by_course(groups), {
            1: ["101"], 2: ["201"], 3: ["315"], 4: ["403"], 5: ["521"]
        })

    def test_course_keyboard_uses_max_callback_buttons(self):
        keyboard = course_keyboard({"101": {}, "315": {}, "403": {}})
        self.assertTrue(keyboard)
        self.assertEqual(keyboard[0][0]["type"], "callback")
        self.assertEqual(keyboard[0][0]["payload"], "group:course:1")

    def test_large_course_list_is_paginated(self):
        groups = {str(300 + i): {} for i in range(1, 46)}
        first = group_page_keyboard(groups, 3, 0)
        self.assertLessEqual(len(first), 30)
        self.assertEqual(first[-2][0]["payload"], "noop")
        self.assertEqual(first[-1][0]["payload"], "menu:groups")

    def test_home_keyboard_respects_max_limits(self):
        keyboard = home_keyboard()
        self.assertLessEqual(len(keyboard), 30)
        self.assertTrue(all(len(row) <= 7 for row in keyboard))
        payloads = {button["payload"] for row in keyboard for button in row}
        self.assertIn("menu:prevweek", payloads)
        self.assertIn("menu:week", payloads)
        self.assertIn("menu:nextweek", payloads)


class MaxApiFormattingTests(unittest.TestCase):
    def test_message_body_keyboard_shape(self):
        buttons = [[{"type": "callback", "text": "Today", "payload": "menu:today"}]]
        body = MaxApi.message_body("Hello", buttons)
        self.assertEqual(body["attachments"][0]["type"], "inline_keyboard")
        self.assertEqual(body["attachments"][0]["payload"]["buttons"], buttons)

    def test_split_text_under_limit(self):
        self.assertEqual(split_text("line 1\nline 2", limit=100), ["line 1\nline 2"])

    def test_split_text_over_limit(self):
        result = split_text("A" * 50 + "\n" + "B" * 50, limit=60)
        self.assertTrue(all(len(part) <= 60 for part in result))
        self.assertEqual("".join(result).replace("\n", ""), "A" * 50 + "B" * 50)

    def test_lesson_line_contains_details(self):
        text = lesson_line({"pair": 2, "subject": "Математика", "room": "2/33", "teacher": "Иванов", "subgroup": 1})
        self.assertIn("🔹 2 пара", text)
        self.assertIn("Математика", text)
        self.assertIn("Каб.: 2/33", text)
        self.assertIn("Иванов", text)
        # The parser's physical first sub-column is not enough to label a lesson
        # as first subgroup; subgroup is displayed only when the source says so.
        self.assertNotIn("1 подгруппа", text)

    def test_class_hour_is_not_duplicated(self):
        text = lesson_line({"pair": 0, "subject": "Классный час", "room": "201"}, date(2026, 11, 9))
        self.assertEqual(text.count("Классный час"), 1)
        self.assertIn("08:15–09:00", text)

    def test_subgroup_marker_is_removed_without_label(self):
        text = lesson_line({
            "pair": 3,
            "subject": "Правовое обеспечение (2 п/г)",
            "room": "129",
            "teacher": "Юранева А.В.",
            "subgroup": 1,
        })
        self.assertIn("Правовое обеспечение", text)
        self.assertNotIn("(2 п/г)", text)
        self.assertNotIn("подгруппа", text)


class ForwardedFileTests(unittest.TestCase):
    def _forward_update(self):
        return {
            "update_type": "message_created",
            "timestamp": 1762545600000,
            "message": {
                "sender": {"user_id": 4242},
                "recipient": {"chat_id": 9999, "chat_type": "dialog", "user_id": 4242},
                # MAX may return null body for a pure forwarded post.
                "body": None,
                "link": {
                    "type": "forward",
                    "chat_id": -69158905315551,
                    "message": {
                        "mid": "original-post-mid",
                        "text": "Расписание занятий",
                        "attachments": [{
                            "type": "file",
                            "payload": {
                                "filename": "schedule.xlsx",
                                "url": "https://files.example/schedule.xlsx",
                                "size": 12345,
                            },
                        }],
                    },
                },
            },
        }

    def test_extracts_file_from_forwarded_post(self):
        update = self._forward_update()
        message = update["message"]
        context = extract_file_context(update, message)
        self.assertTrue(context["is_forward"])
        self.assertEqual(context["origin_chat_id"], -69158905315551)
        self.assertEqual(context["message_mid"], "original-post-mid")
        self.assertEqual(context["attachments"][0]["payload"]["filename"], "schedule.xlsx")

    def test_configured_source_accepts_forward_without_admin_rights(self):
        self.assertTrue(is_allowed_file_source(
            sender_id=4242,
            origin_chat_id=-69158905315551,
            source_chat_id="-69158905315551",
            admin_ids=set(),
        ))

    def test_other_forward_source_is_rejected(self):
        self.assertFalse(is_allowed_file_source(
            sender_id=4242,
            origin_chat_id=-123,
            source_chat_id="-69158905315551",
            admin_ids=set(),
        ))

    def test_admin_can_upload_directly(self):
        self.assertTrue(is_allowed_file_source(
            sender_id=4242,
            origin_chat_id=9999,
            source_chat_id="-69158905315551",
            admin_ids={4242},
        ))


class WeekAndBellViewTests(unittest.TestCase):
    def setUp(self):
        self.old_storage = max_bot.storage
        self.old_schedule = max_bot.schedule
        self.old_bells = max_bot.bell_schedule
        self.old_today = max_bot.local_today
        self.temp = tempfile.TemporaryDirectory()
        max_bot.storage = UserStorage(Path(self.temp.name) / "users.sqlite3")
        max_bot.storage.set_group(12345, "315")
        max_bot.local_today = lambda: date(2026, 11, 7)
        max_bot.bell_schedule = {"overrides": {}}

    def tearDown(self):
        max_bot.storage = self.old_storage
        max_bot.schedule = self.old_schedule
        max_bot.bell_schedule = self.old_bells
        max_bot.local_today = self.old_today
        self.temp.cleanup()

    def test_current_week_includes_already_passed_days(self):
        max_bot.schedule = {
            "schedule_updated_at": "2026-11-07T21:33:00",
            "groups": {"315": {"days": {
                "2026-11-02": [{"pair": 1, "subject": "Понедельник предмет", "room": "2/17"}],
                "2026-11-03": [{"pair": 2, "subject": "Вторник предмет", "room": "2/18"}],
            }}},
        }
        text = max_bot.week_text(12345, 0)
        self.assertIn("ТЕКУЩАЯ НЕДЕЛЯ", text)
        self.assertIn("ПОНЕДЕЛЬНИК 02.11.2026", text)
        self.assertIn("ВТОРНИК 03.11.2026", text)
        self.assertNotIn("оставшиеся дни", text)
        self.assertIn("🎓 ГРУППА: 315", text)
        self.assertIn("🕒 Обновлено: 21:33", text)
        self.assertIn("🔹 1 пара · 09:20–10:05 / 10:15–11:00", text)
        self.assertIn("📍 Каб.: 2/17", text)

    def test_previous_week_is_available_in_retention_window(self):
        max_bot.schedule = {
            "groups": {"315": {"days": {
                "2026-10-26": [{"pair": 1, "subject": "Предыдущая неделя"}],
            }}},
        }
        text = max_bot.week_text(12345, -1)
        self.assertIn("ПРЕДЫДУЩАЯ НЕДЕЛЯ", text)
        self.assertIn("26.10.2026 — 01.11.2026", text)
        self.assertIn("Предыдущая неделя", text)

    def test_weekend_bells_show_next_school_day(self):
        text = max_bot.bells_text(date(2026, 11, 7))
        self.assertIn("выходной, занятий нет", text)
        self.assertIn("Понедельник, 09.11.2026", text)
        self.assertIn("09:20–10:05", text)


class StorageTests(unittest.TestCase):
    def test_max_user_group_round_trip(self):
        with tempfile.TemporaryDirectory() as temp:
            storage = UserStorage(Path(temp) / "max_users.sqlite3")
            storage.set_group(987654321, "315")
            self.assertEqual(storage.get_group(987654321), "315")
            self.assertEqual(storage.users_for_groups(["315"]), [(987654321, "315")])
            self.assertEqual(storage.all_users(), [(987654321, "315")])


if __name__ == "__main__":
    unittest.main()


class WebhookSecurityTests(unittest.TestCase):
    def test_correct_secret_is_accepted(self):
        self.assertTrue(verify_webhook_secret("abc-DEF_12345", "abc-DEF_12345"))

    def test_wrong_or_missing_secret_is_rejected(self):
        self.assertFalse(verify_webhook_secret("abc-DEF_12345", "wrong"))
        self.assertFalse(verify_webhook_secret("abc-DEF_12345", None))
        self.assertFalse(verify_webhook_secret("", "anything"))

    def test_webhook_event_key_uses_original_forwarded_message_mid(self):
        update = {
            "update_type": "message_created",
            "timestamp": 123456,
            "message": {
                "link": {"type": "forward", "message": {"mid": "forwarded-mid"}},
                "body": {},
            },
        }
        self.assertEqual(
            max_bot.ScheduleMaxBot.event_key(update),
            "message_created:forwarded-mid:123456:",
        )
