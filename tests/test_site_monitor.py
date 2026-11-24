from __future__ import annotations

import unittest
from datetime import datetime
from zoneinfo import ZoneInfo

from site_monitor import next_check_at, invalidate_site_cache, load_site_state, save_site_state, prune_site_archive
from pdf_schedule import pdf_to_schedule
from group_picker import home_keyboard, more_keyboard
from schedule_manager import merge_schedule_data


class SiteScheduleTests(unittest.TestCase):
    def test_schedule_window_boundaries(self):
        tz = ZoneInfo("Europe/Moscow")
        self.assertEqual(
            next_check_at(datetime(2026, 11, 19, 11, 55, tzinfo=tz)),
            datetime(2026, 11, 19, 12, 0, tzinfo=tz),
        )
        self.assertEqual(
            next_check_at(datetime(2026, 11, 19, 12, 0, tzinfo=tz)),
            datetime(2026, 11, 19, 12, 0, tzinfo=tz),
        )
        self.assertEqual(
            next_check_at(datetime(2026, 11, 19, 12, 1, tzinfo=tz)),
            datetime(2026, 11, 19, 12, 15, tzinfo=tz),
        )
        self.assertEqual(
            next_check_at(datetime(2026, 11, 19, 23, 44, tzinfo=tz)),
            datetime(2026, 11, 19, 23, 45, tzinfo=tz),
        )
        self.assertEqual(
            next_check_at(datetime(2026, 11, 19, 23, 46, tzinfo=tz)),
            datetime(2026, 11, 20, 12, 0, tzinfo=tz),
        )

    def test_site_pdf_archive_is_bounded(self):
        import tempfile
        from pathlib import Path
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            for index in range(5):
                (folder / f"{index}.pdf").write_bytes(b"%PDF-test")
            removed = prune_site_archive(folder, max_files=3)
            self.assertEqual(len(removed), 2)
            self.assertEqual(len(list(folder.glob("*.pdf"))), 3)

    def test_new_xlsx_reference_invalidates_processed_pdf_cache(self):
        import tempfile
        from pathlib import Path
        with tempfile.TemporaryDirectory() as directory:
            state_path = Path(directory) / "site_sync_state.json"
            save_site_state(state_path, {"files": {
                "https://example.test/weekly.pdf": {"sha256": "abc", "status": "processed"},
                "https://example.test/broken.pdf": {"sha256": "def", "status": "failed"},
                "https://example.test/nohash.pdf": {"status": "failed"},
            }})
            self.assertEqual(invalidate_site_cache(state_path), 2)
            state = load_site_state(state_path)
            self.assertEqual(state["files"]["https://example.test/weekly.pdf"]["status"], "reference_changed")
            self.assertEqual(state["files"]["https://example.test/broken.pdf"]["status"], "reference_changed")
            self.assertEqual(state["files"]["https://example.test/nohash.pdf"]["status"], "failed")

    def test_home_menu_hides_secondary_actions(self):
        payloads = {button["payload"] for row in home_keyboard() for button in row}
        self.assertIn("menu:more", payloads)
        self.assertNotIn("menu:resetgroup", payloads)
        self.assertNotIn("menu:help", payloads)
        self.assertNotIn("menu:bells", payloads)
        self.assertNotIn("menu:mygroup", payloads)

    def test_more_menu_contains_secondary_actions(self):
        payloads = {button["payload"] for row in more_keyboard() for button in row}
        self.assertTrue({"menu:resetgroup", "menu:help", "menu:bells", "menu:mygroup", "menu:home"}.issubset(payloads))

    def test_weekly_pdf_fills_only_missing_or_uncertain_blocks_from_exact_week_xlsx(self):
        from unittest.mock import patch
        parsed = {
            "document_type": "weekly",
            "source_file": "schedule.pdf",
            "date_from": "2026-11-16",
            "date_to": "2026-11-20",
            "dates": ["2026-11-18"],
            "entries": [
                {
                    "group": "170", "date": "2026-11-18", "pair": 1,
                    "lesson_numbers": [1], "subject": "Математика", "room": "2/41",
                    "teacher": "Преподаватель А.А.", "subgroup": 1, "needs_review": False,
                },
                {
                    "group": "171", "date": "2026-11-18", "pair": 1,
                    "lesson_numbers": [1], "subject": "Обрезанное", "room": "2/36",
                    "teacher": "Преподаватель Б.Б.", "subgroup": 1, "needs_review": True,
                },
            ],
        }
        reference = {
            "date_from": "2026-11-16", "date_to": "2026-11-20",
            "groups": {
                "170": {"days": {"2026-11-18": [{"pair": 1, "subject": "НЕ ПОДМЕНЯТЬ", "room": "2/41", "teacher": "Преподаватель А.А.", "subgroup": 1}]}},
                "171": {"days": {"2026-11-18": [{"pair": 1, "subject": "История", "room": "2/36", "teacher": "Преподаватель Б.Б.", "subgroup": 1}]}},
                "172": {"days": {"2026-11-18": [{"pair": 2, "subject": "Физика", "room": "117", "teacher": "Преподаватель В.В.", "subgroup": 1}]}},
            },
        }
        with patch("pdf_schedule.parse_xlsx", return_value=reference):
            converted, skipped = pdf_to_schedule(parsed, reference_path="reference.xlsx")
        lessons_170 = converted["groups"]["170"]["days"]["2026-11-18"]
        lessons_171 = converted["groups"]["171"]["days"]["2026-11-18"]
        self.assertEqual(lessons_170[0]["subject"], "Математика")
        self.assertEqual(lessons_171[0]["subject"], "История")
        self.assertEqual(lessons_171[0]["quality_status"], "same_week_xlsx_fallback")
        self.assertIn("172", converted["groups"])
        self.assertEqual(len(converted["reference_fallback_blocks"]), 2)
        self.assertEqual(skipped, [])

    def test_daily_replacement_never_uses_weekly_xlsx_as_fallback(self):
        from unittest.mock import patch
        parsed = {
            "document_type": "replacements", "replacement_mode": "replace",
            "source_file": "Zamena.pdf", "date_from": "2026-11-18", "date_to": "2026-11-18",
            "dates": ["2026-11-18"],
            "entries": [{
                "group": "171", "date": "2026-11-18", "pair": 1,
                "lesson_numbers": [1], "subject": "Обрезанное", "room": "2/36",
                "teacher": "Преподаватель Б.Б.", "subgroup": 1, "needs_review": True,
            }],
        }
        with patch("pdf_schedule.parse_xlsx") as parser_mock:
            with self.assertRaises(ValueError):
                pdf_to_schedule(parsed, reference_path="reference.xlsx")
            parser_mock.assert_not_called()

    def test_pdf_converter_keeps_good_blocks_and_holds_uncertain_group_day(self):
        parsed = {
            "document_type": "replacements",
            "replacement_mode": "replace",
            "source_file": "Zamena_zanyatiy_na_19.11.2026.pdf",
            "page_count": 1,
            "dates": ["2026-11-19"],
            "entries": [
                {
                    "group": "170", "date": "2026-11-19", "pair": 1,
                    "lesson_numbers": [1, 2], "subject": "Математика", "room": "2/41",
                    "teacher": "Майбурова Г.Н.", "subgroup": 1, "self_study": False,
                    "needs_review": False,
                },
                {
                    "group": "171", "date": "2026-11-19", "pair": 1,
                    "lesson_numbers": [1, 2], "subject": "Основы гидравлик", "room": "2/36",
                    "teacher": "Королев С.Г.", "subgroup": None, "self_study": False,
                    "needs_review": True,
                },
            ],
        }
        converted, skipped = pdf_to_schedule(parsed)
        self.assertIn("170", converted["groups"])
        self.assertNotIn("171", converted["groups"])
        self.assertEqual(converted["groups"]["170"]["days"]["2026-11-19"][0]["subject"], "Математика")
        self.assertEqual(skipped, [{"group": "171", "date": "2026-11-19", "reason": "сомнительных/пустых записей: 1"}])

    def test_supplement_pdf_overrides_existing_pair_subgroup_without_lesson_number(self):
        current = {
            "groups": {"315": {"days": {"2026-11-19": [
                {"pair": 1, "subject": "Старый предмет", "subgroup": 1},
                {"pair": 2, "subject": "Оставить", "subgroup": 1},
            ]}}},
        }
        incoming = {
            "groups": {"315": {"days": {"2026-11-19": [
                {"pair": 1, "subject": "Новый предмет", "subgroup": 1},
            ]}}},
            "dates": ["2026-11-19"],
            "replacement_mode": "supplement",
        }
        merged, _ = merge_schedule_data(current, incoming, today=datetime(2026, 11, 19).date())
        lessons = merged["groups"]["315"]["days"]["2026-11-19"]
        self.assertEqual(len(lessons), 2)
        self.assertEqual(lessons[0]["subject"], "Новый предмет")
        self.assertEqual(lessons[1]["subject"], "Оставить")


if __name__ == "__main__":
    unittest.main()
