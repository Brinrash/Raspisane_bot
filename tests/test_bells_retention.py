import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from bells import (
    DEFAULT_BELLS, bell_for_date, format_bells, merge_bell_override,
    pair_time, parse_bell_document, prune_bell_data,
)
from max_bot import lesson_line


class BellScheduleTests(unittest.TestCase):
    def test_default_monday_pair_maps_to_two_lessons(self):
        monday = date(2026, 11, 9)
        self.assertEqual(pair_time(monday, 1), "09:20–10:05 / 10:15–11:00")
        self.assertEqual(pair_time(monday, 2), "11:10–11:55 / 12:35–13:20")
        self.assertEqual(pair_time(monday, 4), "15:30–16:15 / 16:25–17:10")

    def test_default_tuesday_to_friday_pair_maps_to_two_lessons(self):
        tuesday = date(2026, 11, 10)
        self.assertEqual(pair_time(tuesday, 1), "08:10–08:55 / 09:15–10:00")
        self.assertEqual(pair_time(tuesday, 6), "18:10–18:55 / 19:05–19:50")

    def test_lesson_line_includes_time(self):
        text = lesson_line({"pair": 1, "subject": "МДК 03.03", "room": "2/17"}, date(2026, 11, 9))
        self.assertIn("МДК 03.03", text)
        self.assertIn("09:20–10:05 / 10:15–11:00", text)
        self.assertIn("Каб.: 2/17", text)

    def test_date_override_changes_only_selected_slot(self):
        day = date(2026, 11, 10)
        data = merge_bell_override({"overrides": {}}, day, {"lessons": {"1": {"start": "08:30", "end": "09:15"}}}, "bells.docx")
        self.assertEqual(pair_time(day, 1, data), "08:30–09:15 / 09:15–10:00")
        self.assertEqual(pair_time(day, 2, data), "10:10–10:55 / 11:05–11:50")
        self.assertIn("изменение звонков", format_bells(day, data))

    def test_full_bell_timetable_includes_breaks_and_practice(self):
        monday_text = format_bells(date(2026, 11, 9))
        weekday_text = format_bells(date(2026, 11, 10))
        for value in ("Классный час", "Завтрак", "Обед", "Полдник", "УЧЕБНАЯ ПРАКТИКА", "1 смена", "2 смена"):
            self.assertIn(value, monday_text)
        for value in ("Десятиминутка", "Завтрак", "Обед", "Полдник", "УЧЕБНАЯ ПРАКТИКА", "1 смена", "2 смена"):
            self.assertIn(value, weekday_text)
        self.assertIn("1-й урок — 09:20–10:05", monday_text)
        self.assertIn("2-й урок — 10:15–11:00", monday_text)
        self.assertIn("3-й урок — 11:10–11:55", monday_text)
        self.assertNotIn("первая половина", monday_text)
        self.assertNotIn("вторая половина", monday_text)
        self.assertNotIn("пара (", monday_text)

    def test_docx_parser_reads_all_rows_and_practice_from_documents(self):
        from bells import _parse_docx
        monday_path = Path("/mnt/data/Распорядок учебных занятий на понедельник(1)(1).docx")
        weekday_path = Path("/mnt/data/Распорядок учебных занятий со вторника по пятницу(1)(1).docx")
        if monday_path.exists() and weekday_path.exists():
            monday, _ = _parse_docx(monday_path)
            weekday, _ = _parse_docx(weekday_path)
            self.assertEqual(set(monday["lessons"]), {str(i) for i in range(1, 9)})
            self.assertEqual(set(weekday["lessons"]), {str(i) for i in range(1, 13)})
            self.assertEqual([b["after_lesson"] for b in monday["breaks"]], [0, 3, 6])
            self.assertEqual([b["after_lesson"] for b in weekday["breaks"]], [1, 4, 7])
            self.assertEqual(len(monday["practice_shifts"]), 2)
            self.assertEqual(len(weekday["practice_shifts"]), 2)

    def test_bell_overrides_follow_three_week_retention(self):
        data = {"overrides": {
            "2026-10-25": {"lessons": {}},
            "2026-10-26": {"lessons": {}},
            "2026-11-15": {"lessons": {}},
            "2026-11-16": {"lessons": {}},
        }}
        pruned = prune_bell_data(data, date(2026, 11, 7))
        self.assertEqual(set(pruned["overrides"]), {"2026-10-26", "2026-11-15"})


if __name__ == "__main__":
    unittest.main()
