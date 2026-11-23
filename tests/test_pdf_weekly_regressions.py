from __future__ import annotations
import unittest
from datetime import date
from pathlib import Path

from pdf_parser import load_subject_dictionary, parse_pdf
from pdf_schedule import pdf_to_schedule
from parser import parse_xlsx
from schedule_manager import merge_schedule_data

ROOT = Path(__file__).resolve().parents[1]
PDF = ROOT / "tests" / "fixtures" / "weekly_16_20_11_2026.pdf"
REFERENCE = ROOT / "assets" / "reference" / "Расписание_16.11.2026-20.11.2026.xlsx"


class WeeklyPdfRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dictionary = load_subject_dictionary(ROOT / "subjects_dictionary.json")
        cls.pdf = parse_pdf(PDF, cls.dictionary)
        cls.xlsx = parse_xlsx(REFERENCE)
        cls.converted, cls.skipped = pdf_to_schedule(cls.pdf, reference_path=REFERENCE)

    def test_detects_all_weekdays_even_when_first_visible_pair_is_two(self):
        self.assertEqual(
            self.pdf["dates"],
            ["2026-11-16", "2026-11-17", "2026-11-18", "2026-11-19", "2026-11-20"],
        )
        self.assertEqual(self.pdf["group_count"], 75)

    def test_group_315_pair_numbers_match_xlsx_for_each_day(self):
        for day, xlsx_lessons in self.xlsx["groups"]["315"]["days"].items():
            pdf_lessons = self.converted["groups"]["315"]["days"][day]
            pdf_pairs = sorted(item["pair"] for item in pdf_lessons if item["pair"] > 0)
            xlsx_pairs = sorted(item["pair"] for item in xlsx_lessons if item["pair"] > 0)
            self.assertEqual(pdf_pairs, xlsx_pairs, msg=f"group 315, {day}")

    def test_group_315_thursday_starts_at_pair_three_and_is_not_renumbered(self):
        thursday = "2026-11-19"
        pdf_lessons = self.converted["groups"]["315"]["days"][thursday]
        xlsx_lessons = self.xlsx["groups"]["315"]["days"][thursday]
        pdf_pairs = sorted(item["pair"] for item in pdf_lessons if item["pair"] > 0)
        xlsx_pairs = sorted(item["pair"] for item in xlsx_lessons if item["pair"] > 0)
        self.assertEqual(pdf_pairs, [3, 4, 5, 7])
        self.assertEqual(pdf_pairs, xlsx_pairs)

    def test_group_315_thursday_and_friday_are_separate_blocks(self):
        thursday = "2026-11-19"
        friday = "2026-11-20"
        pdf_thursday = sorted(
            item["pair"] for item in self.converted["groups"]["315"]["days"][thursday]
            if item["pair"] > 0
        )
        pdf_friday = sorted(
            item["pair"] for item in self.converted["groups"]["315"]["days"][friday]
            if item["pair"] > 0
        )
        xlsx_thursday = sorted(
            item["pair"] for item in self.xlsx["groups"]["315"]["days"][thursday]
            if item["pair"] > 0
        )
        xlsx_friday = sorted(
            item["pair"] for item in self.xlsx["groups"]["315"]["days"][friday]
            if item["pair"] > 0
        )
        self.assertEqual(pdf_thursday, xlsx_thursday)
        self.assertEqual(pdf_friday, xlsx_friday)
        self.assertNotEqual(pdf_thursday, pdf_friday)

    def test_weekly_pdf_keeps_explicitly_empty_group_days(self):
        # Group 452 has no classes Monday in this file. The empty day must be
        # retained so it can clear stale data from an earlier cached schedule.
        self.assertIn("2026-11-16", self.converted["groups"]["452"]["days"])
        self.assertEqual(self.converted["groups"]["452"]["days"]["2026-11-16"], [])

    def test_empty_pdf_day_clears_previous_cached_lessons(self):
        current = {
            "groups": {
                "452": {
                    "sheet": "old cache",
                    "days": {
                        "2026-11-16": [{"pair": 1, "subject": "Старое занятие", "subgroup": 1}],
                    },
                }
            }
        }
        merged, _ = merge_schedule_data(
            current,
            self.converted,
            today=date(2026, 11, 19),
            source="site_pdf",
        )
        self.assertEqual(merged["groups"]["452"]["days"]["2026-11-16"], [])

    def test_all_five_days_exist_for_every_group_detected_on_pdf_pages(self):
        for page in self.pdf["pages"]:
            for group in page.get("groups", []):
                self.assertEqual(
                    [day["date"] for day in group["days"]],
                    ["2026-11-16", "2026-11-17", "2026-11-18", "2026-11-19", "2026-11-20"],
                )


if __name__ == "__main__":
    unittest.main()
