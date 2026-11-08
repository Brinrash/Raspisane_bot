import unittest
from pathlib import Path

from parser import parse_xlsx

ROOT = Path(__file__).resolve().parents[1]


class TestScheduleParser(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data = parse_xlsx(ROOT / "assets" / "schedule_seed.xlsx")

    def test_detects_groups_and_week(self):
        self.assertGreaterEqual(self.data["group_count"], 60)
        self.assertEqual(self.data["date_from"], "2026-11-09")
        self.assertEqual(self.data["date_to"], "2026-11-13")
        self.assertIn("315", self.data["groups"])

    def test_group_315_monday(self):
        lessons = self.data["groups"]["315"]["days"]["2026-11-09"]
        subjects = [item["subject"] for item in lessons]
        self.assertIn("Классный час", subjects)
        self.assertIn("МДК 03.03 (1 п/г)", subjects)
        self.assertIn("Физическая культура (1 п/г)", subjects)
        self.assertIn("МДК 03.01 (1 п/г)", subjects)

    def test_group_315_friday(self):
        lessons = self.data["groups"]["315"]["days"]["2026-11-13"]
        pairs = [item["pair"] for item in lessons]
        self.assertEqual(pairs, [1, 2, 3, 4])
        self.assertEqual(lessons[0]["room"], "2/18")
        self.assertEqual(lessons[0]["teacher"], "Пунгин И.В.")

    def test_subgroup_is_preserved(self):
        wednesday = self.data["groups"]["315"]["days"]["2026-11-11"]
        legal = next(item for item in wednesday if "Правовое обеспечение" in item["subject"])
        self.assertIn("2 п/г", legal["subject"])
        self.assertEqual(legal["room"], "129")


if __name__ == "__main__":
    unittest.main()
