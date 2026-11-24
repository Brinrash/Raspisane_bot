import unittest
from datetime import date
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from group_picker import course_for_group, groups_by_course
from schedule_manager import merge_schedule_data, prune_schedule_data, retention_bounds


class TestGroupPicker(unittest.TestCase):
    def test_course_from_group_code(self):
        self.assertEqual(course_for_group("315"), 3)
        self.assertEqual(course_for_group("201"), 2)
        self.assertEqual(course_for_group("406"), 4)
        self.assertEqual(course_for_group("521"), 5)
        self.assertIsNone(course_for_group("AB1"))

    def test_groups_are_bucketed_by_course(self):
        result = groups_by_course({"315": {}, "101": {}, "201": {}, "521": {}})
        self.assertEqual(result[1], ["101"])
        self.assertEqual(result[2], ["201"])
        self.assertEqual(result[3], ["315"])
        self.assertEqual(result[5], ["521"])


class TestScheduleMerge(unittest.TestCase):
    def test_daily_update_does_not_erase_other_days_or_groups(self):
        current = {
            "groups": {
                "315": {"days": {
                    "2026-11-09": [{"pair": 1, "subject": "Old Monday"}],
                    "2026-11-10": [{"pair": 1, "subject": "Tuesday"}],
                }},
                "316": {"days": {
                    "2026-11-09": [{"pair": 1, "subject": "Other group"}],
                }},
            },
            "table_count": 3,
        }
        incoming = {
            "groups": {"315": {"days": {
                "2026-11-09": [{"pair": 1, "subject": "Replacement Monday"}],
            }}},
            "dates": ["2026-11-09"],
            "date_from": "2026-11-09",
            "date_to": "2026-11-09",
            "table_count": 1,
        }
        result, diff = merge_schedule_data(current, incoming, today=date(2026, 11, 7), source="MAX")
        self.assertEqual(result["groups"]["315"]["days"]["2026-11-09"][0]["subject"], "Replacement Monday")
        self.assertEqual(result["groups"]["315"]["days"]["2026-11-10"][0]["subject"], "Tuesday")
        self.assertIn("316", result["groups"])
        self.assertIn("315", diff["changed_groups"])
        self.assertEqual(diff["changed_dates"], ["2026-11-09"])

    def test_retains_previous_current_and_next_calendar_weeks(self):
        current = {"groups": {"315": {"days": {
            "2026-10-25": [{"pair": 1, "subject": "Too old"}],
            "2026-10-26": [{"pair": 1, "subject": "Previous week"}],
            "2026-11-07": [{"pair": 1, "subject": "Current week"}],
            "2026-11-15": [{"pair": 1, "subject": "Next week"}],
            "2026-11-16": [{"pair": 1, "subject": "Too far"}],
        }}}}
        result, info = prune_schedule_data(current, date(2026, 11, 7))
        days = result["groups"]["315"]["days"]
        self.assertNotIn("2026-10-25", days)
        self.assertIn("2026-10-26", days)
        self.assertIn("2026-11-07", days)
        self.assertIn("2026-11-15", days)
        self.assertNotIn("2026-11-16", days)
        self.assertEqual(info["retention_start"], "2026-10-26")
        self.assertEqual(info["retention_end"], "2026-11-15")

    def test_retention_bounds_are_three_monday_to_sunday_weeks(self):
        self.assertEqual(retention_bounds(date(2026, 11, 7)), (date(2026, 10, 26), date(2026, 11, 15)))

    def test_incoming_outside_window_is_ignored_without_erasing_valid_data(self):
        current = {"groups": {"315": {"days": {"2026-11-09": [{"pair": 1, "subject": "Current"}]}}}}
        incoming = {"groups": {"315": {"days": {"2026-10-19": [{"pair": 1, "subject": "Too old"}]}}}, "dates": ["2026-10-19"], "table_count": 1}
        result, diff = merge_schedule_data(current, incoming, today=date(2026, 11, 7))
        self.assertEqual(result["groups"]["315"]["days"]["2026-11-09"][0]["subject"], "Current")
        self.assertNotIn("2026-10-19", result["groups"]["315"]["days"])
        self.assertEqual(diff["changed_groups"], [])



if __name__ == "__main__":
    unittest.main()
