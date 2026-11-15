"""Integration tests against the real MIC files (source.xlsx / history_source.xlsx).

Run: python3 test_real_data.py
Expected values were measured from the 令和6年1月1日更新版 on 2026-11-15.
"""
import unittest
from collections import Counter

from build_codes import (
    build_records,
    build_history_events,
    derive_lifespans,
    validate_records,
)


class TestCurrentData(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.records = build_records("source.xlsx")

    def test_total_count(self):
        self.assertEqual(len(self.records), 1965)

    def test_type_counts(self):
        counts = Counter(r["type"] for r in self.records)
        self.assertEqual(counts, {
            "prefecture": 47,
            "designated_city": 20,
            "ward": 171,
            "special_ward": 23,
            "city": 772,
            "town": 743,
            "village": 189,
        })

    def test_all_check_digits_valid(self):
        validate_records(self.records)  # must not raise

    def test_known_records(self):
        by_code = {r["code"]: r for r in self.records}
        self.assertEqual(by_code["010006"]["prefecture_name"], "北海道")
        self.assertIsNone(by_code["010006"]["municipality_name"])
        self.assertEqual(by_code["011002"]["municipality_name"], "札幌市")
        self.assertEqual(by_code["011002"]["type"], "designated_city")
        self.assertEqual(by_code["131016"]["municipality_name"], "千代田区")
        self.assertEqual(by_code["131016"]["type"], "special_ward")
        self.assertEqual(by_code["131016"]["municipality_kana"], "チヨダク")
        self.assertEqual(by_code["131016"]["municipality_kana_halfwidth"], "ﾁﾖﾀﾞｸ")
        self.assertEqual(by_code["131016"]["code5"], "13101")

    def test_designated_cities_deduplicated(self):
        codes = [r["code"] for r in self.records]
        self.assertEqual(len(codes), len(set(codes)))


class TestHistoryData(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.events = build_history_events("history_source.xlsx")
        cls.records = build_records("source.xlsx")
        cls.lifespans = derive_lifespans(cls.records, cls.events)

    def test_block_count(self):
        groups = {e["group_id"] for e in self.events}
        self.assertEqual(len(groups), 417)

    def test_all_events_dated(self):
        for e in self.events:
            self.assertIsNotNone(e["date"], e)

    def test_event_date_range(self):
        dates = sorted(e["date"] for e in self.events)
        self.assertEqual(dates[0], "2005-04-01")
        self.assertEqual(dates[-1], "2024-01-01")

    def test_change_types(self):
        types = {e["change_type"] for e in self.events}
        self.assertEqual(types, {"created", "retired", "continued"})

    def test_nagaoka_deletion_rows_are_retired(self):
        # The 2006-01-01 Nagaoka merger rows carry kind 削除 (not 欠番);
        # all four must be classified retired in history_events itself.
        ev = [e for e in self.events
              if e["date"] == "2006-01-01"
              and e["code_before"] in ("152153", "154032", "154041", "154067")]
        self.assertEqual(len(ev), 4)
        for e in ev:
            self.assertEqual(e["change_type"], "retired", e)

    def test_hamamatsu_ward_reorg(self):
        ev = [e for e in self.events
              if e["date"] == "2024-01-01" and e["code_after"] == "221384"]
        self.assertEqual(len(ev), 1)
        self.assertEqual(ev[0]["change_type"], "created")
        self.assertEqual(ev[0]["name_after"], "中央区")

    def test_shinoyama_rename(self):
        ev = [e for e in self.events if e["code_before"] == "282219"]
        self.assertEqual(len(ev), 1)
        self.assertEqual(ev[0]["change_type"], "continued")
        self.assertEqual(ev[0]["code_after"], "282219")
        self.assertEqual(ev[0]["name_after"], "丹波篠山市")
        self.assertEqual(ev[0]["date"], "2019-05-01")

    def test_lifespan_consistency(self):
        current = {r["code"] for r in self.records}
        by_code = {s["code"]: s for s in self.lifespans}
        for code in current:
            self.assertIn(code, by_code)
            self.assertIsNone(by_code[code]["valid_to"], code)
        for code, span in by_code.items():
            if span["valid_to"] is not None:
                self.assertNotIn(code, current)
            if span["valid_from"] and span["valid_to"]:
                # SPEC: same-day valid_from == valid_to is allowed
                self.assertLessEqual(span["valid_from"], span["valid_to"])

    def test_retired_code_lifespan(self):
        by_code = {s["code"]: s for s in self.lifespans}
        # 浅羽町 224812 was abolished in the 2005-04-01 merger into 袋井市
        self.assertEqual(by_code["224812"]["valid_to"], "2005-04-01")
        self.assertIsNone(by_code["224812"]["valid_from"])  # predates coverage

    def test_abolished_codes_not_missing_from_lifespans(self):
        # Regression (2026-11-16): codes abolished via kind 削除 rows
        # (長岡市 merger, 2006-01-01) and via continued merger-mapping rows
        # (佐久市 merger, 2005-04-01) used to vanish from lifespans entirely.
        by_code = {s["code"]: s for s in self.lifespans}
        for code in ("152153", "154032", "154041", "154067"):
            self.assertEqual(by_code[code]["valid_to"], "2006-01-01", code)
            self.assertIsNone(by_code[code]["valid_from"], code)
        for code in ("203017", "203220"):
            self.assertEqual(by_code[code]["valid_to"], "2005-04-01", code)
            self.assertIsNone(by_code[code]["valid_from"], code)

    def test_validate_lifespans_passes_on_real_data(self):
        from build_codes import validate_lifespans
        validate_lifespans(self.records, self.events, self.lifespans)


class TestSchema(unittest.TestCase):
    def test_shipped_outputs_validate_against_schema(self):
        import json

        from build_codes import _build_schema

        with open("schema.json", encoding="utf-8") as f:
            schema = json.load(f)
        self.assertEqual(schema, _build_schema())
        with open("local_gov_codes.json", encoding="utf-8") as f:
            data = json.load(f)
        try:
            import jsonschema
        except ImportError:
            jsonschema = None
        if jsonschema is not None:
            jsonschema.validate(data, schema)
        else:
            # Structural fallback: array root, required keys, code pattern.
            import re

            self.assertIsInstance(data, list)
            for rec in data:
                for key in schema["items"]["required"]:
                    self.assertIn(key, rec)
                self.assertRegex(rec["code"], r"^\d{6}$")


if __name__ == "__main__":
    unittest.main(verbosity=2)
