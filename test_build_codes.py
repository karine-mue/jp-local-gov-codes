"""Tests for build_codes.py. Run: python3 test_build_codes.py"""
import datetime
import os
import tempfile
import unittest

import openpyxl

from build_codes import (
    _is_data_code_cell,
    build_history_events,
    build_records,
    check_digit,
    classify_type,
    derive_lifespans,
    parse_date_cell,
    to_fullwidth_kana,
    validate_lifespans,
    validate_records,
)


class TestCheckDigit(unittest.TestCase):
    def test_known_codes(self):
        self.assertEqual(check_digit("13101"), 6)   # 千代田区 131016
        self.assertEqual(check_digit("01100"), 2)   # 札幌市 011002
        self.assertEqual(check_digit("01000"), 6)   # 北海道 010006 (S=5 < 11)

    def test_remainder_zero_gives_one(self):
        self.assertEqual(check_digit("01213"), 1)   # 012131, S=22, S mod 11 = 0

    def test_remainder_one_gives_zero(self):
        self.assertEqual(check_digit("01205"), 0)   # 012050, S=23, S mod 11 = 1

    def test_out_of_range_raises(self):
        # S=0 gives digit 11, which is out of range
        with self.assertRaises(ValueError):
            check_digit("00000")

    def test_fullwidth_digits_rejected(self):
        with self.assertRaises(ValueError):
            check_digit("１２３４５")


class TestParseDateCell(unittest.TestCase):
    def test_datetime(self):
        self.assertEqual(parse_date_cell(datetime.datetime(2005, 4, 1)),
                         datetime.date(2005, 4, 1))

    def test_serial_int(self):
        self.assertEqual(parse_date_cell(40260), datetime.date(2010, 3, 23))

    def test_era_strings(self):
        self.assertEqual(parse_date_cell("R6.1.1"), datetime.date(2024, 1, 1))
        self.assertEqual(parse_date_cell("H30.10.1"), datetime.date(2018, 10, 1))
        self.assertEqual(parse_date_cell("R1.5.1"), datetime.date(2019, 5, 1))

    def test_ditto_and_none(self):
        self.assertIsNone(parse_date_cell("〃"))
        self.assertIsNone(parse_date_cell(None))

    def test_empty_string_is_none(self):
        # SPEC: blank date cells are filled from the block date
        self.assertIsNone(parse_date_cell(""))
        self.assertIsNone(parse_date_cell("  "))


class TestKana(unittest.TestCase):
    def test_halfwidth_to_fullwidth(self):
        self.assertEqual(to_fullwidth_kana("ｻｯﾎﾟﾛｼ"), "サッポロシ")
        self.assertEqual(to_fullwidth_kana("ﾎｯｶｲﾄﾞｳ"), "ホッカイドウ")
        self.assertEqual(to_fullwidth_kana(None), None)


class TestClassifyType(unittest.TestCase):
    def test_prefecture(self):
        self.assertEqual(classify_type("北海道", None, in_designated_sheet=False),
                         "prefecture")

    def test_designated_city_and_ward(self):
        self.assertEqual(classify_type("北海道", "札幌市", in_designated_sheet=True),
                         "designated_city")
        self.assertEqual(classify_type("北海道", "札幌市中央区", in_designated_sheet=True),
                         "ward")

    def test_special_ward(self):
        self.assertEqual(classify_type("東京都", "千代田区", in_designated_sheet=False),
                         "special_ward")

    def test_city_town_village(self):
        self.assertEqual(classify_type("北海道", "函館市", in_designated_sheet=False), "city")
        self.assertEqual(classify_type("埼玉県", "寄居町", in_designated_sheet=False), "town")
        self.assertEqual(classify_type("埼玉県", "横瀬町", in_designated_sheet=False), "town")
        self.assertEqual(classify_type("長野県", "川上村", in_designated_sheet=False), "village")


class TestValidateRecords(unittest.TestCase):
    def test_out_of_range_code_is_aggregated(self):
        # check_digit("00000") raises internally; validate_records must
        # report the code in its aggregated error, not abort raw.
        with self.assertRaises(ValueError) as ctx:
            validate_records([{"code": "000000"}])
        self.assertIn("Invalid check digits for codes: 000000", str(ctx.exception))


class TestDataCodeCell(unittest.TestCase):
    def test_ascii_and_markers(self):
        self.assertTrue(_is_data_code_cell("012345"))
        self.assertTrue(_is_data_code_cell("〃"))
        self.assertTrue(_is_data_code_cell("同左"))
        self.assertTrue(_is_data_code_cell("削除"))

    def test_fullwidth_rejected(self):
        self.assertFalse(_is_data_code_cell("１２３４５６"))

    def test_none_and_short(self):
        self.assertFalse(_is_data_code_cell(None))
        self.assertFalse(_is_data_code_cell("12345"))


def _save_wb(wb, tmpdir, name):
    path = os.path.join(tmpdir, name)
    wb.save(path)
    return path


def _current_wb(main_rows, designated_rows, extra_sheet=False):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "現在の団体"
    ws.append(["コード", "都道府県名", "市区町村名", "都道府県カナ", "市区町村カナ"])
    for r in main_rows:
        ws.append(list(r))
    ws2 = wb.create_sheet("政令指定都市")
    ws2.append(["コード", "都道府県名", "市区町村名", "都道府県カナ", "市区町村カナ"])
    for r in designated_rows:
        ws2.append(list(r))
    if extra_sheet:
        wb.create_sheet("予備")
    return wb


class TestBuildRecordsFixtures(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def test_minimal_valid(self):
        wb = _current_wb(
            [("010006", "北海道", None, "ﾎｯｶｲﾄﾞｳ", None),
             ("011002", "北海道", "札幌市", "ﾎｯｶｲﾄﾞｳ", "ｻｯﾎﾟﾛｼ")],
            [("011002", "北海道", "札幌市", "ﾎｯｶｲﾄﾞｳ", "ｻｯﾎﾟﾛｼ")],
        )
        records = build_records(_save_wb(wb, self.tmp.name, "c.xlsx"))
        self.assertEqual(len(records), 2)
        by_code = {r["code"]: r for r in records}
        self.assertEqual(by_code["011002"]["type"], "designated_city")

    def test_three_sheets_rejected(self):
        wb = _current_wb(
            [("010006", "北海道", None, "ﾎｯｶｲﾄﾞｳ", None)],
            [],
            extra_sheet=True,
        )
        with self.assertRaises(ValueError):
            build_records(_save_wb(wb, self.tmp.name, "c.xlsx"))

    def test_duplicate_code_in_main_sheet_rejected(self):
        wb = _current_wb(
            [("010006", "北海道", None, "ﾎｯｶｲﾄﾞｳ", None),
             ("010006", "北海道", None, "ﾎｯｶｲﾄﾞｳ", None)],
            [],
        )
        with self.assertRaises(ValueError):
            build_records(_save_wb(wb, self.tmp.name, "c.xlsx"))

    def test_designated_mismatch_rejected(self):
        wb = _current_wb(
            [("011002", "北海道", "札幌市", "ﾎｯｶｲﾄﾞｳ", "ｻｯﾎﾟﾛｼ")],
            [("011002", "北海道", "札幌市X", "ﾎｯｶｲﾄﾞｳ", "ｻｯﾎﾟﾛｼ")],
        )
        with self.assertRaises(ValueError):
            build_records(_save_wb(wb, self.tmp.name, "c.xlsx"))

    def test_invalid_code_cell_rejected(self):
        wb = _current_wb(
            [("ABC", "北海道", None, "ﾎｯｶｲﾄﾞｳ", None)],
            [],
        )
        with self.assertRaises(ValueError):
            build_records(_save_wb(wb, self.tmp.name, "c.xlsx"))

    def test_fullwidth_code_cell_rejected(self):
        wb = _current_wb(
            [("０１０００６", "北海道", None, "ﾎｯｶｲﾄﾞｳ", None)],
            [],
        )
        with self.assertRaises(ValueError):
            build_records(_save_wb(wb, self.tmp.name, "c.xlsx"))


def _history_wb(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "改正一覧表"
    ws.append(["h"] * 14)
    for r in rows:
        ws.append(list(r))
    return wb


def _hrow(pref=None, cb=None, nb=None, kind=None, ca=None, na=None,
          reason=None, date=None):
    row = [None] * 14
    row[4] = pref
    row[5] = cb
    row[6] = nb
    row[8] = kind
    row[9] = ca
    row[10] = na
    row[12] = reason
    row[13] = date
    return row


class TestBuildHistoryFixtures(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)

    def _events(self, rows):
        wb = _history_wb(rows)
        return build_history_events(_save_wb(wb, self.tmp.name, "h.xlsx"))

    def test_numeric_code_cells_zero_padded(self):
        events = self._events([
            _hrow(pref="北海道", cb=10006, nb="北海道", kind="欠番",
                  date=datetime.datetime(2010, 1, 1)),
        ])
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0]["code_before"], "010006")
        self.assertEqual(events[0]["change_type"], "retired")

    def test_unknown_kind_rejected(self):
        with self.assertRaises(ValueError):
            self._events([
                _hrow(pref="北海道", cb="010006", nb="北海道", kind="変種",
                      date=datetime.datetime(2010, 1, 1)),
            ])

    def test_reason_ditto_resolved(self):
        events = self._events([
            _hrow(pref="青森県", kind="新設", ca="026807", na="某市",
                  reason="合併", date=datetime.datetime(2010, 4, 1)),
            _hrow(cb="026801", nb="某町", kind="欠番", reason="〃",
                  date=datetime.datetime(2010, 4, 1)),
        ])
        self.assertEqual(events[1]["reason"], "合併")

    def test_blank_kind_with_deletion_marker_is_retired(self):
        events = self._events([
            _hrow(pref="新潟県", cb="152153", nb="栃尾市", kind=None,
                  ca="削除", na="〃", date=datetime.datetime(2006, 1, 1)),
        ])
        self.assertEqual(events[0]["change_type"], "retired")
        self.assertIsNone(events[0]["code_after"])

    def test_empty_history_rejected(self):
        with self.assertRaises(ValueError):
            self._events([])


def _rec(code, name="某市", pref="某県"):
    return {"code": code, "municipality_name": name, "prefecture_name": pref}


def _ev(date, ctype, cb=None, nb=None, ca=None, na=None, pref="某県"):
    return {"group_id": 1, "date": date, "prefecture_name": pref,
            "change_type": ctype, "code_before": cb, "name_before": nb,
            "code_after": ca, "name_after": na, "reason": None}


class TestDeriveLifespans(unittest.TestCase):
    def test_created_without_code_after_rejected(self):
        with self.assertRaises(ValueError):
            derive_lifespans([], [_ev("2010-04-01", "created", ca=None)])

    def test_retired_without_code_before_rejected(self):
        with self.assertRaises(ValueError):
            derive_lifespans([], [_ev("2010-04-01", "retired", cb=None)])

    def test_name_comes_from_latest_event(self):
        events = [
            _ev("2010-04-01", "created", ca="026807", na="旧名市"),
            _ev("2020-04-01", "continued", cb="026807", nb="旧名市",
                ca="026807", na="新名市"),
        ]
        spans = {s["code"]: s for s in derive_lifespans([], events)}
        self.assertEqual(spans["026807"]["name"], "新名市")

    def test_order_independence(self):
        events = [
            _ev("2005-04-01", "retired", cb="203017", nb="臼田町"),
            _ev("2005-04-01", "continued", cb="203220", nb="望月町",
                ca="202171", na="佐久市"),
            _ev("2010-04-01", "created", ca="026807", na="某市"),
            _ev("2020-04-01", "retired", cb="026807", nb="某市"),
        ]
        forward = derive_lifespans([], events)
        shuffled = derive_lifespans([], list(reversed(events)))
        self.assertEqual(forward, shuffled)

    def test_mapping_rule_needs_last_appearance(self):
        # Code mapped away in 2010 but appearing again in 2015: the 2010
        # mapping row must not close its lifespan.
        events = [
            _ev("2010-04-01", "continued", cb="026801", nb="某町",
                ca="026807", na="某市"),
            _ev("2015-04-01", "continued", cb="026801", nb="某町",
                ca="026801", na="某町"),
        ]
        spans = {s["code"]: s for s in derive_lifespans([], events)}
        self.assertIsNone(spans["026801"]["valid_to"])

    def test_mapping_rule_fires_on_last_appearance(self):
        events = [
            _ev("2005-04-01", "continued", cb="203017", nb="臼田町",
                ca="202171", na="佐久市"),
        ]
        spans = {s["code"]: s for s in derive_lifespans([], events)}
        self.assertEqual(spans["203017"]["valid_to"], "2005-04-01")


class TestValidateLifespans(unittest.TestCase):
    def _base(self):
        records = [_rec("010006", name="北海道", pref="北海道")]
        events = [_ev("2010-04-01", "retired", cb="026801", nb="某町")]
        lifespans = [
            {"code": "010006", "name": "北海道", "prefecture_name": "北海道",
             "valid_from": None, "valid_to": None},
            {"code": "026801", "name": "某町", "prefecture_name": "某県",
             "valid_from": None, "valid_to": "2010-04-01"},
        ]
        return records, events, lifespans

    def test_valid_passes(self):
        validate_lifespans(*self._base())

    def test_same_day_from_to_passes(self):
        records, events, lifespans = self._base()
        lifespans[1]["valid_from"] = "2010-04-01"
        validate_lifespans(records, events, lifespans)

    def test_non_current_without_valid_to_rejected(self):
        records, events, lifespans = self._base()
        lifespans[1]["valid_to"] = None
        with self.assertRaises(ValueError):
            validate_lifespans(records, events, lifespans)

    def test_null_name_rejected(self):
        records, events, lifespans = self._base()
        lifespans[1]["name"] = None
        with self.assertRaises(ValueError):
            validate_lifespans(records, events, lifespans)

    def test_valid_to_before_from_rejected(self):
        records, events, lifespans = self._base()
        lifespans[1]["valid_from"] = "2011-04-01"
        with self.assertRaises(ValueError):
            validate_lifespans(records, events, lifespans)

    def test_event_code_missing_rejected(self):
        records, events, lifespans = self._base()
        events.append(_ev("2012-04-01", "created", ca="999999", na="幻市"))
        with self.assertRaises(ValueError):
            validate_lifespans(records, events, lifespans)


if __name__ == "__main__":
    unittest.main(verbosity=2)
