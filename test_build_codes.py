"""Tests for build_codes.py. Run: python3 test_build_codes.py"""
import datetime
import unittest

from build_codes import (
    check_digit,
    classify_type,
    parse_date_cell,
    to_fullwidth_kana,
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


if __name__ == "__main__":
    unittest.main(verbosity=2)
