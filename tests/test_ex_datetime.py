"""
Tests for ex_datetime

By Michael Grazebrook of Joined Up Finance Ltd
"""

from excel2py.ex_datetime import ex_datetime
import datetime
import unittest


class TestExDateTime(unittest.TestCase):
    def test_new_from_int(self):
        for date_as_number, expect, note in (
            (43376.74, datetime.datetime(2018, 10, 3, 17, 45, 36), "03 Oct 2018 17:45:36"),
            # This test fails because of Microsoft incorrectly treats 1900 as a leap year.
            # https://support.microsoft.com/en-gb/help/214326/excel-incorrectly-assumes-that-the-year-1900-is-a-leap-year
            # (60, datetime.datetime(1900, 6, 28, 0, 0, 0), "01 Jan 1900 00:00:00"),
            (61, datetime.datetime(1900, 3, 1, 0, 0, 0), "01 Jan 1900 00:00:00"),
            (366, datetime.datetime(1900, 12, 31, 0, 0, 0), "01 Jan 1901 00:00:00"),
            (367, datetime.datetime(1901, 1, 1, 0, 0, 0), "01 Jan 1901 00:00:00"),
        ):
            # with self.subTest(msg=note, date_as_number=date_as_number, expect=expect):
            with self.subTest():
                dt = ex_datetime(date_as_number)
                self.assertEqual(dt, expect, "Constructed from a float")

    def test_add(self):
        dt = ex_datetime(43376)
        expect = ex_datetime(2018, 10, 5)
        self.assertEqual(dt + 2, expect, "date + days")
        self.assertEqual(2 + dt, expect, "days + date")

    def test_sub(self):
        dt = ex_datetime(2018, 11, 20, 13, 30, 20)
        expect = ex_datetime(2018, 11, 10, 1, 30, 20)
        self.assertEqual(dt - 10.5, expect, "dt - 10.5")


if __name__ == "__main__":
    d = ex_datetime(2018, 5, 10)
    # unittest.main()
