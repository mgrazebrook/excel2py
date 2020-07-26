"""
Test the Python implementation of Excel functions

By Michael Grazebrook of Joined Up Finance Ltd
"""
import unittest
import datetime
import excel2py.excel_functions as ef
from excel2py.ex_datetime import ex_datetime

# skipping:
# ISERROR(val), IFERROR(value, value_if_error),
# IFNA(value, value_if_na), DATE(year, month, day),


class TestToNumber(unittest.TestCase):
    def test_ok(self):
        delta = datetime.timedelta(days=2, hours=2, minutes=2, seconds=2)
        dt = datetime.datetime(2010, 5, 10, 12, 30, 20)
        test_cases = (
            (-1, -1, "Integer"),
            (1/7, 1/7, "float"),
            (dt,  40308.5210648, "datetime"),
            (delta, 2.084745370368180, "timedelta")
        )
        for val, expect, note in test_cases:
            with self.subTest(val=val, expect=expect, note=note):
                result = ef._to_number(val)
                self.assertAlmostEqual(result, expect, msg=note)

    def test_bad_type(self):
        with self.assertRaises(TypeError):
            ef._to_number("1")


class TestIterArgs(unittest.TestCase):
    def test_ok(self):
        d = datetime.datetime(2010, 5, 10, 12, 30, 20)
        test_cases = (
            ([1], [1], "one value"),
            ([2, 4, 16], [2, 4, 16], "list"),
            ([d], [40308.5210648], "date"),
        )
        for args, expect, msg in test_cases:
            with self.subTest(args=args, expect=expect, msg=msg):
                result = list(ef._iter_args(args))
                self.assertEqual(len(result), len(expect), msg + ": length check")
                for i in range(len(result)):
                    self.assertAlmostEqual(result[i], expect[i], msg=msg)

    def test_tuple_of_tuples(self):
        # Not supported since so far we only need it in a numerical context
        t = ((1, 2), (3, 4),)

        with self.assertRaises(TypeError):
            list(ef._iter_args(t))


class TestIf(unittest.TestCase):
    def test_true(self):
        self.assertEqual(ef.IF(True, 1, 0), 1)

    def test_false(self):
        self.assertEqual(ef.IF(False, 1, 0), 0)


class TestIsBlank(unittest.TestCase):
    def test_is_blank(self):
        self.assertTrue(ef.ISBLANK('   '))

    def test_not_blank(self):
        self.assertFalse(ef.ISBLANK(' x  '))


class TestRoundDown(unittest.TestCase):
    def test_ok(self):
        test_cases = (
            ((1.77, 1), 1.7),
            ((-1.77, 1), -1.7),
            ((123.45, -1), 120),
        )
        for args, expect in test_cases:
            with self.subTest(args=args, expect=expect):
                self.assertEqual(ef.ROUNDDOWN(*args), expect)


class TestSimpleFunctions(unittest.TestCase):
    def test_args(self):
        self.assertEqual(ef.SUM(1, 2, 3), 6)
        self.assertEqual(ef.SUM(3), 3)

    def test_min(self):
        dt = ex_datetime(2018, 7, 5)
        self.assertEqual(ef.MIN(1, 7, -4, 11), -4)
        self.assertEqual(ef.MIN(dt), dt)
        self.assertEqual(ef.MIN([dt+4, dt, dt+2]), dt)

    def test_max(self):
        self.assertEqual(ef.MAX(1, 7, -4, 11), 11)

    def test_year(self):
        self.assertEqual(ef.YEAR(datetime.datetime(2018, 3, 3)), 2018)

    def test_product(self):
        self.assertEqual(ef.PRODUCT([1, 2, 3, 4]), 24)

    def test_round(self):
        for val, places, result in (
            (128, -1, 130),
            (123.456, -1, 120),
            (123.456, 0, 123),
            (123.456, 1, 123.5),
            (123.456, 2, 123.46)
        ):
            with self.subTest(val=val, places=places, result=result):
                self.assertEqual(ef.ROUND(val, places), result)


class TestVLookup(unittest.TestCase):
    def test_vlookup(self):
        table = ((1, 2), (3, 4))

        for expected, val, range_lookup, name in (
                (None, 0, True, "range, before"),
                (2, 1, True, "range, first"),
                (2, 2, True, "range, mid"),
                (4, 3, True, "range, second"),
                (4, 4, True, "range, after"),
                (None, 0, False, "exact, before"),
                (2, 1, False, "exact, first"),
                (None, 2, False, "exact, mid"),
                (4, 3, False, "exact, second"),
                (None, 4, False, "exact, after"),
        ):
            with self.subTest(name):
                self.assertEqual(expected, ef.VLOOKUP(val, table, 2, range_lookup), name)


class TestBoolean(unittest.TestCase):
    def test_and(self):
        self.assertTrue(ef.AND(True))
        self.assertTrue(ef.AND(True, True))
        self.assertTrue(ef.AND([True, True]))
        self.assertFalse(ef.AND(False))
        self.assertFalse(ef.AND([True, False, True]))

    def test_or(self):
        self.assertTrue(ef.OR(True))
        self.assertTrue(ef.OR([True]))
        self.assertTrue(ef.OR([True, False]))
        self.assertFalse(ef.OR(False))
        self.assertFalse(ef.OR([False, False]))

    def test_no_args(self):
        with self.assertRaises(TypeError):
            self.assertTrue(ef.AND())  # Not legal in Excel either


class TestInt(unittest.TestCase):
    def test_ok(self):
        for val, expect in (
            (1.00, 1.00),
            (-1.00, -1.00),
            ( 1.50, 1.00),
            (-1.50, -2.00),
            (-1.01, -2.00),
        ):
            with self.subTest(val=val, expect=expect):
                self.assertEqual(ef.INT(val), expect)


# RIGHT test
class TestRight(unittest.TestCase):
    def test_with_optional_value(self):
        """
        Test to extract characters based on the second argument
        """
        expect = "cel"
        result = ef.RIGHT("Excel", 3)
        self.assertEqual(expect, result)

    def test_with_no_optional_value(self):
        """
        Test to extract characters if no second argument is passed
        """
        expect = 'l'
        result = ef.RIGHT("Excel")
        self.assertEqual(expect, result)
   
    def test_with_zero_as_len(self):
        """
        Test to return any empty string if 0 is passed as the len
        """
        result = ef.RIGHT("ABC",0)
        self.assertEqual('',result)

    def test_value_errors(self):
        """
        Throw a value error if the len is a string or a negative value
        """
        with self.assertRaises(ValueError):
            ef.RIGHT("ABC","len")
        with self.assertRaises(ValueError):
            ef.RIGHT("ABC",-1)
        with self.assertRaises(ValueError):
            ef.RIGHT("ABC",True)
        with self.assertRaises(ValueError):
            ef.RIGHT("ABC",[5,10,15])
        with self.assertRaises(ValueError):
            ef.RIGHT("ABC",(25,30,35))




if __name__ == "__main__":
    unittest.main()
