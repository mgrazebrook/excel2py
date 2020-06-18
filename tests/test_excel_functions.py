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

    # def test_min_mixed_range(self):
    #     result = ef.MIN(((8,4), 7, 3), (9, 2.5), 2.0)
    #     expect = 2.0
    #     self.assertEqual(expect, result)

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


# Count test
class TestCount(unittest.TestCase):
    """
    The COUNT function counts the number of cells that contain numbers
    """
    def test_one(self):
        expect = 1
        result = ef.COUNT((7, ))
        self.assertEqual(expect, result)

    def test_different_types(self):
        """
        This is from the Excel documentation excluding the #DIV/0:
        Excel  Formulas and functions  Functions  COUNT function
        https://support.microsoft.com/en-gb/office/count-function-a59cd7fc-b623-4d93-87a4-d23bf411294c?ui=en-us&rs=en-gb&ad=gb#:~:text=Use%20the%20COUNT%20function%20to,numbers%2C%20the%20result%20is%205.
        """
        data = (ex_datetime(2008, 8, 8), 19, 22.24, True)
        expect = 3
        result = ef.COUNT(data)
        self.assertEqual(expect, result)

    def test_mixed_range(self):
        """
        Counts the number of cells that contain numbers in cells A2 through A7,
        and the value 2 : e.g =COUNT(A2:A7,2)
        """
        expect = 5
        range = ((2, 4.8), (1, 12))
        # unpack the nested tuple
        data = [element for rang in range for element in rang]
        result = ef.COUNT((*data, 28))
        self.assertEqual(expect, result)

    def test_count_error(self):
        pass  # TODO: #DIV/0 etc


# Median test
class TestMedian(unittest.TestCase):
    """
    Test to get the median of a list of values
    Will raise an error if strings are passed
    """
    def test_ok(self):
        data = [1, 2, 3, 12, 15, 45]
        result = ef.MEDIAN(data)
        self.assertTrue(result)

    def test_error(self):
        data = [1, 2, 3, 12, 15, 45, 'one', 25, 'two', ' ']
        with self.assertRaises(TypeError, msg="Only "): ef.MEDIAN(data)


# Trim test
class TestTrim(unittest.TestCase):
    """
    Test to remove all whitespace from a string
    """
    def test_ok(self):
        text = ' PYTHON       UNITTEST  TRIM  '
        result = ef.TRIM(text)
        self.assertTrue(result)


# Concatenate test
class TestConcatenate(unittest.TestCase):
    """
    Test to join to cells together to form one
    """
    def test_ok(self):
        cell1 = 'one'
        cell2 = 'two'
        option = ''
        result = ef.CONCATENATE(cell1, cell2, option)
        self.assertTrue(result)


# Counta test
class TestCounta(unittest.TestCase):
    """
    Test to counts all cells regardless of type but only skips empty cell
    """
    def test_ok(self):
        data = ['one', 1, 'two', '', 3, 'three', 25, '']
        result = ef.COUNTA(data)
        self.assertTrue(result)



if __name__ == "__main__":
    unittest.main()
