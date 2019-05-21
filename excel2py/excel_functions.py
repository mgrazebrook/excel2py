"""
Python implementation of Excel functions

By Michael Grazebrook of Joined Up Finance Ltd
"""
import math
import datetime
import numbers
from collections import Iterable
from functools import reduce
from excel2py.ex_datetime import ex_datetime, to_excel_number


def _to_number(arg):
    if isinstance(arg, numbers.Number):
        return arg

    if isinstance(arg, datetime.datetime):
        # Excel actually displays zero as 00 Jan 1900, an invalid date.
        # It can't handle earlier dates.
        return to_excel_number(arg)

    if isinstance(arg, datetime.timedelta):
        return arg.total_seconds() / 60 / 60 / 24

    raise TypeError(f"{repr(arg)} {arg.__class__} doesn't seem to be a number")


def _iter_args(*args):
    """
    For functions which act on a range or a list of values, e.g. max(1,2) or max(A1:A2)

    If a value isn't numeric, convert it to a number if possible.
    :return: iterator
    """
    def arg_generator(iterable):
        for val in iterable:
            if isinstance(val, str):
                continue  # Excel group functions skip text
            yield _to_number(val)

    if len(args) == 1 and isinstance(args[0], Iterable):
        return arg_generator(args[0])
    return arg_generator(args)


def _group_function(fn, *args):
    """
    functools.reduce but with the additional option that args is a single iterable value

    Return the first arg if there is only one.
    :param fn: function passed to reduce: fn(cumulative_value, value)
    :param args: Either one iterable arg (e.g. a range) or one or more values
            Values can be numbers or datetime.
    :return: an ex_datetime or Number
    """
    # Use a datetime context if all arguments are datetime
    if len(args) == 1 and isinstance(args[0], Iterable):
        args = args[0]
    is_datetime = reduce(
        lambda truth, val: truth and isinstance(val, (datetime.datetime, str)),  # strings are ignored
        args, True)
    # Excel group functions simply ignore text
    ret = reduce(fn, _iter_args(*args))

    if is_datetime:
        return ex_datetime(ret)
    return ret


def IF(test, ok, bad):
    if test:
        return ok
    return bad


def ISBLANK(val):
    return val is None or val.strip() == ''


def ISERROR(val):
    """
    This is imperfect: Excel has many kinds of error:
    "Value refers to any error value (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!)."
    :param val:
    :return:
    """
    return val is None


def ROUND(value, digits):
    assert digits == int(digits)
    if value is None:
        return None
    return round(value, int(digits))


def IFERROR(value, value_if_error):
    if ISERROR(value):
        return value_if_error
    return value


def ROUNDDOWN(val, decimal_places):
    """
    Excel rounds towards zero

    Examples:
        ROUNDDOWN(1.77, 1) = 1.7
        ROUNDDOWN(-1.77, 1) = -1.7
        ROUNDDOWN(123.45, -1) = 120
    :param val:
    :param decimal_places:
    :return:
    """
    # Cute alternative:
    # return val // 10**-decimal_places / 10**decimal_places
    multiplier = 10**decimal_places
    if val > 0:
        return math.floor(val * multiplier) / multiplier
    return math.ceil(val * multiplier) / multiplier


def IFNA(value, value_if_na):
    if value is None:
        return value_if_na
    return value


def INT(val):
    """Round down to the nearest integer, e.g. -1.01 returns -2"""
    return math.floor(val)


def DATE(year, month, day):
    return datetime.datetime(year, month, day)


def DAY(date_value):
    return  date_value.day


def MONTH(date_value):
    return date_value.month


def YEAR(date_value):
    return date_value.year


# Range functions
def MIN(*args):
    """Return the minimum of a range or list of Number or datetime"""
    return _group_function(min, *args)


def MAX(*args):
    return _group_function(max, *args)


def SUM(*args):
    """
    Return the sum of a range the sum or its arguments

    :param args: Range or non-string arguments
    :return: a number or ex_datetime (though the latter would be a bit odd)
    """
    return _group_function(lambda x, y: x + y, *args)


def PRODUCT(*args):
    return _group_function(lambda x, y: x * y, *args)


def OR(*args):
    return _group_function(lambda a, b: a or b, *args)


def AND(*args):
    return _group_function(lambda a, b: a and b, *args)


def VLOOKUP(value, table, column, range_lookup=True):
    for i, row in enumerate(table):
        if value == row[0]:
            return row[column-1]
        if range_lookup and value < row[0]:
            if i:
                return table[i-1][column-1]
            return None
    if range_lookup:
        return table[-1][column-1]
    return None

#
# dt = ex_datetime(2018, 7, 5)
# print(MIN(dt))
