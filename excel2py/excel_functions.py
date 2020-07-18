"""
Python implementation of Excel functions

By Michael Grazebrook of Joined Up Finance Ltd
"""
import math
import datetime
import numbers
from collections import Iterable
from functools import reduce
from statistics import median
from excel2py.ex_datetime import ex_datetime, to_excel_number


def _to_number(arg):
    """
    cast arg to a number of equivalent meaning.

    datetime is a float of days since the Excel era

    timedelta is a float expressed in days

    anything else (e.g. a number stored as text) raises TypeError
    """
    
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


# count function
def COUNT(*args, value2=0):
    """
    Counts all values that are numeric in a column not excluding dates, e.g. COUNT(A1:A7)
    :param args: the arguments to count
    :param value2: the default argument
    :return: the length of the counted tuple of elements
    """
    return len(tuple(filter(lambda x: not isinstance(x, (bool, str)), *args)))


# median function
def MEDIAN(*args): # *args
    """
    Returns the median of a given list of values, e.g. =MEDIAN(1,2,3,4,5) returns 3
    """
    try:
        return median(*args)
    except TypeError:
        raise TypeError("median works on only strings values")


# trim function
def TRIM(val):
    """
    Trim removes all white spaces from a given cell, e.g. =TRIM(A1)
    """
    return " ".join(val.split())


# concatenate function
def CONCATENATE(value1,value2,option):
    """
    Concatenate joins two cells into a combined cell, e.g. =CONCATENATE(A2,"",B2)
    :param value1: represents the first cell
    :param value2: represents the second cell
    :param option: represents any white space(s) to be included between the two values 
    """
    if option == ' ':
        return option.join([value1, value2])
    return ''.join([value1, value2])


# counta function
def COUNTA(*args):
    """
    Counta counts all cells regardless of type but only skips empty cell, e.g. =COUNTA(value1, [value2], …)
    """
    return [len([a for a in arg if a != '']) for arg in args]


# mid function
def MID(text, start_num, num_chars):
    """
    The Excel MID function extracts a given number of characters from the middle of a supplied text string.
    For example, =MID("apple",2,3) returns "ppl". =MID (text, start_num, num_chars)
    
    :param text: The text to extract from.
    :param start_num: The location of the first character to extract.
    :param num_chars: The number of characters to extract.
    :return: The extracted text
    """
    
    if start_num and num_chars > 0:
        return text[start_num-1:(start_num+num_chars)-1]
    
    # raise ValueError if param start_num value is not positive integer
    if start_num == 0:
        raise ValueError('start_num value must be positive integers eg: 1')

    # raise ValueError if param start_num value is not a positive integer
    if num_chars == 0:
        raise ValueError('num_chars value must be positive integers eg: 1')

    # raise ValueError if param start_num or num_chars are negative values
    if start_num < 0 or num_chars < 0:
        raise ValueError('start_num and num_chars values cannot be negative values')


# replace function
def REPLACE(text, start_num, num_chars, option):
    """
    Replace, replaces text by position.
    For example, =REPLACE("apple##",2,3,"*") returns "p*l"
    
    :param text: The text to extract from.
    :param start_num: The location of the first character to extract.
    :param num_chars: The number of characters to extract.
    :param option: integer(s), string(s),char(s) or symbol(s) to be inserted into the position.
    :return: The replaced text
    """
    # raise TypeError if a non int param start_num value
    if not isinstance(start_num, int):
        raise TypeError('start_num should be of type int')

    # raise TypeError if a non int param num_chars value
    if not isinstance(num_chars, int):
        raise TypeError('num_chars should be of type int')

    # raise TypeError for negative param values
    if start_num < 0 or num_chars < 0:
        raise TypeError('start_num or num_chars should be of positive integers')

    # return the string in place
    return option.join([text[:start_num-1], text[(start_num+num_chars)-1:]])


# search function
def SEARCH(find_text, within_text):
    """
    Search for a word in a string and returns the position.
    For example, =SEARCH("we love python", "love") -> output(4)
    
    :param find_text: The text to find.
    :param within_text: The string of text.
    :return: The numeric position of matching text.
    """
    try:
        return within_text.index(find_text) + 1
    except ValueError:
        raise ValueError('No such charater available in string')


# abs function
def ABS(val):
    """
    Will return the absolute value of the param passed. 
    For example, =ABS(-13.40) -> output(13.40)
    
    :param val: The value to check
    :return: The absolute value of a number
    """
    if not isinstance(val, str):
        return abs(val)
    raise TypeError('Value must be of type str')


# exact function
def EXACT(val1, val2):
    """
    Check for equality between two text strings in a case-sensitive manner. 
    For example, =EXACT("Test","test") -> output(FALSE)
    
    :param val1: First value
    :param val2: Second Value
    :return: True/False
    """
    if val1 == val2:
        return True
    return False


# right function
def RIGHT(text, num_chars=1):
    """
    Extract text from the right of a string eg: apple
    :param text: The text from which to extract characters on the right.
    :param num_chars: [optional] The number of characters to extract, starting on the right. Optional, default = 1.
    :return : The Extracted text
    """
    tx = [text[-x:] for x in range(1, (num_chars+1))]
    return tx[-1]


# left function
def LEFT(text, num_chars=1):
    """
    Extract text from the left of a string
    :param text: The text from which to extract characters on the right.
    :param num_chars: [optional] The number of characters to extract, starting on the right. Optional, default = 1.
    """
    tx = [text[:x] for x in range(1, (num_chars+1))]
    return tx[-1]


# len function
def LEN(val):
    """
    Get the length of the value. Object of type 'int' are converted to type 'str'
    before their length can be returned
    :param val: The value to be checked
    :return : The length
    """
    if isinstance(val, (int, float)):
        return len(str(val))
    return len(val)


# ceiling function
def CEILING(number, significance=0):
    """
    Rounds up a number to the nearest multiple of significance
    :param number: The value you want to round
    :param significance: The multiple to which you want to round
    :return: The rounded up ceiling value or value error if any
    """
    if not isinstance(number, str):
        if significance == 0:
            return number
        
        # get the modulos of the number ans significance
        remainder = number % significance
        if remainder == 0:
            return number

        ceiling_value = number + significance - remainder
        return round(ceiling_value, 2)
    raise ValueError("#VALUE! error value")


#
# dt = ex_datetime(2018, 7, 5)
# print(MIN(dt))
