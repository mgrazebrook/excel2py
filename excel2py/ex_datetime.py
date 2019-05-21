"""
ex_datetime: modify datetime so it works with excel-style arithmetic

Excel dates are numbers since 1900-01-00 (I know, not a date). Since the
date itself is stored as a number, you can do arithmetic on it.

Subclass datetime and timedelta so they work in numeric contexts:
 - Define __add__, __radd__ etc to work with numbers too
 - return ex_datetime where datetime would be returned (e.g. date subtraction)

CODE READER QUESTION:
Might it be better to monkey patch datetime instead of subclassing it?
I'm concerned about mixing normal datetimes in subclass customisation code
and the risk of operations which haven't been overridden and don't return the
subclassed version.

By Michael Grazebrook of Joined Up Finance Ltd
"""

import datetime
import numbers

EXCEL_ERA = datetime.datetime(1900, 1, 1) - datetime.timedelta(days=2)
EXCEL_ERA_BASE = EXCEL_ERA.toordinal()


class ex_datetime(datetime.datetime):
    def __new__(cls, *args, **kwargs):
        """
        Add the ability to construct it from a number with the normal Excel meaning

        Excel dates are numbers since 1900-01-00 (I know, not a date)
        If the only arg is numeric, construct excel-style, else construct as usual.
        """
        if len(args) == 1:
            val = args[0]
            if isinstance(val, numbers.Number):
                date = EXCEL_ERA + datetime.timedelta(days=val)
                return super().__new__(cls, *date.timetuple()[:6])
            if isinstance(val, datetime.datetime):
                # Return a copy as ex_datetime instead of datetime
                return super().__new__(
                    cls, val.year, val.month, val.day,
                    val.hour, val.minute, val.second, val.microsecond, val.tzinfo
                )
        else:
            return super().__new__(cls, *args, **kwargs)

    def __add__(self, other):
        if isinstance(other, numbers.Number):
            return ex_datetime(to_excel_number(self) + other)
        # This generates an error
        return ex_datetime(super().__add__(other))

    def __radd__(self, days):
        return self.__add__(days)

    def __sub__(self, other):
        if isinstance(other, numbers.Number):
            return ex_datetime(to_excel_number(self) - other)
        return to_excel_number(self) - to_excel_number(other)

    def __rsub__(self, other):
        if isinstance(other, numbers.Number):
            raise NotImplemented("<Number> - <datetime> implies the Number ought to be a date - fix it in the Excel")
        return to_excel_number(other) - to_excel_number(self)
        # return ex_datetime(datetime.datetime.__rsub__(self, other))

    def __ge__(self, other):
        if other is None:
            return None
        return datetime.datetime.__ge__(self, other)

    def __gt__(self, other):
        if other is None:
            return None
        return datetime.datetime.__gt__(self, other)

    def __le__(self, other):
        if other is None:
            return None
        return datetime.datetime.__le__(self, other)

    def __lt__(self, other):
        if other is None:
            return None
        return datetime.datetime.__lt__(self, other)


def to_excel_number(datetime_value):
    """
    Convert a Python datetime to an Excel number
    """
    assert isinstance(datetime_value, datetime.datetime)
    delta = datetime.datetime.__sub__(datetime_value, EXCEL_ERA)
    return delta.total_seconds() / 24 / 60 / 60
