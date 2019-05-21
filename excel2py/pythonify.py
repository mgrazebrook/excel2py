"""
TatSu semantics class to convert an Excel expression into Python.
https://tatsu.readthedocs.io/en/stable/semantics.html

By Michael Grazebrook of Joined Up Finance Ltd
"""

import re
import keyword


# A list of names used by Excel: these shouldn't be prefixed with 'self.'
# List taken from:
# https://support.office.com/en-us/article/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188
# 30 Aug 2018
ALL_EXCEL_FUNCTIONS = {
   'MID','MIDB','REPLACE','REPLACEB','RIGHT','RIGHTB','SEARCH','SEARCHB',
   'FIND','FINDB','LEFT','LEFTB','LEN','LENB','SUMX2PY2','SUMXMY2',
   'SYD','T','TAN','TANH','TBILLEQ','TBILLPRICE','TBILLYIELD','TDIST',
   'TEXT','TIME','TIMEVALUE','TINV','TODAY','TRANSPOSE','TREND','TRIM',
   'TRIMMEAN','TRUE','TRUNC','TTEST','TYPE','UPPER','VALUE','VAR','VARA',
   'VARP','VARPA','VDB','VLOOKUP','WEEKDAY','WEEKNUM','WEIBULL','WORKDAY',
   'XIRR','XNPV','YEAR','YEARFRAC','YIELD','YIELDDISC','YIELDMAT','ZTEST',
   'ABS','ACCRINT','ACCRINTM','ACOS','ACOSH','AGGREGATE','ADDRESS','AMORDEGRC',
   'AMORLINC','AND','AREAS','ASC','ASIN','ASINH','ATAN','ATAN2','ATANH',
   'AVEDEV','AVERAGE','AVERAGEA','AVERAGEIF','AVERAGEIFS','BAHTTEXT',
   'BASE','BESSELI','BESSELJ','BESSELK','BESSELY','BETADIST','BETAINV',
   'BIN2DEC','BIN2HEX','BIN2OCT','BINOMDIST','BINOM.DIST.RANGE','CALL',
   'CEILING','CEILING.PRECISE','CELL','CHAR','CHIDIST','CHIINV','CHITEST',
   'CHOOSE','CLEAN','CODE','COLUMN','COLUMNS','COMBIN','COMPLEX','CONCATENATE',
   'CONFIDENCE','CONVERT','CORREL','COS','COSH','COUNT','COUNTA','COUNTBLANK',
   'COUNTIF','COUNTIFS','COUPDAYBS','COUPDAYS','COUPDAYSNC','COUPNCD',
   'COUPNUM','COUPPCD','COVAR','CRITBINOM','CUBEKPIMEMBER','CUBEMEMBER',
   'CUBEMEMBERPROPERTY','CUBERANKEDMEMBER','CUBESET','CUBESETCOUNT','CUBEVALUE',
   'CUMIPMT','CUMPRINC','DATE','DATEDIF','DATEVALUE','DAVERAGE','DAY',
   'DAYS360','DB','DCOUNT','DCOUNTA','DDB','DEC2BIN','DEC2HEX','DEC2OCT',
   'DEGREES','DELTA','DEVSQ','DGET','DISC','DMAX','DMIN','DOLLAR','DOLLARDE',
   'DOLLARFR','DPRODUCT','DSTDEV','DSTDEVP','DSUM','DURATION','DVAR',
   'DVARP','EDATE','EFFECT','EOMONTH','ERF','ERFC','ERROR.TYPE','EUROCONVERT',
   'EVEN','EXACT','EXP','EXPONDIST','FACT','FACTDOUBLE','FALSE','FDIST',
   'FINV','FISHER','FISHERINV','FIXED','FLOOR','FLOOR.PRECISE','FORECAST',
   'FORECAST.ETS.STAT','FREQUENCY','FTEST','FV','FVSCHEDULE','GAMMADIST',
   'GAMMAINV','GAMMALN','GAMMALN.PRECISE','GCD','GEOMEAN','GESTEP','GETPIVOTDATA',
   'GROWTH','HARMEAN','HEX2BIN','HEX2DEC','HEX2OCT','HLOOKUP','HOUR',
   'HYPERLINK','HYPGEOM.DIST','HYPGEOMDIST','IF','IFERROR','IMABS','IMAGINARY',
   'IMARGUMENT','IMCONJUGATE','IMCOS','IMDIV','IMEXP','IMLN','IMLOG10',
   'IMLOG2','IMPOWER','IMPRODUCT','IMREAL','IMSIN','IMSQRT','IMSUB',
   'IMSUM','INDEX','INDIRECT','INFO','INT','INTERCEPT','INTRATE','IPMT',
   'IRR','ISBLANK','ISERR','ISERROR','ISEVEN','ISLOGICAL','ISNA','ISNONTEXT',
   'ISNUMBER','ISODD','ISREF','ISTEXT','ISPMT','JIS','KURT','LARGE',
   'LCM','LINEST','LN','LOG','LOG10','LOGEST','LOGINV','LOGNORMDIST',
   'LOOKUP','LOWER','MATCH','MAX','MAXA','MDETERM','MDURATION','MEDIAN',
   'MIN','MINA','MINUTE','MINVERSE','MIRR','MMULT','MOD','MODE','MONTH',
   'MROUND','MULTINOMIAL','N','NA','NEGBINOMDIST','NETWORKDAYS','NOMINAL',
   'NORMDIST','NORMINV','NORMSDIST','NORMSINV','NOT','NOW','NPER','NPV',
   'OCT2BIN','OCT2DEC','OCT2HEX','ODD','ODDFPRICE','ODDFYIELD','ODDLPRICE',
   'ODDLYIELD','OFFSET','OR','PEARSON','PERCENTILE','PERCENTRANK','PERMUT',
   'PHONETIC','PI','PMT','POISSON','POWER','PPMT','PRICE','PRICEDISC',
   'PRICEMAT','PROB','PRODUCT','PROPER','PV','QUARTILE','QUOTIENT','RADIANS',
   'RAND','RANDBETWEEN','RANK','RATE','RECEIVED','REGISTER.ID','REPT',
   'ROMAN','ROUND','ROUNDDOWN','ROUNDUP','ROW','ROWS','RSQ','RTD','SECOND',
   'SERIESSUM','SIGN','SIN','SINH','SKEW','SLN','SLOPE','SMALL','SQL.REQUEST',
   'SQRT','SQRTPI','STANDARDIZE','STDEV','STDEVA','STDEVP','STDEVPA',
   'STEYX','SUBSTITUTE','SUBTOTAL','SUM','SUMIF','SUMIFS','SUMPRODUCT',
   'SUMSQ','SUMX2MY2','MINIFS','MODE.MULT','MODE.SNGL','MUNIT','NEGBINOM.DIST',
   'NETWORKDAYS.INTL','NORM.DIST','NORM.INV','NORM.S.DIST','NORM.S.INV',
   'NUMBERVALUE','PDURATION','PERCENTILE.EXC','PERCENTILE.INC','PERCENTRANK.EXC',
   'PERCENTRANK.INC','PERMUTATIONA','PHI','POISSON.DIST','QUARTILE.EXC',
   'QUARTILE.INC','RANK.AVG','RANK.EQ','RRI','SEC','SECH','SHEET',
   'SHEETS','SKEW.P','STDEV.P','STDEV.S','SWITCH','T.DIST','T.DIST.2T',
   'T.DIST.RT','TEXTJOIN','T.INV','T.INV.2T','T.TEST','UNICHAR',
   'UNICODE','VAR.P','VAR.S','WEBSERVICE','WEIBULL.DIST','WORKDAY.INTL',
   'XOR','Z.TEST','ACOT','ACOTH','ARABIC','BETA.DIST','BETA.INV',
   'BINOM.DIST','BINOM.INV','BITAND','BITLSHIFT','BITOR','BITRSHIFT',
   'BITXOR','CEILING.MATH','CHISQ.DIST','CHISQ.DIST.RT','CHISQ.INV',
   'CHISQ.INV.RT','CHISQ.TEST','COMBINA','CONCAT','CONFIDENCE.NORM',
   'CONFIDENCE.T','COT','COTH','COVARIANCE.P','COVARIANCE.S','CSC',
   'CSCH','DAYS','DBCS','DECIMAL','ENCODEURL','ERF.PRECISE','ERFC.PRECISE',
   'EXPON.DIST','F.DIST','F.DIST.RT','FILTERXML','F.INV','F.INV.RT',
   'FLOOR.MATH','FORECAST.ETS','FORECAST.ETS.CONFINT','FORECAST.ETS.SEASONALITY',
   'FORECAST.LINEAR','FORMULATEXT','F.TEST','GAMMA','GAMMA.DIST',
   'GAMMA.INV','GAUSS','IFNA','IFS','IMCOSH','IMCOT','IMCSC',
   'IMCSCH','IMSEC','IMSECH','IMSINH','IMTAN','ISFORMULA','ISO.CEILING',
   'ISOWEEKNUM','LOGNORM.DIST','LOGNORM.INV','MAXIFS',
}


class Pythonify:
    """
    Handle a single Excel expression. Convert it into Python.

    Usage:
        parser = tatsu.compile(grammar)
        parser.parse(excel_expression, semantics=Pythonify)
    The Excel expression is everything after '='
    """

    def __init__(self, functions: set, aliases: dict = {}):
        """
        Parse an Excel expression and reformulate it as Python

        :param globals: Globally defined name other than Excel functions
                Anything in this list doesn't get a 'this.' prefix.
                e.g. { 'my_tk_function' }
        :param aliases: Used to ether rename a name or give a range a name
                e.g. { 'lambda': 'my_lambda', 'Sheet1!A1' : 'limburger_amount' }
        """
        # If ranges used here don't have names, the caller will need to look them up.
        # This is built up as we progress.
        self.ranges = set()

        # TODO: kwlist should surely be managed via aliases?
        self.globals = set(keyword.kwlist) | ALL_EXCEL_FUNCTIONS | functions

        self.aliases = aliases

        self.sheet = None

    @staticmethod
    def start(ast):
        py_expression = Pythonify._default(ast)
        # TODO: Can I use autopep8 (pycodestyle) as a library to pretty print it?
        # https://github.com/hhatto/autopep8
        return py_expression + '\n'

    @staticmethod
    def operator(ast):
        """

        # TODO: The equals problem is, I think, generally true. It's possible
        # all operators need to be replaced with functions.
        :param ast: string
        :return: Python equivalent
        """
        assert isinstance(ast, str)
        if ast == '=':
            # TODO: Excel string comparison is case insensitive.
            # TODO: Excel comparison with an error returns the error
            # e.g. #NUM!=7 returns #NUM, #DIV/0!=7 returns #DIV! etc
            return ' == '
        if ast == '^':
            return '**'
        if ast == '&':  # string concatenation
            return ' + '
        if ast == '<>':
            return ' != '
        return ast

    def range(self, ast):
        range_name = _flatten(ast)
        if self.sheet and '!' not in range_name:
            if re.search('\W', range_name):
                range_name = f"'{self.sheet}'!{range_name}"
            else:
                range_name = f"{self.sheet}!{range_name}"
        try:
            name = self.aliases[range_name]
        except KeyError:  # NB: Not an error
            name = self.py_name(range_name)
            # Track un-named ranges so we can get their formulae later
            self.ranges.add(range_name)
        if name in self.globals:
            return name
        return f'self.{name}'

    @staticmethod
    def py_name(range_name):
        """
        :param range_name: Name such as "'Proforma Preserved'!$E$46"
        :return: python legal name such as "ProformaPreservedE46"
        """
        return re.sub("[ _!:$']+", '', range_name)

    @staticmethod
    def number(ast):
        text = _flatten(ast)
        if text[-1] == '%':
            return repr(float(text[:-1])/100)
        return text

    def name(self, ast):
        text = _flatten(ast)
        if text in self.aliases:
            text = self.aliases[text]
        if text == 'TRUE':
            return 'True'
        if text == 'FALSE':
            return 'False'
        if text in self.globals:
            return text
        return 'self.' + text

    @staticmethod
    def _default(ast):
        return _flatten(ast)


def _flatten(ast):
    if isinstance(ast, str):
        return ast
    if isinstance(ast, list):
        return ''.join([_flatten(bit) for bit in ast])
    assert False, repr(ast)  # Should be unreachable
    return repr(ast)


if __name__ == "__main__":
    from expression_parser import expression_parser

    lines = [
        "2",
        "-2",
        "tk_function(D18, Sheet1!B3:C5)",
        "SUM(E4:F5)",
        "E4^Sheet1!A1",
        'ROUND(IF(G6="FULL",G21,IF(G6="FIXED",G23,G25))*G12,2)-G26'
    ]
    aliases = {
        'Sheet1!A1': 'my_input',
        'Sheet1!B3:C5': 'a_table'
    }
    globals = {
        'tk_function'
    }
    semantics = Pythonify(globals, aliases)
    parser = expression_parser()
    for line in lines:
        print(">", line)
        semantics.sheet = 'Sheet1'
        result = parser.parse(line, semantics=semantics)
        print("<", result)
