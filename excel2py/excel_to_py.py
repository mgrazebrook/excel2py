"""
Core implementation of the excel2py utility

TODO: One generated class per tab, or perhaps one per tab named in the config
If from the config, there'd be a tab name (=class name) and optional sub-class name for customisation

TODO: Use the variable naming convention which includes the field ID and links with automation.
To maintain the generality of the tool, this might be done via some sort of special configuration for overriding
names.

By Michael Grazebrook of Joined Up Finance Ltd
"""
import win32com.client as win32
import re
import sys
from io import StringIO
import difflib
import datetime

from excel2py.expression_parser import expression_parser
from excel2py.pythonify import Pythonify


class FileSection:
    """
    Abstract base class. Gather the text for a section of the output file.
    """
    def __init__(self, comment):
        self.text = StringIO()
        self.text.write(f"\n    # {comment}\n\n")
        self.preamble()

    def result(self):
        """
        :return: the text for a section of the output file
        """
        self.postscript()
        return self.text.getvalue()

    def preamble(self):
        """
        Text to write before the per-cell section
        """
        pass

    def do_name(self, name):
        """
        Process a variable name (generally a cell or range)
        Write any text for the variable, any additional processing
        :param name: Excel Name object
        :return: True if handled, else None
        """
        pass

    def postscript(self):
        """
        Text to write after the per-cell section
        """
        pass


class BadSection(FileSection):
    """
    List as comments things which we want the user to check
    For example variable names we couldn't do anything with
    This section also 'handles' anything we deliberately ignore.
    """
    def do_name(self, name):
        """
        Handle names we ignore or don't understand

        If we ignore them, silently handle them. If we don't understand them,
        include them as a comment.
        :param name: Excel Name object
        :return: True if handled, else None
        """
        if isinstance(name, DuckTypeName):
            return None  # Not relevant
        if name.RefersTo == '=#NAME?':
            return True
        if "!Print_Area" in name.Name:
            return True
        if "#REF!" in name.RefersTo:
            self.text.write(f"    # {name.Name} = '{name.RefersTo}'\n")
            return True


class CalculationSection(FileSection):
    """
    Prints the calculate(...) method

    This processes Names from the Inputs and Outputs sheets.
    """
    def __init__(self, comment, inputs, outputs):
        self.inputs = inputs
        self.outputs = outputs
        super().__init__(comment)

    def preamble(self):
        # TODO: Know which inputs are datetime and convert them to ex_datetime
        # It's a detail which should be hidden from the caller.
        self.text.write("    def __init__(self,\n")
        for var in self.inputs.values():
            self.text.write(f"        {var},\n")
        self.text.write("        **args):\n")
        for var in self.inputs.values():
            self.text.write(f"        self.{var} = {var}\n")
        self.text.write("        self.private_construction()\n")
        self.text.write("\n")
        self.text.write("    def calculate(self):\n")
        output_names = wrap_text(repr(list(self.outputs.values())), 120, 12*' ')
        self.text.write(f"        return namedtuple('CalcResult',\n")
        self.text.write(f"{output_names})(\n")
        for value in self.outputs.values():
            self.text.write(f"            {value}=self.{value},\n")
        self.text.write("    )\n\n")

    def do_name(self, name):
        """
        Exclude inputs from class constants: they are instance variables.
        :param name:
        :return:
        """
        if name.Name in self.inputs.values():
            return True


class ConstantSection(FileSection):
    """
    Constants become class variables.
    prerequisite: Non-formula cell with a value
    """
    def __init__(self, comment, valid_date_formats):
        self.valid_date_formats = valid_date_formats
        super().__init__(comment)

    def do_name(self, name):
        cells = name.RefersToRange

        if isinstance(cells.Value2, str):
            if '"' in cells.Value2:
                value = f"'{cells.Value2}'"
            else:
                value = f'"{cells.Value2}"'
        elif is_date(cells, self.valid_date_formats):
            value = f'ex_datetime({cells.Value2})'
        else:
            value = cells.Value2

        self.text.write(f"    {name.Name} = {value}\n")
        return True


class PropertySection(FileSection):
    """
    Property sections translate Excel formulae to Python @property methods
    """
    def __init__(self, comment, excel_to_py):
        super().__init__(comment)
        self.inputs = excel_to_py.config.inputs
        self.excel_to_py = excel_to_py
        self.names = []

    def do_name(self, name):
        cells = name.RefersToRange
        if not cells.HasFormula:
            return None

        if is_tuple_formula(cells):
            # TODO: Proper implementation of this
            value = deduce_tuple_formula(cells)
        else:
            value = self.excel_to_py.reformulate(cells)

        for input_ref in self.inputs:
            # TODO: Works for the current case but would could fail. Alias list for parse?
            # e.g. if Inputs!D1 and Inputs!D12 are both valid
            if isinstance(value, str) and input_ref in value:
                assert value, f"{value} {name.Name}"
                value = value.replace(input_ref, self.inputs[input_ref])

        self.names.append("_" + name.Name)

        self.text.write(
            "    @property\n"
            f"    def {name.Name}(self):\n"
            f"        if self._{name.Name} is not None:\n"
            f"            return self._{name.Name}\n"
            f"        self._{name.Name} = {value}\n"
            f"        return self._{name.Name}\n\n"
        )
        return True

    def postscript(self):
        """
        Initialise private variables to support calculating properties once only.
        """
        self.text.write(
            "\n\n"
            "    def private_construction(self):\n"
        )
        for name in self.names:
            self.text.write(f"        self.{name} = None\n")
        self.text.write("\n")


class DuckTypeName:
    """
    For when I want to use a Excel Name-like object without changing the spreadsheet.
    """
    def __init__(self, range_name, cells):
        self.Name = range_name
        self.RefersToRange = cells


class ExcelToPy:
    def __init__(self, config):
        self.config = config
        self.parser = expression_parser()

        # Prefix aliases with 'self.'
        self.inouts.update(self.formulae_on_sheets(config.input_sheets, config.inputs))
        self.outputs.update(self.formulae_on_sheets(config.output_sheets, config.outputs))
        aliases = {}
        aliases.update(config.inputs)
        aliases.update(config.outputs)
        self.pythonify = Pythonify(config.globals, aliases)

    def reformulate(self, cells):
        assert cells.Formula.startswith('='), cells.Formula
        formula = cells.Formula[1:]  # skip the '='
        self.pythonify.sheet = cells.Worksheet.Name
        return self.parser.parse(formula, semantics=self.pythonify)

    def formulae_on_sheets(self, sheets, aliases):


    def add_names_as_aliases(self, book):
        """
        Update Pythonify's aliases with the named cells.

        This means that if a named cell is used in R1C1 form, it still gets its name.
        Aliases from the config take precedence.
        :param book: Workbook object
        :return: None
        """
        range_to_name = {
            name.RefersTo[1:]: name.Name
            for name in book.Names
            if name.RefersTo != '=#NAME?'
            and "!Print_Area" not in name.Name
        }
        range_to_name.update(self.pythonify.aliases)
        self.pythonify.aliases = range_to_name

    def generate(self):
        xl, book = self._connect_to_excel()

        self.add_names_as_aliases(book)

        # The section which uses a name defines section order.
        sections = [
            BadSection("EXCEL VARIABLES WITH NO USABLE FORMULA"),
            CalculationSection("External interface", self.config.inputs, self.config.output_sheets, self.config.outputs),
            PropertySection("PROPERTIES", self),
            ConstantSection("CONSTANTS", self.config.valid_date_formats),
        ]
        for name in book.Names:
            for section in sections:
                if section.do_name(name):
                    # Names are handled by the first section which will accept them.
                    break
            else:
                print("Do something about", name.Name)

        # Warning: Changes Pythonify, moving a name from ranges to aliases
        while self.pythonify.ranges:

            range_name = self.pythonify.ranges.pop()
            sheet_name, cell_name = range_name.split('!')
            cells = book.Sheets[sheet_name.strip("'")].Range(cell_name)
            py_name = Pythonify.py_name(range_name)
            name = DuckTypeName(py_name, cells)
            self.pythonify.aliases[range_name] = py_name
            for section in sections:
                if section.do_name(name):
                    break
            # get the new name
            # add it to aliases
            # process as above - which could add new ranges
            # TODO:

        self._write_class(reversed(sections))
        print(self.pythonify.ranges)

    # def _formula_cell(self, name, cells):
    #     if is_tuple_formula(cells):
    #         # raise NotImplementedError("Tuple formulae are ranges of related cells."
    #         #                          "I'd like to replace them with a function.")
    #         value = deduce_tuple_formula(cells)
    #     else:
    #         assert cells.Formula.startswith('='), cells.Formula
    #         formula = cells.Formula[1:]  # skip the '='
    #         value = ExcelToPy.parser.parse(formula, semantics=Pythonify)
    #
    #     for input_ref in self.config.inputs:
    #         # TODO: Works for the current case but would could fail
    #         # e.g. if Inputs!D1 and Inputs!D12 are both valid
    #         if isinstance(value, str) and input_ref in value:
    #             assert value, f"{value} {name.Name}"
    #             value = value.replace(input_ref, self.config.inputs[input_ref])
    #
    #     return (
    #         "    @property\n"
    #         f"    def {name.Name}(self):\n"
    #         f"        return {value}\n"
    #     )

    def _write_class(self, sections):
        with open(self.config.output, 'w') as f:
            f.write(
                '# WARNING: AUTOMATICALLY GENERATED CODE\n'
                f'# Override this class using {self.config.class_name}\n'
                f'# Generated at {datetime.datetime.now().isoformat()}\n'
                f'# {" ".join(sys.argv)}\n'
                '\n'
                'from collections import namedtuple\n'
                'from excel2py.excel_functions import *\n'
                'from excel2py.base_proforma_calc import BaseProformaCalc\n'
                f'{self.config.imports}\n'
                'from excel2py.ex_datetime import ex_datetime\n'
                '\n'
                '\n'
                f'class {self.config.gen_class_name}(BaseProformaCalc):\n'
            )
            for section in sections:
                f.write(section.result())

    def _inputs_and_outputs(self):
        """
        Text for class variables defining the function interface

        TODO: For now, I'm listing them all in config. It would be nice to have
        the config just declare the input/output sheets and assume all variables there
        are inputs/outputs.

        :return: The text for the class member varaibles listing the inputs and outputs
        """
        out = StringIO()
        out.write("    inputs = {'")
        out.write("', '".join(self.config.inputs.keys()))
        out.write("'}\n")
        out.write("    outputs = {'")
        out.write("', '".join(self.config.outputs.keys()))
        out.write("'}\n\n")
        return out.getvalue()

    def _connect_to_excel(self):
        xl = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            book = xl.Workbooks(self.config.spreadsheet)
            return xl, book
        except win32.pywintypes.com_error as err:
            print(f"You must have '{self.config.spreadsheet}' open in Excel first.")
            print(f"{err}")
            exit(1)


def is_date(cells, valid_formats):
    """
    Does the Excel number format contain a date convertion?

    :param cells: A single cell with a value which could be a date
    :param valid_formats: Sequence of valid Excel date format strings
            e.g.  = ('dd/mm/yyyy',)
    :return: bool
    """
    try:
        float(cells.Value2)
    except ValueError:
        return False
    except TypeError:  # typically a tuple (range of cells)
        return False

    return cells.NumberFormat in valid_formats


def is_tuple_formula(cells):
    """
    Is it a range containing many cells containing formulae?

    Such formulae are usually related so we may be able to replace them
    with a function.
    :param cells:
    :return:
    """
    if (isinstance(cells.Formula, tuple)
        and cells.HasFormula
        and re.match('\$[A-Z]+\$[0-9]+:\$[A-Z]+\$[0-9]+', cells.Address)):
        return True
    return False


def deduce_tuple_formula(cells):
    """
    Try to work out a generic formula instead of a tuple of calculations

    This will work if the cells' formulae are essentially the same and
    refer to a named range. We'll try to deduce the other named range.
    :param cells:
    :return: Python text with equivalent functionality
    """
    # diff the first row and use it to check if it relates to another range.
    first = cells.Formula[0][0]
    second = cells.Formula[1][0]
    matches = difflib.SequenceMatcher(a=first, b=second)

    blocks = [
        b
        for b in matches.get_matching_blocks()
        if b.size # discard the zero length one at the end
    ]

    diffs = [
        first[b1.a + b1.size:b2.a]
        for b1, b2 in zip(blocks[:-1], blocks[1:])
    ]
    # TODO: work out the related arrays
    # Diffs should be the row indexes into the other arrays.
    # I'm assuming all arrays are one dimensional.
    # From the top left cell of this array, I can now find
    # the top left of the array(s) this formula uses.
    # I'll need to search book.Names for arrays starting there.

    # Must contain all the common bits
    common = [
        first[match.a:match.a + match.size]
        for match in blocks
    ]
    for row in cells.Formula:
        for formula in row:
            for fragment in common:
                if fragment not in formula:
                    return cells.Formula  # we can't do it.


def wrap_text(text: str, max_length: int, prefix: str):
    """
    Reformat text to a line limit, breaking on spaces

    :param text: Some text with spaces in it
    :param max_length: Max line length including the indent
    :param prefix: Prefix for each line
    :return: Wrapped text
    """
    rest = text
    result = []
    max_length -= len(prefix)
    while True:
        end = rest[:max_length].rfind(' ')
        if end == -1:
            assert max_length > len(rest), "not expecting very long lines with no spaces"
            result.append(prefix + rest)
            break
        else:
            result.append(prefix + rest[:end])
            rest = rest[end+1:]
    return '\n'.join(result)
