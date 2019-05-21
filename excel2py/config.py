"""
Command line args and configuration for exel2py

TODO: Command line args should override configuration values
TODO: Use config instead of hard coding

By Michael Grazebrook of Joined Up Finance Ltd
"""

import argparse
import os.path



# TODO: Mechanism for importing embedded functions which aren't from Excel
TOOLKIT_FUNCTIONS = {
    # This used to have client specific function names
}


def config(description):
    """
    Take values from the command line and/or config

    :param description:
    :return: Provides an object with these attributes:
        - spreadsheet, name of the file open in Excel
        - config: TODO: parsed contents of the config file
        - prefix: prefix for generated classes and file names
        - output: path to the main output file
        - gen_class_name: Name of the class containing generated code
        - class_name: Name of the stub class subclassing gen_class_name
    """
    args = _parse_args(description)
    _parse_config(args)
    return args


def _parse_args(description):
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument("spreadsheet",
                        help="Excel file name (including extension)")
    parser.add_argument("config_file", nargs='?',
                        help="yaml configuration file; '-' means create a template containing defaults")
    parser.add_argument("prefix", default="gen", nargs='?',
                        help="Default prefix name for files and classes")
    parser.add_argument("input_sheet", nargs='?', default="Inputs",
                        help="Name of the input sheet (default: Inputs)")
    parser.add_argument("result_sheet", nargs='?', default="Results",
                        help="Name of the result sheet (default: Results)")
    parser.add_argument(
        "output", nargs='?',
        help=(
            "Name of the main output file.\n"
            "The default is based on the prefix and spreadsheet path"))

    return parser.parse_args()


def _parse_config(args):
    """
    TODO: Parse a yaml config file. For now, hard code.

    Command line args override config values
    :param args:
    """
    args.ignore_sheets = ['Notes', ]
    args.valid_date_formats = ('dd/mm/yyyy',)

    _file_and_class_names(args)
    _inputs(args)
    _outputs(args)
    args.imports = "from excel2py.toolkit import *\n"
    args.globals = TOOLKIT_FUNCTIONS


def _outputs(args):
    """All formulae on an output sheet count as outputs. Constants do not.

    args.outputs can be used to include values which are not on an output sheet as well as (re-)defining names.

    I considered allowing column aliases here. They create a maintenance problem if the cell ranges change.
    However it seems best to permit them so the original spreadsheet can remain unchanged.
    """
    args.output_sheets = {
        'Results',
    }
    args.outputs = {
        'Results!C5': 'first_output',
        'Resluts!C6': 'second_output',
    }


def _inputs(args):
    args.input_sheets = {
        'Inputs'
    }
    args.inputs = {
        # TODO: Variables in pep8 style are non-standard and should be standardised
        'Inputs!D6': 'first_input',
        'Inputs!D7': 'second_input',
    }


def _file_and_class_names(args):
    """
    Update args with class_name, gen_class_name and possibly outputs

    e.g. myCalc.v1.xlsm => MyCalc, GenMyCalc, gen_mycalc.py
    :param args: argparse return value
    """
    base_name = args.spreadsheet if args.output is None else args.output

    path, name = os.path.split(base_name)
    assert '.' in name
    base = name[:name.find('.')]  # e.g. from mycalc.v1.xlsm, we just want mycalc

    args.class_name = base.capitalize()
    args.gen_class_name = args.prefix.capitalize() + args.class_name
    if args.output is None:
        args.output = os.path.join(path, f"{args.prefix}_{base}.py").lower()
