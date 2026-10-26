"""
main functions for any excel_enhancer application
"""
__version__ = "0.2.0"

import sys
import os
import time
import numpy as np
import pandas as pd

from pywintypes import com_error
import win32api
from numbers import Number
from functools import wraps
from contextlib import contextmanager
import logging
from datetime import datetime
from collections import OrderedDict
import unicodedata
import traceback
import threading as th
import importlib
import importlib.util
import pickle

import xlwings as xw
try:
    from xlwings import xlplatform
except ImportError:
    try:
        from xlwings import _xlwindows as xlplatform
    except ImportError:
        try:
            from xlwings import _xlmac as xlplatform
        except ImportError:
            xlplatform = None
if not hasattr(xw, 'xlplatform'):
    xw.xlplatform = xlplatform

from xlwings.udfs import udf_modules

from xlwings.conversion.standard import (BaseAccessor, ValueAccessor, ReadValueFromRangeStage, Ensure2DStage,
                                         CleanDataFromReadStage, TransposeStage, AdjustDimensionsStage,
                                         Pipeline, WriteValueToRangeStage, CleanDataForWriteStage)

from xlwings.conversion import Converter, PandasDataFrameConverter


# class keeping special values to standardize them on VBA functions results or interpretations
class Special:
    # https://github.com/xlwings/xlwings/issues/776#issuecomment-476183843
    # this needs to be something like sheet.range("A5").formula = '#N/A'
    # all errors https://support.microsoft.com/en-us/office/formulas-and-functions-294d9486-b332-48ed-b489-abe7d0f9eda9?ui=en-US&rs=en-US&ad=US#id0eaabaaa=errors

    # The #N/A error generally indicates that a formula can’t find what it’s been asked to look for.
    nan = empty = '#N/A'

    err_div = '#DIV/0!'

    # The top reason why the #NAME? error appears in your formula is because there is a typo in the formula name
    err_name = '#NAME?'

    err_null = '#NULL!'

    # if function contains numeric values that aren't valid
    err_num = '#NUM!'

    # The #REF! error shows when a formula refers to a cell that’s not valid
    err_ref = '#REF!'

    # #VALUE! There's something wrong with the way your formula is typed. Or,
    # there's something wrong with the cells you are referencing.
    err_value = '#VALUE!'

    error_dict = {-2146826281: np.inf, -2146826246: np.nan}
    row_sep = ';'  # separator for rows in excel equation


general_configuration = {}  # configuration to modify function modes in @xw.func i.e. {"async_mode": 'threading'}
custom_group_ops = {}  # custom functions used in VBAs


def find_attribute_in_object_tree(att, obj, addrs="", max_re=None):
    """
    helper function to find attributes in an object

    >>> find_attribute_in_object_tree("calculate", xw, max_re= 10)

    :param att: attribute to find
    :param obj: object to look on
    :param addrs: starting address. Leave as "" if search is starting.
    :param max_re: maximun deep in tree
    :return:
    """
    if max_re is not None and addrs.count(".") > max_re:
        return
    try:
        got = getattr(obj, att, None)
        if got is not None:
            print(f"found {att} in address {addrs}")
    except:
        pass

    for i in dir(obj):
        if not i.startswith("__"):  # ignore special methods
            if att.lower() in i.lower():
                print(f"found {i} in address {addrs}")
            try:
                got = getattr(obj, i, None)
                if got is not None:
                    find_attribute_in_object_tree(att, got, addrs + "." + i, max_re)
            except:
                continue


def find_similar_attribute_in_object_tree(att, obj, addrs="", max_re=None):
    """
    helper function to find similar attributes in an object

    >>> find_similar_attribute_in_object_tree("calculate", xw, max_re= 10)

    :param att: attribute to find
    :param obj: object to look on
    :param addrs: starting address. Leave as "" if search is starting.
    :param max_re: maximun deep in tree
    :return:
    """
    if max_re is not None and addrs.count(".") > max_re:
        return

    for i in dir(obj):
        if not i.startswith("__"):  # ignore special methods
            if att.lower() in i.lower():
                print(f"found {i} in address {addrs}")
            try:
                got = getattr(obj, i, None)
                if got is not None:
                    find_attribute_in_object_tree(att, got, addrs + "." + i, max_re)
            except:
                continue


def register(register_in=None, name=None):
    """
    register function in group. This is useful to make function available for a process or VBA that uses the group
    that is registered to.

    :param register_in: dictionary group. If None then register in default custom_group_ops in the module
    :param name: name to give in group
    :return: registered function
    """
    if register_in is None:
        register_in = custom_group_ops
    def to_register(f):
        """register function to """
        if name is None:
            register_in[f.__name__] = f
        else:
            register_in[name] = f
        return f
    return to_register


def strip_accents(s, throw=False):
    """helper function used to normalize text so that accents are ignored"""
    # https://stackoverflow.com/a/518232/5288758
    try:
        return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    except TypeError:
        if throw:
            raise
        return s


def col_name(n):
    # https://stackoverflow.com/a/23862195/5288758
    # replaces from xlwings.utils import col_name
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def convert_to_string(value):
    """
    convert value to string in the most simple way to be human readable
    """

    if value is None:
        return ""  # should not return anything

    # float case to handle e.g 0 instead of 0.0
    if isinstance(value, float) and value == int(value):
        return str(int(value))

    return str(value)


def helper_correct_to_str(values):
    """
    ensure lists hast all values as str

    :param values: 2D list
    :return: whether invalid values were found
    """
    found_invalid = False
    for row in values:
        for j, col in enumerate(row):
            if not isinstance(col, str):
                row[j] = str(col)
                found_invalid = True
    return found_invalid


# --------------------------- ACCESSORS AND CONVERTERS ---------------------------


class ReadValueFromRangeStageInfiniteAdjusted:
    """Converter to fit values to table data from infinite range. Finite ranges are untouched"""

    def __call__(self, c):
        if c.range:
            if is_range_infinite(c.range):
                fitted_range = fit_table(c.range)
                # raw_range changes None to '' automatically
                c.value = fitted_range.raw_value
            else:
                c.value = c.range.raw_value


class ConverterAdjustInfiniteRange:
    """Converter to fit infinite range to finite range using table data. Finite ranges are untouched"""

    def __call__(self, c):
        if c.range:
            if is_range_infinite(c.range):
                c.range = fit_table(c.range)


class EnsureDataIsWritable:

    def __call__(self, c):
        # value must be a list of lists
        if isinstance(c.value, (list, tuple)) and isinstance(c.value[0], (list, tuple)):
            found_invalid = helper_correct_to_str(c.value)
            if found_invalid:
                print("values on invalid output corrected")


class ValueAccessorFitted(ValueAccessor):
    """
    Accessor that ensures values of tables fitting infinite rows to table data
    so that data result data don't have many None values and take time to convert
    """

    @staticmethod
    def reader(options):
        return (
            BaseAccessor.reader(options)
            .append_stage(ConverterAdjustInfiniteRange())
            .append_stage(ReadValueFromRangeStage())
            .append_stage(Ensure2DStage())
            .append_stage(CleanDataFromReadStage(options))
            .append_stage(TransposeStage(), only_if=options.get('transpose', False))
            .append_stage(AdjustDimensionsStage(options))
        )

    @staticmethod
    def writer(options):
        return (
            Pipeline()
            #.append_stage(EnsureDataIsWritable())
            .prepend_stage(WriteValueToRangeStage(options))
            .prepend_stage(CleanDataForWriteStage())
            .prepend_stage(TransposeStage(), only_if=options.get('transpose', False))
            .prepend_stage(Ensure2DStage())
        )


ValueAccessorFitted.register(None)  # register as default


def is_range_infinite(rng):
    """determine if a range is infinite e.g A:C"""
    address = rng.address
    is_finite = any([str(i) in address for i in range(10)])
    return not is_finite


def range_bounding_box(rng):
    """
    get bounding box to fit contents of table

    :param rng: range
    :return: first_row, last_row, first_col, last_col
    """

    # https://github.com/xlwings/xlwings/pull/1177#issuecomment-558322707
    # x.sheet[row_range,column_range]
    # https://stackoverflow.com/a/53102362/5288758
    # x.sheet.range("B"+str(x.last_cell.row)).end("up").row  # fit the data
    last_row = rng.last_cell.row  # get last row
    sampled_rows = []
    last_column = 0  # counter for last column
    for col in (i.column for i in rng.columns):  # see columns
        last_column = col
        last_range_str = rng.sheet[0, last_column - 1].address.replace("1", str(last_row))
        sampled_rows.append(rng.sheet.range(last_range_str).end("up").row)
    last_row_sampled = max(sampled_rows)
    return rng.row - 1, last_row_sampled, rng.column - 1, last_column


def fit_table(rng):
    """
    adjust range to fit contents of table in the rows so that infinite ranges like A:B
    don't go to infinite but the actual last row with data e.g. A1:B500

    :param rng: range
    :return: fitted_range
    """
    first_row, last_row, first_col, last_col = range_bounding_box(rng)
    return rng.sheet[first_row:last_row, first_col:last_col]


class MyConverter(Converter):
    # TODO, this is a test

    @staticmethod
    def read_value(value, options):
        myoption = options.get('myoption', None)
        return_value = value  # Implement your conversion here
        return return_value

    @staticmethod
    def write_value(value, options):
        myoption = options.get('myoption', None)
        return_value = value  # Implement your conversion here
        return return_value

    @classmethod
    def reader(self, options):
        base_reader = self.base_reader(options)
        # modify stage of reading default
        base_reader.insert(0, ConverterAdjustInfiniteRange())
        return (
            base_reader.append_stage(self.FromValueStage(self.read_value, options))
        )


# helper converters


def pandas_dataframe_converter(value, *args, **kwargs):
    """ pandas converter from xlwings for particular conversions when needed"""
    if isinstance(value, pd.DataFrame):
        return value
    if args and not kwargs:
        kwargs = args[0]
    elif args:
        raise ValueError(f"cannot deal with options args {args} and kwargs {kwargs}")
    return PandasDataFrameConverter.read_value(value, kwargs)


# --------------------------- DEBUGGERS ---------------------------


def script_path():
    """
    get the script path of the file that starts the program
    """
    current_script = os.path.realpath(sys.argv[0])  # script directory
    return os.path.splitext(current_script)[0]  # name, ext


class OperationDebugger(logging.Logger):
    _is_setting_up = False

    def __init__(self, file_name=None, level=logging.DEBUG, stream=None, formatter=None):
        """
        create a file debugger to log relevant info, errors and debug data to a log.

        :param file_name: name of debugger and file
        :param level: logging level. logging.DEBUG by default
        :param stream: use a stream to print to. None by default
        :return:
        """
        self._is_setting_up = True
        self._file_created = False

        file_timestamp = '%Y-%m-%d %H:%M:%S'
        if formatter is None:
            formatter = logging.Formatter(fmt='%(asctime)s# %(levelname)-8s# %(message)s',
                                          datefmt=file_timestamp)
        now = datetime.now()
        if file_name is None:
            file_name = '{name}-{timestamp}.log'.format(timestamp=now.strftime(file_timestamp).replace(":", "_"),
                                                        name=script_path())
        self.formatter = formatter

        #logger = logging.getLogger(name)
        #logger.setLevel(level)
        super().__init__(file_name, level)

        if stream is not None:
            screen_handler = logging.StreamHandler(stream=stream)
            screen_handler.setFormatter(formatter)
            self.addHandler(screen_handler)

        self._is_setting_up = False

    def __getattribute__(self, item):
        # register this only when logging starts
        if not object.__getattribute__(self, "_is_setting_up"):
            if not object.__getattribute__(self, "_file_created"):  # this must run once
                self._is_setting_up = True
                file_name = object.__getattribute__(self, "name")
                formatter = object.__getattribute__(self, "formatter")
                addHandler = object.__getattribute__(self, "addHandler")
                handler = logging.FileHandler(file_name, mode='w')
                handler.setFormatter(formatter)
                addHandler(handler)
                self._file_created = True
                self._is_setting_up = False
        return super(OperationDebugger, self).__getattribute__(item)


file_debugger = OperationDebugger()


def exception_to_stream(e, stream_handler=None, template=None):
    """
    prints Exception message to stream much like traceback.format_exception or traceback.print_exception

    :param e: Exception
    :param stream_handler:
    :param template:
    :return:
    """
    # gotten from book "The Python 3 Standard Library by Example: Pyth 3 Stan Libr Exam _2"
    #template = template or "{filename:<23}:{linenum}:{funcname}:\n{source}"
    template = template or "File \"{filename:<23}\", line {linenum}, in {funcname}\n{source}"
    lines = []
    for tb_info in traceback.extract_tb(e.__traceback__):
        filename, linenum, funcname, source = tb_info
        filename = os.path.basename(filename)
        if funcname != '<module>':
            funcname += '()'
        msg = template.format(filename=filename, linenum=linenum, funcname=funcname, source=source)
        lines.append(msg)

    if stream_handler:
        stream_handler("\n".join(lines))
    return lines


serializers = {}


class TestSerializer:
    """
    Class serializer to convert data into natural python language where it is possible.
        If not tries to generate code to load serialized data into python
    """

    @classmethod
    def get_serializer(cls, type):
        type = cls.normalize_type(type)
        return serializers.get(type, serializers.get(None))

    @classmethod
    def register(cls, *types):
        for type in types:
            type = cls.normalize_type(type)
            serializers[type] = cls

    @classmethod
    def serializer(cls, data):
        """
        serializes legacy units of data or sub-data in a recursive way so as to return
        natural python str code to print of the data to be interpreted by the python interpreter
        """
        # catches all python legacy types
        # sanitize and return as python natural language
        if data is None:
            return "None"
        elif isinstance(data, bool):
            return "{}".format(data)
        elif isinstance(data, list):
            return "[{}]".format(", ".join(f"{cls.serialize(i)}" for i in data))
        elif isinstance(data, tuple):
            return "({})".format(", ".join(f"{cls.serialize(i)}" for i in data) + ("," if len(data) == 1 else ""))
        elif isinstance(data, set):
            return "{{}}".format(", ".join(f"{cls.serialize(i)}" for i in data))
        elif isinstance(data, dict):
            return "{{}}".format(", ".join(f"{cls.serialize(k)}: {cls.serialize(v)}" for k, v in data.items()))
        elif isinstance(data, str):
            return repr(data)
        else:
            return cls.serialize(data)

    @classmethod
    def deserializer(cls, data):
        """
        deserializes unit of data. which in this case is itself as it must be natural python code
        interpreted by the python interpreter.

        sub-classes must implement the deserializer so that python reconstruct serialized data into
        a python object or readable data
        """
        return data

    @classmethod
    def normalize_type(cls, type):
        """
        normalizes type of data to get str name
        """
        if type is not None and not isinstance(type, str):
            type = type.__name__
        return type

    @classmethod
    def prepare(cls, serialized, type):
        """
        pos-process serialized intended to help children implement serialize method in a clean way
        """
        if isinstance(serialized, str):
            return serialized
        type = cls.normalize_type(type)
        return f"{TestSerializer.__name__}.deserialize({repr(serialized)}, {repr(type)})"

    @classmethod
    def serialize(cls, data):
        """
        solves serialization for data and must return natural python str code to print
        """
        data_type = type(data)
        serializer = cls.get_serializer(data_type)
        return cls.prepare(serializer.serializer(data), data_type)

    @classmethod
    def deserialize(cls, data, type=None):
        """
        deserializes data and returns object
        """
        deserializer = cls.get_serializer(type)
        return deserializer.deserializer(data)


# register supported legacy python types
TestSerializer.register(str, tuple, list, dict, set, type(None), bool)


class LastSerializer(TestSerializer):
    """last serializer to user in case none matches data type"""

    @classmethod
    def serializer(cls, data):
        return pickle.dumps(data, 0)  # most human readable. This returns byte so it is processed

    @classmethod
    def deserializer(cls, data):
        return pickle.loads(data)


# register general serializer for any type of data not cached at the end of the serialization
LastSerializer.register(None)


class DataFrameSerializer(TestSerializer):
    """serializer for Dataframes"""

    @classmethod
    def serializer(cls, data):
        columns = list(data.columns)
        body = data.values.tolist()
        #return f"{TestSerializer.__name__}.deserialize({repr((columns, body))}, 'DataFrame')"
        return columns, body

    @classmethod
    def deserializer(cls, data):
        columns, body = data
        return pd.DataFrame(body, columns=columns)


DataFrameSerializer.register(pd.DataFrame)


def inheritors(klass):
    """helper function to get inheritors from a base class"""
    # https://stackoverflow.com/a/5883218/5288758
    subclasses = set()
    work = [klass]
    while work:
        parent = work.pop()
        for child in parent.__subclasses__():
            if child not in subclasses:
                subclasses.add(child)
                work.append(child)
    return subclasses


class UnitTestCreator:
    """
    Class helper to manage unittest creations of main and custom functions or UDFs
    """
    def __init__(self):
        self.tests = {}
        self.indent = "    "

    def _get_indents(self):
        """creates basic indents i0, i1, i2, i3"""
        return {f"i{i}": self.indent * i for i in range(4)}

    def collect_test(self, func, ret, args, kwargs):
        """
        collect test to functions that will be printed for unittest

        :param func: function
        :param ret: function result
        :param args: arguments
        :param kwargs: keyword arguments
        :return:
        """
        name = func.__name__
        cache_key = self.get_cache_key(func, (ret, args, kwargs))

        if name in self.tests:
            test_cases = self.tests[name]
            if cache_key not in test_cases:
                # add new test case
                test_cases[cache_key] = (func, ret, args, kwargs)
        else:
            # create new unittest
            test_cases = {cache_key: (func, ret, args, kwargs)}
            self.tests[name] = test_cases

    def print_unittest(self, test_name):
        """
        convenient function to format and print requested test

        :param test_name: collected test of function name
        :return: formatted test to be placed in an unittest with extension .py
        """
        formatted = self.format_unittest(test_name)
        print(formatted)
        return formatted

    def format_unittest(self, test_name):
        """

        :param test_name: collected test of function name
        :return: formatted test to be placed in an unittest with extension .py
        """
        if not isinstance(test_name, str):
            test_name = test_name.__name__
        test_cases = self.tests[test_name]
        template_unittest = """
import unittest
{module_imports}


class {test_name}(unittest.TestCase):
{test_cases}

if __name__ == '__main__':
{i1}unittest.main()
"""
        # get default formatting
        formats = self._get_indents()
        # unittest name in camel case
        formats["test_name"] = "MyTestFor"+"".join([i.capitalize() for i in test_name.split("_")])

        # add test cases
        lines = []
        functions = set()
        for (i, (cached_key, (func, ret, args, kwargs))) in enumerate(test_cases.items()):
            test_case_name = f"test_{func.__name__}{i}"
            functions.add(("excel_enhancer", func.__name__))
            formatted = self.format_test_case(test_case_name, func, ret, args, kwargs)
            lines.append(formatted)

        i1 = formats["i1"]
        formats["test_cases"] = "".join(["\n".join([f"{i1}"+f"{j}" for j in i.split("\n")]) for i in lines])

        # add imports
        main_imports = [("core", UnitTestCreator.__name__),
                        ("core", TestSerializer.__name__),
                        ("core", "debug")]
        imports = main_imports + list(functions)
        formats["module_imports"] = "\n".join([f"from {mod} import {f}" for mod, f in imports])

        # additional lines
        formats["module_imports"] += "\n\ndebug.testing = True"

        # format template
        formatted = template_unittest.format(**formats)
        return formatted

    def format_test_case(self, case_name, func, ret, args, kwargs):
        """
        helper function to format individual test case method

        :param case_name: name of the test case
        :param func: function
        :param ret: function result
        :param args: arguments
        :param kwargs: keyword arguments
        :return:
        """
        template_test_case = """
def {case_name}(self):
{arg_serializations}
{kwarg_serializations}
{ret_serializations}
{function_execution}
{function_test}
"""
        # get default formatting
        formats = self._get_indents()
        i1 = formats["i1"]
        # add case name
        formats["case_name"] = case_name

        # for args
        if args:
            lines_arg_serializations = []
            names_arg = []
            for i, arg in enumerate(args):
                ser_arg = self.serialize(arg)
                arg_name = f"arg{i}"
                line = f"{i1}{arg_name} = {ser_arg}"
                names_arg.append(arg_name)
                lines_arg_serializations.append(line)

            formats["arg_serializations"] = f"{i1}# get arguments\n"+"\n".join(lines_arg_serializations)
        else:
            formats["arg_serializations"] = ""
            names_arg = []

        # for kwargs
        if kwargs:
            lines_kwarg_serializations = []
            names_kwarg = []
            for i, (key, value) in enumerate(kwargs.items()):
                ser_val = self.serialize(value)
                line = f"{i1}{key} = {ser_val}"
                names_kwarg.append(key)
                lines_kwarg_serializations.append(line)

            formats["kwarg_serializations"] = f"{i1}# get key word arguments\n"+"\n".join(lines_kwarg_serializations)
        else:
            formats["kwarg_serializations"] = ""
            names_kwarg = []

        # for ret
        ret_val = self.serialize(ret)
        line = f"{i1}ret_cmp = {ret_val}"
        formats["ret_serializations"] = f"{i1}# get expected result\n"+line

        # function_execution
        str_args = ", ".join(names_arg)
        str_kwargs = (", "+", ".join(names_kwarg)) if names_kwarg else ""

        line = f"{i1}ret = {func.__name__}({str_args}{str_kwargs})"
        formats["function_execution"] = f"{i1}# evaluate function\n"+line

        # function_test
        #line = f"{i1}# compare expected an result\n{i1}self.assertEqual(ret_cmp, ret)"
        line = f"{i1}# compare expected an result\n{i1}self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))"
        formats["function_test"] = line

        # format template
        formatted = template_test_case.format(**formats)
        return formatted

    @classmethod
    def serialize(cls, data):
        """
        serializer to insert into formatted code

        :param data:
        :return:
        """
        return TestSerializer.serialize(data)

    @classmethod
    def deserialize(cls, data, type=None):
        """
        deserializer to insert into formatted code

        :param data:
        :return:
        """
        return TestSerializer.deserialize(data, type)

    @classmethod
    def get_cache_key(cls, func, args):
        """get cache key to know if already used"""
        return func.__name__, str([i.to_string() if isinstance(i, pd.DataFrame) else i for i in args])


test_creator = UnitTestCreator()


class DebuggerWatcher:
    """
    production debugger implemented for Watcher. This debugger can decorate UDF functions
    to accept an additional parameter considered the steps to obtain the result.

    # consider this function
    @xw.func
    @xw.arg('df', pd.DataFrame)
    @xw.ret(expand='table', index=False)
    def my_func(df):
        result = df["MY COLUMN"].sum()
        return result

    # now if we are going to add support for debugging then we just decorate with debug.get_steps
    # and we need to add in any place of the workbook the function '=xl_set_watcher("steps")'
    @debug.get_steps()
    @xw.func
    @xw.arg('df', pd.DataFrame)
    @xw.ret(expand='table', index=False)
    def my_debugged_func(df):
        result = df["MY COLUMN"].sum()
        steps = [df]  # here we just show where the result came from
        return result, steps

    # we can also add a back-end debugger for the terminal as debug.debug to show the results of the functions
    @debug.debug(types=True)
    @debug.get_steps()
    @xw.func
    @xw.arg('df', pd.DataFrame)
    @xw.ret(expand='table', index=False)
    def my_debugged_func(df):
        result = df["MY COLUMN"].sum()
        steps = [df]  # here we just show where the result came from
        return result, steps

    """
    def __init__(self):
        self.active = True
        self.collect_tests = True
        self.testing = False
        self.debug_space = {}
        self._cell_formulas = {}

    def _print(self, data):
        if self.active:
            print(data)

    def get_steps(self):
        """this decorator should be the first"""
        def outerwrap(f):
            @ wraps(f)
            def wrapper(*args, **kwargs):
                func_info = wrapper.__xlfunc__
                args_info = func_info['args']
                ret_info = func_info['ret']
                is_dynamic_array = ret_info['options'].get('expand')

                # for testing:
                if self.testing:
                    # evaluate the UDF
                    # TODO simulate excel behaviour i.e when error is thrown then return what is sent to cell
                    output = f(*args, **kwargs)
                    return output

                # get the application
                if app_present:
                    xl_app = args[app_index]
                else:
                    args = list(args)
                    xl_app = args.pop(app_index)

                # get the caller
                tries = 0
                while True:
                    try:
                        # FIXME, not working for threads, so passed ranged in previous stage
                        if isinstance(xl_app, xw.Range):
                            xw_caller = xl_app
                        else:
                            caller = xl_app.caller
                            xw_caller = xw.Range(impl=xw.xlplatform.Range(xl=caller))
                        break
                    except AttributeError:
                        # wait for caller in threading
                        tries += 1
                        time.sleep(0.1)
                        if tries > 5:
                            raise
                #xw_caller = xw.Range(impl=xw.xlplatform.Range(xl=caller))

                # get the cell of the evaluating formula
                # key = wb.selection.api.FormulaArray
                # xl_app.selection.address
                # xl_app.caller.Address  # address of the actual calling formula
                # FIXME. gets bad addresses
                # (xw_caller.sheet.book.app.pid, xw_caller.sheet.book.name, xw_caller.sheet.name, xw_caller.address)
                # F"'[{xw_caller.sheet.book.name}]{xw_caller.sheet.name}'!{xw_caller.address}"
                range_selection = xw_caller.get_address(True, False, external=True)
                equation_address = F"'[{xw_caller.sheet.book.name}]{xw_caller.sheet.name}'!{xw_caller.address}"
                if range_selection.replace("$", "").replace("'", "") != equation_address.replace("$", "").replace("'", ""):
                    print(F"got bad equation address {range_selection.replace('$', '')} (gotten) != {equation_address.replace('$', '')} (constructed)")
                    pass

                # BAD as it is anywhere in array -> range_selection[:range_selection.find("!")+1]+xl_app.selection.address
                formula_cell = range_selection.split(":")[0]
                # evaluate the UDF
                # TODO catch error that will be printed on cell to feed to test_creator
                output = f(*args, **kwargs)
                if self.collect_tests:
                    test_creator.collect_test(f, output, args, kwargs)

                try:
                    ret, steps = output
                except ValueError:  # no steps fround
                    ret, steps = output, []

                # get the formula that is being evaluated
                try:
                    #formula = xl_app.selection.FormulaArray  # very bad, not evalating formula but the one in selection
                    formula = xw_caller.sheet[formula_cell].formula
                    if func_info["name"] not in formula:
                        print(f"evaluating '{func_info['name']}' but not in formula '{formula}'")

                    file_debugger.debug(F"evaluating '{formula}'")
                    # check if this formula is cached
                    is_cached = formula in self.debug_space
                except com_error as e:
                    # e.excepinfo[5] == -2147220464:
                    # import win32api
                    # win32api.FormatMessage(-2147352565)
                    # getting steps unsuccessful
                    file_debugger.warning(F"could not get formula for debugging: {func_info}")
                    #traceback.format_exception(e, e, e.__traceback__)
                    exception_to_stream(e, stream_handler=file_debugger.warning)
                    return ret

                # collect steps and debugging data
                if is_cached:  # maintain debugging data

                    # get cache
                    formula_cache = self.debug_space[formula]

                    # replace new cache values
                    # ret, steps should be to its specific nested function = func_info["name"]
                    formula_cache[0][func_info["name"]] = (ret, steps)

                    # keep track of formulas in each cell
                    if formula_cell in self._cell_formulas:  # check cell already has formula tracked
                        previous_formula = self._cell_formulas[formula_cell]

                        # keep track of cells that use the same formula
                        if previous_formula != formula and previous_formula in self.debug_space:

                            # formula changed thus maintain
                            previous_formula_cache = self.debug_space[previous_formula]

                            # clean current cell from previous_formula
                            previous_formula_references = previous_formula_cache[1]
                            try:
                                previous_formula_references.remove(formula_cell)
                            except KeyError:
                                pass

                            try:
                                # cleanup if previous formula reaches 0 references
                                # no cell should have the formula anymore
                                if not previous_formula_references:
                                    del self.debug_space[previous_formula]
                            except KeyError:
                                pass

                    else:  # there was no formula previously in this cell
                        # cell with new formula
                        self._cell_formulas[formula_cell] = formula

                    # add this cell to the formula track even if it exists
                    # ensure new formula is accounted for
                    formula_cache[1].add(formula_cell)
                else:  # create debugging data
                    # keep track of cell that contains this formula
                    self._cell_formulas[formula_cell] = formula

                    # create new cached data for formula
                    formula_references = set((formula_cell,))
                    data_cache = OrderedDict()
                    data_cache[func_info["name"]] = (ret, steps)
                    formula_cache = self.debug_space[formula] = [data_cache, formula_references]

                # normalize to xl result
                #xl_result = xw.conversion.write(ret, None, ret_info['options'])
#
                #if is_dynamic_array:
                #    current_size = (caller.Rows.Count, caller.Columns.Count)
                #    result_size = (1, 1)
                #    if type(xl_result) is list:
                #        result_height = len(xl_result)
                #        result_width = result_height and len(xl_result[0])
                #        result_size = (max(1, result_height), max(1, result_width))
                #    if current_size != result_size:
                #        pass
                        #from .server import add_idle_task
                        #add_idle_task(DelayedResizeDynamicArrayFormula(
                        #    Range(impl=xlplatform.Range(xl=caller)).resize(*result_size),
                        #    caller,
                        #    current_size[0] > result_size[0] or current_size[1] > result_size[1]
                        #))
                # register steps
                return ret  # return original data
            try:
                # share __xlfunc__ of original function  to wrapper
                func_info = f.__xlfunc__
                wrapper.__xlfunc__ = func_info

                # determine if Application is present
                args_info = func_info['args']
                applications_tests = [i["vba"] == "Application" for i in args_info]
                app_present = any(applications_tests)
                if app_present:
                    app_index = applications_tests.index(True)
                else:
                    # check if there is *args
                    varargs_i = [i for i, v in enumerate(args_info) if v["vararg"] is True]
                    if varargs_i:
                        # fix vararg new position
                        assert len(varargs_i) == 1
                        app_index = varargs_i[0]  # this is where Application will be placed
                        # update current arg in actual position to new position after Application
                        args_info[app_index]["pos"] += 1
                        args_info[app_index]["doc"] += F'Positional argument {args_info[app_index]["pos"]}'
                    else:
                        app_index = len(args_info)  # -1  # FIXME not always -1

                    # add application as last value last variable
                    app_name = 'xl_app'
                    names = [i["name"] for i in args_info]
                    while app_name in names:
                        app_name += "1"  # update name until there is no conflict

                    arg_app = {'name': app_name,
                               'pos': app_index,
                               'vba': "Application",
                               'doc': F'Positional argument {app_index}',
                               'vararg': False,
                               'options': {}, 'optional': None}
                    args_info.insert(app_index, arg_app)
                    func_info["argmap"][app_name] = arg_app

                    this_copy = func_info.copy()  # copy to restore options
                    # this operation deletes initial options
                    xw.arg(app_name, vba='Application')(wrapper)

                    # restore options
                    for key, val in wrapper.__xlfunc__.items():
                        old_val = this_copy[key]
                        if old_val != val:
                            wrapper.__xlfunc__[key] = old_val
                            #print(f"restored {func_info['name']} in {key} with {old_val}")
            except AttributeError:
                pass
            return wrapper
        return outerwrap

    def debug(self, types=False):
        def outerwrap(f):
            @ wraps(f)
            def wrapper(*args, **kwargs):
                if self.active:
                    print(F'calling function "{f.__name__}" with arguments')
                if self.active:
                    if types:
                        print('args: ('+", ".join([repr(i)+str(type(i)) for i in args])+')')
                        print('kwargs: {'+", ".join([repr(k)+": "+repr(i)+str(type(i)) for k, i in kwargs.items()])+'}')
                    else:
                        print(F'args: {args}')
                        print(F'kwargs: {kwargs}')
                data = f(*args, **kwargs)
                if self.active:
                    if types:
                        print(F'got data: {data}{type(data)}')
                    else:
                        print(F'got data: {data}')
                return data
            return wrapper
        return outerwrap

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    @contextmanager
    def debug_space(self, *args, **kwds):
        """
        this context is intended to be used inside the debugged function

        WARNING: still not clear usage!!

        :param args:
        :param kwds:
        :return:
        """
        # TODO recognize debugged function
        try:
            yield  # automatically collect steps in the context
        finally:
            # clean up
            pass


debugger = DebuggerWatcher()


# --------------------------- WATCHER IMPLEMENTATION ---------------------------


class Watcher:
    """
    Watcher to implement and show DebuggerWatcher
    """

    def __init__(self):
        self.caller = None
        self.watchers = {}
        self.watcher_names = ["address", "plain", "mimic", "print", "steps"]
        self.thread = None

    @xw.func
    @xw.arg('name', doc='Watcher name')
    def set_watcher(self, name="plain"):
        self.caller = wb = xw.Book.caller()
        addrs = wb.selection.get_address(True, False, external=True)  # xl_app.Caller.Address
        name = name.lower()
        if name in self.watcher_names:
            self.watchers[addrs] = name
            return addrs
        return F"watcher_name {name} not recognized. use {self.watcher_names}"


    @xw.func
    def del_watcher(self):
        wb = xw.Book.caller()
        addrs = wb.selection.get_address(True, False, external=True)  # xl_app.Caller.Address
        del self.watchers[addrs]

    def watch(self):
        my_thread = th.Thread(target=self._watch)
        #my_thread.daemon = False
        my_thread.start()
        self.thread = my_thread

    def _watch(self):
        """
        helper function to watch constantly for user selections where data could be being debugged
        :return:
        """
        old_sel = None
        # cache data for cleaning
        val_olds, steps_olds = {}, {}
        while True:
            # FIXME this consumes too many resources
            time.sleep(0.5)
            if debugger.active:
                try:
                    try:
                        wb = xw.books.active
                    except AttributeError:
                        # no more active books
                        print("no more active books. exiting...")
                        break

                    #wb = xw.Book.caller()
                    if wb is not None:
                        #sel = wb.selection.address
                        sel = wb.selection.get_address(True, False, external=True)
                        if sel != old_sel:  # update on change
                            old_sel = sel
                            for address, (watcher_name, opts) in self.watchers.items():
                                # TODO create lazy mode do not update if not in sheet of debugging

                                if watcher_name == "print":
                                    print(F"{sel} = {xw.Range(sel).raw_value}")
                                elif watcher_name == "plain":
                                    xw.Range(address).value = F"{sel} = {xw.Range(sel).value}"
                                elif watcher_name == "address":
                                    xw.Range(address).value = F"{sel}"
                                elif watcher_name == "mimic":

                                    # get selected value
                                    # fixed: determine orientation when data is one dimension
                                    selection = xw.Range(sel).options(ndim=2).value

                                    # clear previous data
                                    val_old = val_olds.get(address)
                                    #xw.Range(address).expand().clear()  # dangerous, this can delete near values
                                    if val_old is not None:
                                        xw.Range(address).value = np.zeros_like(val_old, dtype="<U1")

                                    # set new data
                                    xw.Range(address).value = selection
                                    val_olds[address] = selection
                                elif watcher_name == "steps":
                                    # FIXME. gets bad addresses
                                    # TODO implement using data from DebuggerWatcher
                                    # get selected value of function xw.Range("C2").api.Validation.Formula1[1:]
                                    # see posible values of Validation https://stackoverflow.com/a/54092842/5288758
                                    # fixed: determine orientation when data is one dimension
                                    #val = xw.Range(sel).options(ndim=2).value
                                    #formula = wb.selection.api.FormulaArray  # dangerous, do not produce formula the same as in single cell
                                    formula = wb.selection.sheet[wb.selection.address.split(":")[0]].formula

                                    # remember last shown formula at address
                                    steps_old, last_formula_shown, all_menu_options = steps_olds.get(address, (None, None, {}))

                                    cell_formula = xw.Range(address)
                                    col, row = cell_formula.column - 1, cell_formula.row - 1
                                    cell_validation = cell_formula.sheet[row, col + 1]  # cell to show debugged function
                                    cell_steps = cell_formula.sheet[row, col + 2]  # cell to show steps of debugged formula
                                    update_gui = False
                                    try:
                                        # CASE: debug data exists in current formula
                                        data_cache, refs = debugger.debug_space[formula]
                                        menu_options = all_menu_options.get(formula, {})
                                        formula_cache = formula

                                        # set gui if a new formula is gotten
                                        if formula != last_formula_shown or not menu_options.get("is_set"):
                                            # CASE: current formula is different to shown or not gui set

                                            #row_shift = None
                                            #for row_shift, func_name in enumerate(data_cache.keys(), 1):
                                            #    cell_formula.sheet[row, col + 1 + row_shift].value = func_name
                                            #if row_shift is not None:
                                            #choices_range = cell_formula.sheet[row, col + 2: col + 2 + row_shift]
                                            #choices_range_form = "="+choices_range.address

                                            available_funcs = list(data_cache.keys())
                                            choices_range_form = Special.row_sep.join(available_funcs)
                                            create_validation(cell_validation, choices_range_form)

                                            # choose previous selected function or first
                                            selected_func = menu_options.get("selected_func", available_funcs[0])
                                            if selected_func is not None:
                                                cell_validation.value = selected_func

                                            ret, steps = data_cache[selected_func]

                                            if steps:
                                                create_validation(cell_steps, Special.row_sep.join([str(i) for i in range(len(steps))]))
                                                selected_step = menu_options.get("selected_step", "0")
                                                cell_steps.value = selected_step

                                            menu_options["is_set"] = True
                                            update_gui = True
                                            # unregister last
                                            old_menu_options = all_menu_options.get(last_formula_shown, {})
                                            old_menu_options["is_set"] = False

                                    except KeyError as e:
                                        # CASE: no data for current formula, check previous gui options for updating
                                        menu_options = all_menu_options.get(last_formula_shown, {})

                                        #print(e)
                                        if not menu_options.get("is_set"):
                                            continue

                                        data_cache, refs = debugger.debug_space[last_formula_shown]
                                        formula_cache = last_formula_shown

                                    # CASE: same formula as before, gui already set # TODO this evals always, find a way to be triggered when changed
                                    selected_func = cell_validation.value
                                    selected_step = cell_steps.value
                                    # check controls have changed
                                    selected_func_old = menu_options.get("selected_func")
                                    selected_step_old = menu_options.get("selected_step")
                                    if selected_func != selected_func_old or selected_step != selected_step_old:
                                        # change in options so update
                                        update_gui = True

                                    if not update_gui and steps_old is None:
                                        # no changes to debug table so do not update
                                        continue

                                    # get cache
                                    ret, steps = data_cache[selected_func]

                                    if not steps:
                                        debug_table = None
                                    else:
                                        debug_table = steps[int(selected_step)]

                                    # check change in debug table
                                    if steps_old is not debug_table:
                                        update_gui = True

                                    if not update_gui:
                                        continue

                                    # CASE: Update gui

                                    # creating last handles
                                    cell_debug_table = cell_formula.sheet[row + 1, col]  # cell to show debug table
                                    cell_debug_formula = cell_formula.sheet[row, col + 3]  # cell to show debugged formula

                                    # clean previous data
                                    if steps_old is not None:
                                        if opts.get("expand"):
                                            # expand table potentially deleting more
                                            cell_debug_table.expand().clear()
                                        else:
                                            # we can use expand but it could potentially delete more
                                            if isinstance(steps_old, pd.DataFrame):
                                                # consider index
                                                cache = steps_old.copy()
                                                cache[:] = ""
                                                cache.columns = np.zeros_like(cache.columns, dtype="<U1")
                                                cache.index = np.zeros_like(cache.index, dtype="<U1")
                                                cell_debug_table.value = cache
                                            else:
                                                cell_debug_table.value = np.zeros_like(steps_old, dtype="<U1")
                                    elif opts.get("expand"):  # clean anyway any data in there
                                        # expand table potentially deleting more
                                        cell_debug_table.expand().clear()

                                    # set new data
                                    # DONE support more than one step
                                    if debug_table is not None:
                                        cell_debug_table.value = debug_table  # show debug table

                                    # maintain updated
                                    cell_debug_formula.value = "'" + formula_cache  # show formula
                                    menu_options["selected_func"] = selected_func
                                    menu_options["selected_step"] = selected_step
                                    all_menu_options[formula_cache] = menu_options
                                    steps_olds[address] = debug_table, formula_cache, all_menu_options

                except Exception as e:
                    print("Exception in watcher:")
                    print("-"*60)
                    traceback.print_exc(file=sys.stdout)
                    print("-"*60)

watcher = Watcher()

# register functions
#set_watcher = debug.debug(types=True)(watcher.set_watcher)
#del_watcher = watcher.del_watcher

# place this at the end of the scripts; when the server is started
#debug.active = False  # deactivate debugger
#watcher.watch()

@xw.func
@xw.arg('name', doc='Watcher name')
@xw.arg('options', doc='options to pass to the specific watcher')
@xw.arg('xl_app', vba='Application')
def xf_set_watcher(name="address", options=None, xl_app=None):

    # get the caller
    caller = xl_app.caller
    xw_caller = xw.Range(impl=xw.xlplatform.Range(xl=caller))

    # get the cell of the evaluating formula
    # key = wb.selection.api.FormulaArray
    # xl_app.selection.address
    # xl_app.caller.Address  # address of the actual calling formula
    # FIXED. gets bad addresses
    # (xw_caller.sheet.book.app.pid, xw_caller.sheet.book.name, xw_caller.sheet.name, xw_caller.address)
    # F"'[{xw_caller.sheet.book.name}]{xw_caller.sheet.name}'!{xw_caller.address}"
    range_selection = xw_caller.get_address(True, False, external=True)
    equation_address = F"'[{xw_caller.sheet.book.name}]{xw_caller.sheet.name}'!{xw_caller.address}"
    if range_selection.replace("$", "").replace("'", "") != equation_address.replace("$", "").replace("'", ""):
        print(F"IN xl_set_watcher got bad equation address {range_selection.replace('$', '')} (gotten) != {equation_address.replace('$', '')} (constructed)")

    #addrs = wb.selection.get_address(True, False, external=True)  # xl_app.Caller.Address
#
    #if addrs.replace("$", "").replace("'", "") != equation_address.replace("$", "").replace("'", ""):
    #    print(F"IN xl_set_watcher got bad equation address {addrs.replace('$', '')} (gotten selection) != {equation_address.replace('$', '')} (constructed)")
    #    pass

    name = name.lower().strip()
    if name in watcher.watcher_names:
        _options = {}
        if options is not None:
            if isinstance(options, str):
                for opt in options.split(","):
                    _options[opt.strip()] = True
            else:
                if len(options) % 2:
                    raise ValueError(F"options must be series to form key pairs. got {options} and '{options[-1]}' has no pair")
                # convert series to key pair
                _options.update(zip(options[::2], options[1::2]))
        watcher.watchers[range_selection] = name, _options
        return range_selection
    raise ValueError(F"watcher name '{name}' not recognized. use {watcher.watcher_names}")


@xw.func
@xw.arg('xl_app', vba='Application')
def xf_del_watcher(xl_app=None):

    # get the caller
    caller = xl_app.caller
    xw_caller = xw.Range(impl=xw.xlplatform.Range(xl=caller))

    # get equation where call was originated
    addrs = xw_caller.get_address(True, False, external=True)
    del watcher.watchers[addrs]


# --------------------------- MODULE FACTORY ---------------------------


def publish_functions(module, destine=None):
    """
    make all the module function available in globals

    :param module: module to get functions from
    :param destine: dictionary to publish functions
    :return: names of the published functions
    """

    # https://stackoverflow.com/a/43059528/5288758
    # get a handle on the module
    mdl = module  # importlib.import_module(module)

    # is there an __all__?  if so respect it
    if "__all__" in mdl.__dict__:
        names = mdl.__dict__["__all__"]
    else:
        # otherwise we import all names that don't begin with _
        names = [x for x in mdl.__dict__ if not x.startswith("_")]

    # now drag them in
    if destine is None:
        destine = globals()
    destine.update({k: getattr(mdl, k) for k in names})
    return names


def get_module_name(module):
    """
    get normalized module name

    :param module: module name, module or path to module file
    :return: module name
    """
    if isinstance(module, str):
        if os.path.isfile(module):
            path = module
            # path provided
            root, ext = os.path.splitext(path)
            root, basename = os.path.split(root)
        else:
            basename = module  # name of the module provided
    else:
        basename = module.__name__
    return basename


def load_module(module, reload=None):
    """
    import and reload from custom external modules

    :param module: module, module name or path to module file
    :param reload:
    :return:
    """
    # FIXED replace with a custom loader for user to externally decide modules to be imported and used dynamically
    if isinstance(module, str):
        if os.path.isfile(module):
            path = module
            # path provided
            root, ext = os.path.splitext(path)
            root, basename = os.path.split(root)
            spec = importlib.util.spec_from_file_location(basename, path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            module_name = module.__name__

            # register module in sys only if reload or no registered before
            if reload or module_name not in sys.modules:
                sys.modules[module_name] = module
        else:
            module_name = module  # name of the module provided
            # load module so that we have it present for reload
            module = importlib.import_module(module_name)
        if reload:
            importlib.reload(module)  # FIX ME, does not work if module was loaded from path and path not in sys.path
    return module


def reload_functions(module):
    """
    load module on demand and update namespace

    # equivalent to: from excel_enhancer_xf import *
    >>> reload_functions("udf_json_support")
    >>> reload_functions("./extensions/udf_json_support.py")

    :param module:
    :return:
    """
    # load module
    module = load_module(module)

    # propagate
    publish_functions(module)


#reload_functions("udf_json_support")  # equivalent to: from udf_json_support import *
#reload_functions("./extensions/udf_json_support.py")  # equivalent to: from udf_json_support import *

def register_library(library):
    """
    load external library reference to the udf modules that must be imported in excel
    so that next import adds these libraries

    :param library: module, module name or path to module file
    :return:
    """

    # get module from user
    module = load_module(library, reload=False)  # FIXME reload=True not ready

    # register module
    module_name = module.__name__
    filename = os.path.normcase(module.__file__.lower())

    try:  # getmtime fails for zip imports and frozen modules
        mtime = os.path.getmtime(filename)
    except OSError:
        mtime = None

    udf_modules[module_name] = {
        'filename': filename,
        'filetime': mtime,
        'module': module
    }


def delete_library(library):
    """
    delete external library reference to the udf modules that must be imported in excel
    so that next import deletes these libraries

    :param library: module, module name or path to module file
    :return:
    """

    # get module from user
    module_name = get_module_name(library)
    del udf_modules[module_name]


# --------------------------- EXCEL HELPERS ---------------------------


def create_validation(cell, formula1):
    """
    function used to create drop down lists in excel

    :param cell: range cell to create validation
    :param formula1: string of address or a vector e.g. "A1:A5", "data 1; data 2"
    :return:
    """
    validation = cell.api.Validation
    validation.delete()  # delete any previous drop down list
    # add a drop down list
    validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=formula1)
    validation.IgnoreBlank = -1
    validation.InCellDropdown = -1
    validation.InputTitle = ""
    validation.ErrorTitle = ""
    validation.InputMessage = ""
    validation.ErrorMessage = ""
    validation.ShowInput = -1
    validation.ShowError = -1

# --------------------------- MACROS ---------------------------

@xw.sub
def calculate_all():
    """
    macro to calculate all. Equivalent to :

    Private Sub Workbook_Open()
    Application.CalculateFull
    End Sub
    """
    if not calculate_all.calculating:
        calculate_all.calculating = True
        try:
            wb = xw.Book.caller()
            wb.app.api.CalculateFull()
            return "recalculation done..."
        finally:
            calculate_all.calculating = False
    else:
        return "another recalculation ongoing..."


# global flag to indicate if recalculating
calculate_all.calculating = False
