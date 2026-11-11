"""
Unit test serialization and test creation helpers.
"""

import os
import sys
import pickle
import pandas as pd

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
        formats["test_name"] = "MyTestFor" + "".join([i.capitalize() for i in test_name.split("_")])

        # add test cases
        lines = []
        functions = set()
        for (i, (cached_key, (func, ret, args, kwargs))) in enumerate(test_cases.items()):
            test_case_name = f"test_{func.__name__}{i}"
            functions.add(("excel_enhancer", func.__name__))
            formatted = self.format_test_case(test_case_name, func, ret, args, kwargs)
            lines.append(formatted)

        i1 = formats["i1"]
        formats["test_cases"] = "".join(["\n".join([f"{i1}{j}" for j in i.split("\n")]) for i in lines])

        # add imports
        main_imports = [
            ("core", UnitTestCreator.__name__),
            ("core", TestSerializer.__name__),
            ("core", "debug")
        ]
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
        formats = self._get_indents()
        i1 = formats["i1"]
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

            formats["arg_serializations"] = f"{i1}# get arguments\n" + "\n".join(lines_arg_serializations)
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

            formats["kwarg_serializations"] = f"{i1}# get key word arguments\n" + "\n".join(lines_kwarg_serializations)
        else:
            formats["kwarg_serializations"] = ""
            names_kwarg = []

        # for ret
        ret_val = self.serialize(ret)
        line = f"{i1}ret_cmp = {ret_val}"
        formats["ret_serializations"] = f"{i1}# get expected result\n" + line

        # function_execution
        str_args = ", ".join(names_arg)
        str_kwargs = (", " + ", ".join(names_kwarg)) if names_kwarg else ""

        line = f"{i1}ret = {func.__name__}({str_args}{str_kwargs})"
        formats["function_execution"] = f"{i1}# evaluate function\n" + line

        # function_test
        line = f"{i1}# compare expected an result\n{i1}self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))"
        formats["function_test"] = line

        formatted = template_test_case.format(**formats)
        return formatted

    @classmethod
    def serialize(cls, data):
        """serializer to insert into formatted code"""
        return TestSerializer.serialize(data)

    @classmethod
    def deserialize(cls, data, type=None):
        """deserializer to insert into formatted code"""
        return TestSerializer.deserialize(data, type)

    @classmethod
    def get_cache_key(cls, func, args):
        """get cache key to know if already used"""
        return func.__name__, str([i.to_string() if isinstance(i, pd.DataFrame) else i for i in args])


test_creator = UnitTestCreator()

