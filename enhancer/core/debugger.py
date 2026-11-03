"""
DebuggerWatcher implementation for decorating and tracking UDF evaluation steps.
"""

import time
from collections import OrderedDict
from contextlib import contextmanager
from functools import wraps
from pywintypes import com_error

import xlwings as xw

from enhancer.core.logger import file_debugger, exception_to_stream
from enhancer.core.test_generator import test_creator


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

    @staticmethod
    def _resolve_caller(xl_app):
        """Resolve caller Range from xl_app, handling threading delays."""
        tries = 0
        while True:
            try:
                if isinstance(xl_app, xw.Range):
                    return xl_app
                caller = xl_app.caller
                return xw.Range(impl=xw.xlplatform.Range(xl=caller))
            except AttributeError:
                tries += 1
                time.sleep(0.1)
                if tries > 5:
                    raise

    @staticmethod
    def _get_formula_cell(xw_caller):
        """Extract primary formula cell address from caller range."""
        range_selection = xw_caller.get_address(True, False, external=True)
        equation_address = f"'[{xw_caller.sheet.book.name}]{xw_caller.sheet.name}'!{xw_caller.address}"
        norm_range = range_selection.replace("$", "").replace("'", "")
        norm_equation = equation_address.replace("$", "").replace("'", "")
        if norm_range != norm_equation:
            print(
                f"got bad equation address {range_selection.replace('$', '')} (gotten) != "
                f"{equation_address.replace('$', '')} (constructed)"
            )
        return range_selection.split(":")[0]

    def _record_debug_cache(self, formula, formula_cell, func_name, ret, steps):
        """Store or update debug steps for the evaluating formula."""
        if formula in self.debug_space:
            formula_cache = self.debug_space[formula]
            formula_cache[0][func_name] = (ret, steps)

            if formula_cell in self._cell_formulas:
                prev_formula = self._cell_formulas[formula_cell]
                if prev_formula != formula and prev_formula in self.debug_space:
                    prev_cache = self.debug_space[prev_formula]
                    prev_refs = prev_cache[1]
                    prev_refs.discard(formula_cell)
                    if not prev_refs:
                        self.debug_space.pop(prev_formula, None)
            else:
                self._cell_formulas[formula_cell] = formula

            formula_cache[1].add(formula_cell)
        else:
            self._cell_formulas[formula_cell] = formula
            formula_refs = {formula_cell}
            data_cache = OrderedDict()
            data_cache[func_name] = (ret, steps)
            self.debug_space[formula] = [data_cache, formula_refs]

    @staticmethod
    def _inject_app_argument(f, wrapper):
        """Inject Application parameter into wrapper __xlfunc__ metadata."""
        try:
            func_info = f.__xlfunc__
            wrapper.__xlfunc__ = func_info
            args_info = func_info['args']
            applications_tests = [i["vba"] == "Application" for i in args_info]
            if any(applications_tests):
                app_index = applications_tests.index(True)
                return True, app_index

            varargs_i = [i for i, v in enumerate(args_info) if v["vararg"] is True]
            if varargs_i:
                assert len(varargs_i) == 1
                app_index = varargs_i[0]
                args_info[app_index]["pos"] += 1
                args_info[app_index]["doc"] += f'Positional argument {args_info[app_index]["pos"]}'
            else:
                app_index = len(args_info)

            app_name = 'xl_app'
            names = [i["name"] for i in args_info]
            while app_name in names:
                app_name += "1"

            arg_app = {
                'name': app_name, 'pos': app_index, 'vba': "Application",
                'doc': f'Positional argument {app_index}', 'vararg': False,
                'options': {}, 'optional': None
            }
            args_info.insert(app_index, arg_app)
            func_info["argmap"][app_name] = arg_app

            this_copy = func_info.copy()
            xw.arg(app_name, vba='Application')(wrapper)

            for key, val in wrapper.__xlfunc__.items():
                old_val = this_copy[key]
                if old_val != val:
                    wrapper.__xlfunc__[key] = old_val
            return False, app_index
        except AttributeError:
            return False, -1

    def get_steps(self):
        """this decorator should be the first"""
        def outerwrap(f):
            @wraps(f)
            def wrapper(*args, **kwargs):
                func_info = getattr(wrapper, '__xlfunc__', {})
                if self.testing:
                    return f(*args, **kwargs)

                if app_present:
                    xl_app = args[app_index]
                else:
                    args = list(args)
                    xl_app = args.pop(app_index)

                xw_caller = self._resolve_caller(xl_app)
                formula_cell = self._get_formula_cell(xw_caller)

                output = f(*args, **kwargs)
                if self.collect_tests:
                    test_creator.collect_test(f, output, args, kwargs)

                try:
                    ret, steps = output
                except ValueError:
                    ret, steps = output, []

                try:
                    formula = xw_caller.sheet[formula_cell].formula
                    func_name = func_info.get("name", f.__name__)
                    if func_name not in formula:
                        print(f"evaluating '{func_name}' but not in formula '{formula}'")
                    file_debugger.debug(f"evaluating '{formula}'")
                except com_error as e:
                    file_debugger.warning(f"could not get formula for debugging: {func_info}")
                    exception_to_stream(e, stream_handler=file_debugger.warning)
                    return ret

                self._record_debug_cache(formula, formula_cell, func_name, ret, steps)
                return ret

            app_present, app_index = self._inject_app_argument(f, wrapper)
            return wrapper
        return outerwrap

    def debug(self, types=False):
        def outerwrap(f):
            @wraps(f)
            def wrapper(*args, **kwargs):
                if self.active:
                    print(f'calling function "{f.__name__}" with arguments')
                    if types:
                        print('args: (' + ", ".join([repr(i) + str(type(i)) for i in args]) + ')')
                        print('kwargs: {' + ", ".join([repr(k) + ": " + repr(i) + str(type(i)) for k, i in kwargs.items()]) + '}')
                    else:
                        print(f'args: {args}')
                        print(f'kwargs: {kwargs}')
                data = f(*args, **kwargs)
                if self.active:
                    if types:
                        print(f'got data: {data}{type(data)}')
                    else:
                        print(f'got data: {data}')
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
        """
        try:
            yield
        finally:
            pass


debugger = DebuggerWatcher()
debug = debugger  # alias for backwards compatibility and test serialization

