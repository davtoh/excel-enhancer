"""
UDF execution dispatcher and formula lifecycle coordinator.

Converts arguments from Excel types to Python, dispatches function calls
either synchronously or via background threads, handles result caching,
and coordinates dynamic array delayed resizing in Excel.
"""

import sys
from typing import Any, Callable

from win32com.client import Dispatch
from xlwings import conversion, Range

if sys.platform == "win32":
    try:
        from xlwings import _xlwindows as xlplatform
    except ImportError:
        try:
            from xlwings import xlplatform
        except ImportError:
            xlplatform = None
else:
    xlplatform = None

try:
    from xlwings.server import OutputParameter
except Exception:
    OutputParameter = None

from enhancer.udf.cache import (
    cache,
    func_in_formula_count,
    get_cache_key,
    get_first_function,
    clean_stale_cell_cache,
    register_cell_cache,
)
from enhancer.udf.async_worker import AsyncThread, DelayedResizeDynamicArrayFormula
from enhancer.udf.module_registry import get_udf_module


def _convert_input_args(raw_args: tuple[Any, ...], args_info: list[dict[str, Any]],
                        func: Callable, caller: Any) -> list[Any]:
    """Convert raw COM / xlwings arguments to Python objects based on parameter metadata."""
    args = list(raw_args)
    for i, arg in enumerate(args):
        arg_info = args_info[min(i, len(args_info) - 1)]
        if type(arg) is int and arg == -2147352572:  # missing
            args[i] = arg_info.get('optional', None)
        elif xlplatform is not None and xlplatform.is_range_instance(arg):
            if arg_info.get('output', False) and OutputParameter is not None:
                args[i] = OutputParameter(Range(impl=xlplatform.Range(xl=arg)), arg_info['options'], func, caller)
            else:
                args[i] = conversion.read(Range(impl=xlplatform.Range(xl=arg)), None, arg_info['options'])
        else:
            args[i] = conversion.read(None, arg, arg_info['options'])
    return args


def _resolve_caller_equation(xw_caller: Any) -> tuple[str, str]:
    """Extract cell coordinate and formula string from caller range."""
    eq_range = f"'[{xw_caller.sheet.book.name}]{xw_caller.sheet.name}'!{xw_caller.address}"
    eq_start = eq_range.split(":")[0]
    # FIX BUG # _get_str(xw_caller.formula)
    alternative_eq = xw_caller.sheet[eq_start].formula.replace(",", ";")
    return eq_start, alternative_eq


def _track_iteration_count(func_name: str, this_equation: str,
                           eq_key: tuple[str, str]) -> tuple[bool, int]:
    """Track formula execution iterations for nested expressions."""
    last_func_count = this_equation.count(func_name)
    is_last_function = func_name == get_first_function(this_equation)

    if is_last_function:
        if last_func_count > 1:
            if eq_key in func_in_formula_count:
                func_in_formula_count[eq_key] -= 1
                last_func_count = func_in_formula_count[eq_key]
            else:
                func_in_formula_count[eq_key] = last_func_count
            print(f"executing iteration {last_func_count} in function '{func_name}' result of {eq_key}")
        if last_func_count <= 1:
            func_in_formula_count.pop(eq_key, None)

    return is_last_function, last_func_count


def _execute_async(xw_caller: Any, func: Callable, args: list[Any],
                   cache_key: tuple, is_dynamic_array: bool) -> list[list[str]]:
    """Spawn background thread for async UDF calculation and return waiting marker."""
    thread = AsyncThread(
        xw_caller.sheet.book.app.pid,
        xw_caller.sheet.book.name,
        xw_caller.sheet.name,
        xw_caller.address,
        func,
        args,
        cache_key,
        is_dynamic_array,
    )
    thread.start()
    return [["#N/A waiting..." * xw_caller.columns.count] * xw_caller.rows.count]


def _execute_sync_or_async(func: Callable, args: list[Any], func_info: dict[str, Any],
                           ret_info: dict[str, Any], cache_key: tuple,
                           caller: Any, is_dynamic_array: bool) -> tuple[Any, bool]:
    """Execute UDF either via background thread, cached lookup, or direct call."""
    if func_info.get('async_mode') == 'threading':
        cached_value = cache.get(cache_key)
        if cached_value is not None:
            if not is_dynamic_array:
                del cache[cache_key]
            return cached_value, False
        xw_caller = Range(impl=xlplatform.Range(xl=caller))
        return _execute_async(xw_caller, func, args, cache_key, is_dynamic_array), True

    if is_dynamic_array:
        use_cache = ret_info['options'].get('cache', True)
        cached_value = cache.get(cache_key) if use_cache else None
        if cached_value is not None:
            return cached_value, False
        ret = func(*args)
        cache[cache_key] = ret
        return ret, False

    return func(*args), False


def _schedule_dynamic_array_resize(caller: Any, xl_result: Any, is_last_function: bool,
                                   last_func_count: int, this_equation: str,
                                   func_name: str, eq_key: tuple[str, str]) -> None:
    """Schedule idle task for dynamic array formula resizing if result dimensions changed."""
    current_size = (caller.Rows.Count, caller.Columns.Count)
    result_size = (1, 1)
    if isinstance(xl_result, list):
        result_height = len(xl_result)
        result_width = len(xl_result[0]) if result_height else 0
        result_size = (max(1, result_height), max(1, result_width))

    if current_size == result_size:
        return

    if is_last_function and last_func_count <= 1:
        try:
            from xlwings.server import add_idle_task
        except Exception:
            add_idle_task = None
        print(f"executed with resize at iteration {last_func_count} in function '{func_name}' result of {eq_key}")
        if add_idle_task is not None:
            needs_clear = current_size[0] > result_size[0] or current_size[1] > result_size[1]
            if len(this_equation) > 255:  # SOLUTION 2
                add_idle_task(DelayedResizeDynamicArrayFormula(
                    Range(impl=xlplatform.Range(xl=caller)).resize(1, 1), caller, True))
            else:
                add_idle_task(DelayedResizeDynamicArrayFormula(
                    Range(impl=xlplatform.Range(xl=caller)).resize(*result_size), caller, needs_clear))
    else:
        print(f"needs clearing! but ignored. iteration {last_func_count} in function '{func_name}' result of {eq_key}")


def _prepare_caller_context(this_workbook: Any, caller: Any, func: Callable,
                            converted_args: list[Any]) -> tuple[str, str, tuple, tuple[str, str]]:
    """Configure platform caller workbook and compute caller coordinates and cache key."""
    if this_workbook and xlplatform is not None:
        xlplatform.BOOK_CALLER = Dispatch(this_workbook)
    xw_caller = Range(impl=xlplatform.Range(xl=caller))
    eq_start, this_equation = _resolve_caller_equation(xw_caller)
    cache_key = get_cache_key(func, converted_args, caller)
    return eq_start, this_equation, cache_key, (eq_start, this_equation)


def call_udf(module_name: str, func_name: str, args: tuple[Any, ...],
             this_workbook: Any = None, caller: Any = None) -> Any:
    """
    Main dispatch entry point for Excel User Defined Functions.

    Resolves target module/function, converts arguments, executes function
    (sync or async), manages caching, and schedules dynamic array resizing.
    """
    module = get_udf_module(module_name, this_workbook)
    func = getattr(module, func_name)
    func_info = func.__xlfunc__
    ret_info = func_info['ret']
    is_dynamic_array = ret_info['options'].get('expand')

    if func_info.get('writing') and func_info.get('writing') == caller.Address:
        return func_info['rval']

    converted_args = _convert_input_args(args, func_info['args'], func, caller)
    eq_start, this_equation, cache_key, eq_key = _prepare_caller_context(
        this_workbook, caller, func, converted_args)

    is_last_func, last_count = _track_iteration_count(func_name, this_equation, eq_key)
    if clean_stale_cell_cache(eq_start, this_equation, cache_key):
        register_cell_cache(eq_start, this_equation, cache_key)

    ret, is_waiting = _execute_sync_or_async(
        func, converted_args, func_info, ret_info, cache_key, caller, is_dynamic_array)
    if is_waiting:
        return ret

    xl_result = conversion.write(ret, None, ret_info['options'])
    if is_dynamic_array:
        _schedule_dynamic_array_resize(
            caller, xl_result, is_last_func, last_count, this_equation, func_name, eq_key)

    return xl_result
