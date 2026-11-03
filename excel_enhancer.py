"""
excel_enhancer

Backward-compatible facade for Excel Enhancer functions and server runtime.

This module re-exports all classes, functions, and utilities from
the decomposed `enhancer.functions` and `enhancer.server` packages to preserve
100% backward compatibility for existing VBA workbooks, test suites, and client code.

Convention:
function starting with the prefix are of the group:
    xl_<name> : specializes in processing tables as DataFrames. always returns a DataFrame with header and no index.
            This always must process Dataframes and return Dataframes with headers and no indexes. No list, no ranges, etc.
    xv_<name> : specializes in processing vectors and matrices which are ranges or headless Dataframes.
            This may return a matrix or a single value. May accept range, lists, arrays, and headless dataframes.
    xs_<name> : specializes in processing single values. Inputs and outputs may be strings, floats, None or dates.
    xf_<name> : special functions with varied functionalities intended for special cases. Often for single value processing.
            This group of funcitons can receive any datatype which cannot fit in the above rules.

LINKS:
recalculate books on start https://github.com/xlwings/xlwings/issues/847#issuecomment-379690303
    Private Sub Workbook_Open()
    Application.CalculateFull
    End Sub

dynamic arrays https://www.xlwings.org/blog/dynamic-arrays-in-excel

FIXME
https://github.com/xlwings/xlwings/issues/1010

NOTES:
    * "array" cannot be used as argument to xw.arg as it is reserved
    * if error = -2146826273 use win32api.FormatMessage(error) to describe the error
    * volatile=True allows to update formulas if any value is changed
        even though selected range is not watched by excel
    # use @xw.ret(cache=False) to prevent function to use cache when evaluated
"""

__version__ = "0.2.0"

from difflib import SequenceMatcher
import itertools
import re
import warnings

from core import *

from enhancer.functions import (
    build_group_df,
    custom_find_similar,
    custom_group_by,
    mean_validation,
    operation_resume,
    rmse_validation,
    similar_std_instance,
    xf_collect_unittest,
    xf_delete_library,
    xf_delete_unittest,
    xf_evaluate,
    xf_fit_table,
    xf_get_caller_address,
    xf_get_selection_range,
    xf_load_library,
    xf_make_array,
    xf_permutations,
    xf_print_unittest,
    xf_recalculate_udfs,
    xf_show_formula,
    xf_test_data,
    xl_combinations,
    xl_concat,
    xl_expand_down,
    xl_expand_right,
    xl_expand_table,
    xl_filter_by,
    xl_find_similar,
    xl_format_by,
    xl_group_by,
    xl_rename,
    xl_select_by,
    xl_sort_by,
    xlv_apply_expr,
    xlv_apply_str,
    xlv_map_by,
    xv_concat_h,
    xv_concat_v,
    xv_find_similar,
    xv_index,
    xv_normalize_accents,
    xv_normalize_to_string,
)
from enhancer.functions import __all__ as _functions_all
from enhancer.server import load_udf_modules, serve

_legacy_exports = ["re", "warnings", "SequenceMatcher", "itertools", "load_udf_modules", "serve"]
__all__ = ["__version__"] + list(_functions_all) + _legacy_exports

if __name__ == "__main__":
    # Expects the Excel file next to this source file, adjust accordingly.
    # xw.Book('excel_enhancer.xlsm').set_mock_caller()

    serve()
    print("excel_enhancer server started...")
else:
    print("excel_enhancer udf module loaded...")
