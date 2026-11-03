"""
Backward-compatible facade for Excel Enhancer UDF subsystem.

This module re-exports all decorators, dispatch mechanisms, cache managers,
and VBA generators from the decomposed `enhancer.udf` package to preserve 100%
backward compatibility for existing Excel workbooks, VBA callers, and test suites.
"""

import os
import re
import tempfile
import inspect
from threading import Thread
import pywintypes
import win32api
import pandas as pd
from xlwings import conversion, Range, apps, Book
from xlwings.utils import VBAWriter
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

try:
    from xlwings import PRO
except (ImportError, AttributeError):
    PRO = False

from enhancer.udf import (
    func_sig,
    get_category,
    get_async_mode,
    check_bool,
    xlfunc,
    xlsub,
    xlret,
    xlarg,
    cache,
    cache_count_references,
    cache_to_delete,
    func_in_formula_count,
    first_formula_func,
    get_first_function,
    get_cache_key,
    _get_str,
    clean_stale_cell_cache,
    register_cell_cache,
    has_dynamic_array,
    AsyncThread,
    DelayedResizeDynamicArrayFormula,
    udf_modules,
    get_udf_module,
    call_udf,
    generate_vba_wrapper,
    import_udfs,
)
from enhancer.udf import __all__ as _udf_all

_legacy_exports = ["os", "re", "tempfile", "inspect", "Thread", "pywintypes",
                   "win32api", "pd", "conversion", "Range", "apps", "Book", "VBAWriter", "PRO"]
__all__ = list(_udf_all) + _legacy_exports
