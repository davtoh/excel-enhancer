"""
Backward-compatible facade for Excel Enhancer core subsystem.

This module re-exports all classes, functions, and utilities from
the decomposed `enhancer.core` package to preserve 100% backward
compatibility for existing VBA workbooks, test suites, and client code.
"""

__version__ = "0.2.0"

import os
import sys
import time
import numpy as np
import pandas as pd
import xlwings as xw

from enhancer.core import (
    Special,
    general_configuration,
    custom_group_ops,
    find_attribute_in_object_tree,
    find_similar_attribute_in_object_tree,
    register,
    strip_accents,
    col_name,
    convert_to_string,
    helper_correct_to_str,
    is_range_infinite,
    range_bounding_box,
    fit_table,
    ReadValueFromRangeStageInfiniteAdjusted,
    ConverterAdjustInfiniteRange,
    EnsureDataIsWritable,
    ValueAccessorFitted,
    MyConverter,
    pandas_dataframe_converter,
    script_path,
    OperationDebugger,
    file_debugger,
    exception_to_stream,
    serializers,
    TestSerializer,
    LastSerializer,
    DataFrameSerializer,
    inheritors,
    UnitTestCreator,
    test_creator,
    create_validation,
    calculate_all,
    publish_functions,
    get_module_name,
    load_module,
    reload_functions,
    register_library,
    delete_library,
    DebuggerWatcher,
    debugger,
    debug,
    Watcher,
    watcher,
    xf_set_watcher,
    xf_del_watcher,
)
from enhancer.core import __all__ as _core_all

_legacy_exports = ["xw", "np", "pd", "time", "os", "sys"]
__all__ = ["__version__"] + list(_core_all) + _legacy_exports
