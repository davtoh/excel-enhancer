"""
Excel Enhancer core subsystem: accessors, converters, macros, debugging, and test generation.
"""

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

from enhancer.core.constants import Special
from enhancer.core.text_utils import (
    general_configuration,
    custom_group_ops,
    find_attribute_in_object_tree,
    find_similar_attribute_in_object_tree,
    register,
    strip_accents,
    col_name,
    convert_to_string,
    helper_correct_to_str,
)
from enhancer.core.converters import (
    is_range_infinite,
    range_bounding_box,
    fit_table,
    ReadValueFromRangeStageInfiniteAdjusted,
    ConverterAdjustInfiniteRange,
    EnsureDataIsWritable,
    ValueAccessorFitted,
    MyConverter,
    pandas_dataframe_converter,
)
from enhancer.core.logger import (
    script_path,
    OperationDebugger,
    file_debugger,
    exception_to_stream,
)
from enhancer.core.test_generator import (
    serializers,
    TestSerializer,
    LastSerializer,
    DataFrameSerializer,
    inheritors,
    UnitTestCreator,
    test_creator,
)
from enhancer.core.excel_macros import (
    create_validation,
    calculate_all,
)
from enhancer.core.module_loader import (
    publish_functions,
    get_module_name,
    load_module,
    reload_functions,
    register_library,
    delete_library,
)
from enhancer.core.debugger import (
    DebuggerWatcher,
    debugger,
    debug,
)
from enhancer.core.watcher import (
    Watcher,
    watcher,
    xf_set_watcher,
    xf_del_watcher,
)

__all__ = [
    "Special",
    "general_configuration",
    "custom_group_ops",
    "find_attribute_in_object_tree",
    "find_similar_attribute_in_object_tree",
    "register",
    "strip_accents",
    "col_name",
    "convert_to_string",
    "helper_correct_to_str",
    "is_range_infinite",
    "range_bounding_box",
    "fit_table",
    "ReadValueFromRangeStageInfiniteAdjusted",
    "ConverterAdjustInfiniteRange",
    "EnsureDataIsWritable",
    "ValueAccessorFitted",
    "MyConverter",
    "pandas_dataframe_converter",
    "script_path",
    "OperationDebugger",
    "file_debugger",
    "exception_to_stream",
    "serializers",
    "TestSerializer",
    "LastSerializer",
    "DataFrameSerializer",
    "inheritors",
    "UnitTestCreator",
    "test_creator",
    "create_validation",
    "calculate_all",
    "publish_functions",
    "get_module_name",
    "load_module",
    "reload_functions",
    "register_library",
    "delete_library",
    "DebuggerWatcher",
    "debugger",
    "debug",
    "Watcher",
    "watcher",
    "xf_set_watcher",
    "xf_del_watcher",
]
