"""
Excel Enhancer UDF subsystem.

Provides decorators, argument conversion, dynamic array delayed resizing,
background execution threads, caching, dispatching, and VBA code generation.
"""

from enhancer.udf.decorators import (
    func_sig,
    get_category,
    get_async_mode,
    check_bool,
    xlfunc,
    xlsub,
    xlret,
    xlarg,
)
from enhancer.udf.cache import (
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
)
from enhancer.udf.async_worker import (
    has_dynamic_array,
    AsyncThread,
    DelayedResizeDynamicArrayFormula,
)
from enhancer.udf.module_registry import (
    udf_modules,
    get_udf_module,
)
from enhancer.udf.dispatcher import (
    call_udf,
)
from enhancer.udf.vba_generator import (
    generate_vba_wrapper,
    import_udfs,
)

__all__ = [
    # Decorators & inspection
    "func_sig",
    "get_category",
    "get_async_mode",
    "check_bool",
    "xlfunc",
    "xlsub",
    "xlret",
    "xlarg",
    # Cache & formula tracking
    "cache",
    "cache_count_references",
    "cache_to_delete",
    "func_in_formula_count",
    "first_formula_func",
    "get_first_function",
    "get_cache_key",
    "_get_str",
    "clean_stale_cell_cache",
    "register_cell_cache",
    # Async worker & dynamic array
    "has_dynamic_array",
    "AsyncThread",
    "DelayedResizeDynamicArrayFormula",
    # Module registry
    "udf_modules",
    "get_udf_module",
    # Dispatcher
    "call_udf",
    # VBA code generation
    "generate_vba_wrapper",
    "import_udfs",
]

