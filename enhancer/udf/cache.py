"""
Cache management and formula tracking for Excel UDF evaluations.

Maintains equation result caches, reference counts across cells, and
iteration counters to avoid premature recalculations.
"""

import re
import sys
from typing import Any, Callable

import pandas as pd
from xlwings import Range

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

# save equations results according to inputs. This maps cache_key to result using get_cache_key
# TODO use a memoized cache to be shared bewteen sessions so that work is not lost
cache: dict[tuple, Any] = {}

# count the references of a cache_key to the actual cell that is using it
cache_count_references: dict[tuple, int] = {}

# it is the mapping of the cell that uses a cache to the actual cache_key allowing equal cache
# to be shared between cells that use the same opperation
cache_to_delete: dict[str, tuple[str, set[tuple]]] = {}

# cache used to update the last_func_count that determines if a function evaluation is the last in the formula
# i.e. func_name = is_last_function and not last_func_count;  hole formula ended evaluation
func_in_formula_count: dict[tuple[str, str], int] = {}

# re used to detect firt function in a formula
first_formula_func = re.compile(r"(.*?)\(")


def get_first_function(formula: str) -> str | None:
    """Detect and extract the outermost function name from an Excel formula."""
    res = first_formula_func.match(formula)
    if res is None:
        return None
    return res.group(0).replace("=", "").replace("(", "").strip()


def get_cache_key(func: Callable, args: list[Any] | tuple[Any, ...], caller: Any) -> tuple[str, str]:
    """only use this if function is called from cells, not VBA"""
    xw_caller = Range(impl=xlplatform.Range(xl=caller))
    # FIXED bug in which pandas dataframes were converted to test for the key but only some values appeared in the
    # cache key so when values from source were updated key didn't change getting outdated cached data
    # DONE improved key for any cell so that cache can be for multiple cells
    sanatized_args = []
    for arg in args:
        if isinstance(arg, pd.DataFrame):
            sanatized_args.append(arg.to_string())
        elif isinstance(arg, Range):
            sanatized_args.append(arg.value)
        else:
            sanatized_args.append(arg)
    return (func.__name__, str(sanatized_args))
    #return (func.__name__, str(sanatized_args),
    #        str(xw_caller.sheet.book.app.pid) + xw_caller.sheet.book.name + xw_caller.sheet.name + xw_caller.address.split(':')[0])


def _get_str(data: Any) -> str:
    """Extract string value recursively from nested string or list."""
    if isinstance(data, str):
        return data
    return _get_str(data[0])


def clean_stale_cell_cache(eq_start: str, this_equation: str, cache_key: tuple) -> bool:
    """
    Check if equation changed in cell and evict stale unreferenced cached entries.
    Returns True if a new cache entry should be registered.
    """
    try:
        stored_eq, stored_keys = cache_to_delete[eq_start]

        if stored_eq != this_equation:
            for stored_cache in stored_keys:
                count_ref = cache_count_references.get(stored_cache, None)
                if count_ref is not None:
                    count_ref -= 1

                if count_ref is not None and count_ref <= 0:
                    if stored_cache == cache_key:
                        continue  # do not delete cache to current reference
                    print(f"deleting cache of eq '{stored_eq}' in cell {eq_start}")
                    for my_cache in (cache, cache_count_references):
                        my_cache.pop(stored_cache, None)
                else:
                    cache_count_references[stored_cache] = count_ref
            return True

        if cache_key not in stored_keys:
            count = cache_count_references.get(cache_key, 0)
            cache_count_references[cache_key] = count + 1
            stored_keys.add(cache_key)
        return False
    except KeyError:
        return True


def register_cell_cache(eq_start: str, this_equation: str, cache_key: tuple) -> None:
    """Register or increment cache references for a cell equation."""
    count = cache_count_references.get(cache_key, 0)
    cache_count_references[cache_key] = count + 1
    store_keys = {cache_key}
    cache_to_delete[eq_start] = (this_equation, store_keys)

