"""
Dynamic module registration and reloading for Excel UDFs.

Tracks loaded Python modules, handles embedded code for xlwings PRO,
and manages hot reloading when module files change on disk.
"""

import os
import os.path
from importlib import import_module, reload
from types import ModuleType
from typing import Any

from win32com.client import Dispatch
from xlwings import Book
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

if PRO:
    try:
        from xlwings.pro.embedded_code import dump_embedded_code, get_udf_temp_dir
    except Exception:
        PRO = False

udf_modules: dict[str, dict[str, Any]] = {}

# Setup temp dir for embedded code
tempdir = get_udf_temp_dir() if PRO else None


def _reload_if_modified(module_info: dict[str, Any]) -> ModuleType:
    """Reload module if file modification time has changed."""
    module = module_info['module']
    if module_info['filetime'] is not None:
        mtime = os.path.getmtime(module_info['filename'])
        if mtime != module_info['filetime']:
            module = reload(module)
            module_info['filetime'] = mtime
            module_info['module'] = module
    return module


def _load_new_udf_module(module_name: str, xl_workbook: Any) -> ModuleType:
    """Import and record metadata for a newly requested UDF module."""
    if PRO and tempdir is not None and xl_workbook is not None:
        wb = Book(impl=xlplatform.Book(Dispatch(xl_workbook)))
        dump_embedded_code(wb, tempdir.name)

    module = import_module(module_name)
    filename = os.path.normcase(module.__file__.lower()) if module.__file__ else ""

    try:  # getmtime fails for zip imports and frozen modules
        mtime = os.path.getmtime(filename) if filename else None
    except OSError:
        mtime = None

    udf_modules[module_name] = {
        'filename': filename,
        'filetime': mtime,
        'module': module
    }
    return module


def get_udf_module(module_name: str, xl_workbook: Any = None) -> ModuleType:
    """Retrieve or dynamically import/reload a UDF module by name."""
    module_info = udf_modules.get(module_name, None)
    if module_info is not None:
        return _reload_if_modified(module_info)
    return _load_new_udf_module(module_name, xl_workbook)

