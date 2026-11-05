"""
Asynchronous UDF execution workers and dynamic array resizing.

Handles background worker threads for async UDF calculation and delayed
formula array expansion tasks in Excel.
"""

from threading import Thread
from typing import Any

from win32com.client import CDispatch
import pywintypes
from xlwings import apps, Range

from enhancer.udf.cache import cache


def has_dynamic_array(app: Any) -> bool:
    """This check in this form doesn't work on macOS, that's why it's here and not in utils"""
    try:
        app.api.WorksheetFunction.Unique("dummy")
        return True
    except pywintypes.com_error as e:
        if e.hresult == -2147352567:
            return False
        raise e


class AsyncThread(Thread):
    """Background worker thread for asynchronous UDF execution."""

    def __init__(self, pid: int, book_name: str, sheet_name: str, address: str,
                 func: Any, args: list[Any], cache_key: tuple, expand: bool):
        super().__init__()
        self.pid = pid
        self.book = book_name
        self.sheet = sheet_name
        self.address = address
        self.func = func
        self.args = args
        self.cache_key = cache_key
        self.expand = expand

    def run(self) -> None:
        # apps[self.pid].books[self.book].sheets[self.sheet][self.address].api._inner
        new_caller = None
        for i, arg in enumerate(self.args):
            if isinstance(arg, CDispatch):
                if new_caller is None:
                    # FIXME this should simulate the actual object in main thread
                    r"""
                     File "H:\common_partition\MEGAsync\python_projects\usb_projects\usb_python_env\lib\site-packages\xlwings\_xlwindows.py", line 602, in range
                        xl1 = self.xl.Range(arg1)
                    AttributeError: 'NoneType' object has no attribute 'Range'
                    """
                    new_caller = apps[self.pid].books[self.book].sheets[self.sheet][self.address]  # .api._inner
                self.args[i] = new_caller  # replace stale caller with new
        cache[self.cache_key] = self.func(*self.args)

        cell = apps[self.pid].books[self.book].sheets[self.sheet][self.address]
        if self.expand:
            cell.formula_array = cell.formula_array
        else:
            if has_dynamic_array(apps[self.pid]):
                cell.formula2 = cell.formula2
            else:
                cell.formula = cell.formula


def _handle_long_formula_parts(target_api: Any, formula: str) -> None:
    """Split formula exceeding 255 chars into 200-char parts with placeholders."""
    # formula longer than 255 chars, needs to be devided by parts
    # http://excel.dokry.com/matriz-de-combinacin-con-ms-de-255-caracteres.html
    length = 200
    place_holder = "X_X_X"
    place_holder_eq = "Y_Y_Y"
    steps = list(range(0, len(formula), length))
    parts = [formula[i:i + length] + ("" if i == steps[-1] else place_holder) for i in steps]
    # assert "".join(parts) == formula
    eq_position = parts[0].find("=") + 1
    replace_eq = parts[0][0:eq_position]
    part_no_eq = parts[0][eq_position:]

    # add first part without equal sign (=) to prevent interpretation of formula
    target_api.FormulaArray = place_holder_eq + part_no_eq

    # replace at each step placeholder with next part of formula
    for to_replace in parts[1:]:
        target_api.replace(place_holder, to_replace)

    # SOLUTION 1
    # FIXME This triggers calculations on each cell
    # target_api.replace(place_holder_eq, replace_eq)

    # SOLUTION 2
    # replace the equal sign to interpret equation
    target_api.replace(place_holder_eq, "Array equation is too long. it must be 255 characters'" + replace_eq)
    raise ValueError("Array equation is too long. it must be 255 characters")


class DelayedResizeDynamicArrayFormula:
    """Idle task callable for resizing dynamic array formulas in Excel."""

    def __init__(self, target_range: Any, caller: Any, needs_clearing: bool):
        self.target_range = target_range
        self.caller = caller
        self.needs_clearing = needs_clearing

    def __call__(self, *args: Any, **kwargs: Any) -> None:
        formula = self.caller.FormulaArray
        if self.needs_clearing:
            self.caller.ClearContents()
        try:
            self.target_range.api.FormulaArray = formula
        except pywintypes.com_error as e:
            if e.hresult == -2147352567:
                _handle_long_formula_parts(self.target_range.api, formula)
            else:
                raise e

