"""
Excel UI Watcher implementation for observing selections and rendering debug steps.
"""

import sys
import time
import traceback
import threading as th
import numpy as np
import pandas as pd
import xlwings as xw

from enhancer.core.constants import Special
from enhancer.core.debugger import debugger
from enhancer.core.excel_macros import create_validation


class Watcher:
    """
    Watcher to implement and show DebuggerWatcher
    """

    def __init__(self):
        self.caller = None
        self.watchers = {}
        self.watcher_names = ["address", "plain", "mimic", "print", "steps"]
        self.thread = None

    @xw.func
    @xw.arg('name', doc='Watcher name')
    def set_watcher(self, name="plain"):
        self.caller = wb = xw.Book.caller()
        addrs = wb.selection.get_address(True, False, external=True)
        name = name.lower()
        if name in self.watcher_names:
            self.watchers[addrs] = name
            return addrs
        return f"watcher_name {name} not recognized. use {self.watcher_names}"

    @xw.func
    def del_watcher(self):
        wb = xw.Book.caller()
        addrs = wb.selection.get_address(True, False, external=True)
        del self.watchers[addrs]

    def watch(self):
        my_thread = th.Thread(target=self._watch)
        my_thread.start()
        self.thread = my_thread

    def _update_mimic(self, address, sel, val_olds):
        """Update mirror/mimic watcher display."""
        selection = xw.Range(sel).options(ndim=2).value
        val_old = val_olds.get(address)
        if val_old is not None:
            xw.Range(address).value = np.zeros_like(val_old, dtype="<U1")
        xw.Range(address).value = selection
        val_olds[address] = selection

    @staticmethod
    def _clean_debug_table(cell_debug_table, steps_old, opts):
        """Clear previous debug table data cleanly."""
        if steps_old is not None:
            if opts.get("expand"):
                cell_debug_table.expand().clear()
            else:
                if isinstance(steps_old, pd.DataFrame):
                    cache = steps_old.copy()
                    cache[:] = ""
                    cache.columns = np.zeros_like(cache.columns, dtype="<U1")
                    cache.index = np.zeros_like(cache.index, dtype="<U1")
                    cell_debug_table.value = cache
                else:
                    cell_debug_table.value = np.zeros_like(steps_old, dtype="<U1")
        elif opts.get("expand"):
            cell_debug_table.expand().clear()

    def _update_steps(self, address, opts, wb, steps_olds):
        """Update stepped execution debugger table and validation controls in Excel."""
        formula = wb.selection.sheet[wb.selection.address.split(":")[0]].formula
        steps_old, last_formula_shown, all_menu_options = steps_olds.get(address, (None, None, {}))

        cell_formula = xw.Range(address)
        col, row = cell_formula.column - 1, cell_formula.row - 1
        cell_validation = cell_formula.sheet[row, col + 1]
        cell_steps = cell_formula.sheet[row, col + 2]
        update_gui = False

        try:
            data_cache, refs = debugger.debug_space[formula]
            menu_options = all_menu_options.get(formula, {})
            formula_cache = formula

            if formula != last_formula_shown or not menu_options.get("is_set"):
                available_funcs = list(data_cache.keys())
                choices_range_form = Special.row_sep.join(available_funcs)
                create_validation(cell_validation, choices_range_form)

                selected_func = menu_options.get("selected_func", available_funcs[0])
                if selected_func is not None:
                    cell_validation.value = selected_func

                ret, steps = data_cache[selected_func]
                if steps:
                    create_validation(cell_steps, Special.row_sep.join([str(i) for i in range(len(steps))]))
                    selected_step = menu_options.get("selected_step", "0")
                    cell_steps.value = selected_step

                menu_options["is_set"] = True
                update_gui = True
                old_menu_options = all_menu_options.get(last_formula_shown, {})
                old_menu_options["is_set"] = False
        except KeyError:
            menu_options = all_menu_options.get(last_formula_shown, {})
            if not menu_options.get("is_set"):
                return
            data_cache, refs = debugger.debug_space[last_formula_shown]
            formula_cache = last_formula_shown

        selected_func = cell_validation.value
        selected_step = cell_steps.value
        selected_func_old = menu_options.get("selected_func")
        selected_step_old = menu_options.get("selected_step")
        if selected_func != selected_func_old or selected_step != selected_step_old:
            update_gui = True

        if not update_gui and steps_old is None:
            return

        ret, steps = data_cache[selected_func]
        debug_table = steps[int(selected_step)] if steps else None

        if steps_old is not debug_table:
            update_gui = True

        if not update_gui:
            return

        cell_debug_table = cell_formula.sheet[row + 1, col]
        cell_debug_formula = cell_formula.sheet[row, col + 3]

        self._clean_debug_table(cell_debug_table, steps_old, opts)

        if debug_table is not None:
            cell_debug_table.value = debug_table

        cell_debug_formula.value = "'" + formula_cache
        menu_options["selected_func"] = selected_func
        menu_options["selected_step"] = selected_step
        all_menu_options[formula_cache] = menu_options
        steps_olds[address] = debug_table, formula_cache, all_menu_options

    def _dispatch_watcher(self, address, watcher_entry, sel, wb, val_olds, steps_olds):
        """Dispatch single watcher update based on its registered mode."""
        if isinstance(watcher_entry, tuple):
            watcher_name, opts = watcher_entry
        else:
            watcher_name, opts = watcher_entry, {}

        if watcher_name == "print":
            print(f"{sel} = {xw.Range(sel).raw_value}")
        elif watcher_name == "plain":
            xw.Range(address).value = f"{sel} = {xw.Range(sel).value}"
        elif watcher_name == "address":
            xw.Range(address).value = f"{sel}"
        elif watcher_name == "mimic":
            self._update_mimic(address, sel, val_olds)
        elif watcher_name == "steps":
            self._update_steps(address, opts, wb, steps_olds)

    def _watch(self):
        """
        helper function to watch constantly for user selections where data could be being debugged
        """
        old_sel = None
        val_olds, steps_olds = {}, {}
        while True:
            time.sleep(0.5)
            if not debugger.active:
                continue
            try:
                try:
                    wb = xw.books.active
                except AttributeError:
                    print("no more active books. exiting...")
                    break

                if wb is not None:
                    sel = wb.selection.get_address(True, False, external=True)
                    if sel != old_sel:
                        old_sel = sel
                        for address, watcher_entry in list(self.watchers.items()):
                            self._dispatch_watcher(address, watcher_entry, sel, wb, val_olds, steps_olds)
            except Exception:
                print("Exception in watcher:")
                print("-" * 60)
                traceback.print_exc(file=sys.stdout)
                print("-" * 60)


watcher = Watcher()


@xw.func
@xw.arg('name', doc='Watcher name')
@xw.arg('options', doc='options to pass to the specific watcher')
@xw.arg('xl_app', vba='Application')
def xf_set_watcher(name="address", options=None, xl_app=None):
    caller = xl_app.caller
    xw_caller = xw.Range(impl=xw.xlplatform.Range(xl=caller))
    range_selection = xw_caller.get_address(True, False, external=True)
    equation_address = f"'[{xw_caller.sheet.book.name}]{xw_caller.sheet.name}'!{xw_caller.address}"
    if range_selection.replace("$", "").replace("'", "") != equation_address.replace("$", "").replace("'", ""):
        print(
            f"IN xl_set_watcher got bad equation address {range_selection.replace('$', '')} (gotten) != "
            f"{equation_address.replace('$', '')} (constructed)"
        )

    name = name.lower().strip()
    if name in watcher.watcher_names:
        _options = {}
        if options is not None:
            if isinstance(options, str):
                for opt in options.split(","):
                    _options[opt.strip()] = True
            else:
                if len(options) % 2:
                    raise ValueError(f"options must be series to form key pairs. got {options} and '{options[-1]}' has no pair")
                _options.update(zip(options[::2], options[1::2]))
        watcher.watchers[range_selection] = name, _options
        return range_selection
    raise ValueError(f"watcher name '{name}' not recognized. use {watcher.watcher_names}")


@xw.func
@xw.arg('xl_app', vba='Application')
def xf_del_watcher(xl_app=None):
    caller = xl_app.caller
    xw_caller = xw.Range(impl=xw.xlplatform.Range(xl=caller))
    addrs = xw_caller.get_address(True, False, external=True)
    del watcher.watchers[addrs]

