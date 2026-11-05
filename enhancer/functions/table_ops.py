"""
Table and DataFrame transformation, filtering, grouping, and selection functions for Excel Enhancer.
"""

import re
import warnings
import numpy as np
import pandas as pd
import xlwings as xw
from enhancer.core import (
    convert_to_string,
    debugger,
    general_configuration,
    register,
    strip_accents,
)

custom_group_by = {}  # group for xl_group_by


def build_group_df(selected, operation, c1, c2, c3):
    """
    build a dataframe representation of a group with the selected column and operation

    :param selected: Grouped data with selected column. Must be equivalent to: groupby(c1)[c2]
    :param operation: Applied operation to selected column in group. Must be equivalent to: selected.operation()
    :param c1: column to group
    :param c2: selected column in group to show
    :param c3: name of column resulting from operation
    :return: dataframe
    """
    group_cols = [c1] if isinstance(c1, str) else list(c1)
    sel_cols = [c2] if isinstance(c2, str) else list(c2)
    col_names = [c3] if isinstance(c3, str) else list(c3)

    rows = []
    op_values = operation.values.tolist()
    for (group_key, group_data), op_val in zip(selected, op_values):
        k = list(group_key) if isinstance(group_key, tuple) else [group_key]
        if not isinstance(op_val, list):
            op_val = [op_val]
        if isinstance(group_data, pd.Series):
            for val in group_data:
                rows.append(k + [val] + op_val)
        else:
            for _, r in group_data.iterrows():
                rows.append(k + r.tolist() + op_val)
    cols = group_cols + sel_cols + col_names
    return pd.DataFrame(rows, columns=cols)


@register(custom_group_by, "resume")
def operation_resume(group, sep=", "):
    return sep.join(group)


def _normalize_group_inputs(df, groups, selection, mode):
    """Validate and normalize grouping inputs."""
    if not groups:
        raise ValueError("parameter 'groups' not defined")
    if not selection:
        raise ValueError("parameter 'selection' not defined")

    mode_int = int(mode)
    if mode_int not in [1, 2]:
        raise ValueError(f"mode '{mode}' not recognized")

    norm_groups = [groups] if not isinstance(groups, list) else groups
    norm_sel = list(df.columns) if selection is None else ([selection] if not isinstance(selection, list) else selection)
    return norm_groups, norm_sel, mode_int


def _normalize_group_params(selection, params):
    """Normalize params to match selections."""
    if params is None:
        return [[]] * len(selection)
    if not isinstance(params, list):
        return [[params]] * len(selection)
    if len(params) == len(selection):
        return [p if isinstance(p, (list, tuple)) else [p] for p in params]
    if len(selection) == 1:
        return [params]
    raise ValueError(
        f"params must match number of selections. "
        f"Got {len(selection)} selections and {len(params)} params with {params}"
    )


def _normalize_group_operations(selection, operation, params):
    """Normalize operations and matching parameters for group by."""
    if operation is not None and not isinstance(operation, list):
        operations = [operation]
    else:
        operations = operation

    if operations is not None and len(operations) < len(selection):
        if len(operations) == 1:
            operations *= len(selection)
        else:
            raise ValueError(
                f"operations must match number of selections. "
                f"Got {len(selection)} selections and {len(operations)} operations with {operations}"
            )
    return operations, _normalize_group_params(selection, params)


def _execute_group_operations(grouped, selection, operations, parameters, col_operations):
    """Apply configured built-in or custom aggregation functions."""
    def operation_func(x):
        ops = {}
        for new_col_name, col, operation, param in zip(col_operations, selection, operations, parameters):
            selected = x[col]
            param_list = [] if param is None else (param if isinstance(param, (list, tuple)) else [param])
            try:
                custom_op = custom_group_by[operation]
                ops[new_col_name] = custom_op(selected, *param_list)
            except KeyError:
                try:
                    custom_op_selected = getattr(selected, operation)
                except AttributeError:
                    raise KeyError(
                        f"operation '{operation}' not found. use: mean, sum, size, count, std, var, sem, "
                        "describe, first, last, nth, min, max. Additional custom operations "
                        f"are: {', '.join(custom_group_by)}"
                    )
                ops[new_col_name] = custom_op_selected(*param_list)
        return pd.Series(ops, index=col_operations)

    return grouped.apply(operation_func)


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", pd.DataFrame, index=False, doc="selected table or cell to expand")
@xw.arg("groups", doc="column name to group")
@xw.arg("selection", doc="column name to apply operation")
@xw.arg("operation", doc="operation to do: mean, sum, count, etc.")
@xw.arg("mode", doc="1 - to return normal; 2 - to return only result")
@xw.arg("params", doc="custom parameters to pass to operation function")
@xw.ret(expand="table", index=False)
def xl_group_by(df, groups, selection=None, operation="sum", mode="1", params=None):
    """groups table values applying the desired operations"""
    # DONE and TESTED
    groups, selection, mode_int = _normalize_group_inputs(df, groups, selection, mode)
    operations, parameters = _normalize_group_operations(selection, operation, params)
    grouped = df.groupby(groups)
    selected = grouped[selection]

    if operations is None:
        only_groups = selected.head(len(df))
        return only_groups, [only_groups]

    col_operations = [f"{op_str} {sel}".upper() for op_str, sel in zip(operations, selection)]
    op = _execute_group_operations(grouped, selection, operations, parameters, col_operations)

    db = []
    if debugger.active:
        try:
            db = [build_group_df(selected, op, groups, selection, col_operations)]
        except Exception:
            db = [op]

    if mode_int == 1:
        group_keys = [list(k) if isinstance(k, tuple) else [k] for k in [i[0] for i in grouped[groups]]]
        arrays = [k + j for k, j in zip(group_keys, op.values.tolist())]
        names = groups + col_operations
        return pd.DataFrame(arrays, columns=names), db
    elif mode_int == 2:
        return op, db
    raise ValueError(f"mode '{mode}' not recognized. use 1 - to return normal; 2 - to return only result")


def _parse_filter_modes(mode):
    """Parse and validate filtering modes for inclusion and exclusion."""
    modes = ["exact", "startswith", "endswith", "contains", "match"]
    mode_list = [mode] if not isinstance(mode, list) else mode
    if len(mode_list) > 2:
        raise ValueError(f"mode cannot be larger than 2 for inclusion and exclusion modes. got {mode}")

    parsed = []
    for mo in mode_list:
        try:
            idx = modes.index(mo.strip().lower()) if isinstance(mo, str) else int(mo)
            modes[idx]
            parsed.append(idx)
        except (IndexError, ValueError):
            raise ValueError(f"mode '{mo}' not supported. use {[(i, j) for i, j in enumerate(modes)]}")

    if len(parsed) == 1:
        parsed *= 2
    return modes[parsed[0]], modes[parsed[1]]


def _normalize_filter_strings(df, column, include, exclude, ignore_case, ignore_accents):
    """Normalize selected column, include terms, and exclude terms."""
    inc = [str(i) for i in ([include] if not isinstance(include, list) else include)] if include is not None else []
    exc = [str(i) for i in ([exclude] if not isinstance(exclude, list) else exclude)] if exclude is not None else []
    selected = df[column].apply(convert_to_string)

    if ignore_accents:
        selected = selected.apply(strip_accents)
        inc = [strip_accents(x) for x in inc]
        exc = [strip_accents(x) for x in exc]

    if ignore_case:
        selected = selected.apply(lambda x: x.lower())
        inc = [x.lower() for x in inc]
        exc = [x.lower() for x in exc]

    return selected, inc, exc


def _apply_filter_transform(selected, terms, mode_str, na, flags, regex):
    """Apply string matching operations for filter terms."""
    ops = ["startswith", "endswith"] if mode_str == "exact" else [mode_str]
    operated = None
    for op in ops:
        for t in terms:
            extra = {} if op not in ("match", "contains") else {"regex": regex, "flags": flags}
            res = getattr(selected.str, op)(t, na=na, **extra)
            operated = res if operated is None else (operated & res)
    return operated


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", pd.DataFrame, index=False, doc="selected table or cell to expand")
@xw.arg("column", doc="columns to find values to filter containing rows")
@xw.arg("include", doc="values that has to be found")
@xw.arg("exclude", doc="values that must be ignored")
@xw.arg("mode", doc="filter mode: 0-exact, 1-startswith, 2-endswith, 3-contains, 4-pattern")
@xw.arg("ignore_case", doc="True to ignore Case")
@xw.arg("ignore_accents", doc="True to ignore accents in words")
@xw.arg("na", doc="replacement for Nan values e.g. #N/A")
@xw.arg("regex", doc="True to assume values as regular expression else literal string")
@xw.ret(expand="table", index=False)
def xl_filter_by(df, column, include, exclude=None, mode=0, ignore_case=None, ignore_accents=None, na=False, regex=True):
    """filter table values in columns"""
    flags = re.RegexFlag.IGNORECASE if ignore_case else 0
    inc_mode, exc_mode = _parse_filter_modes(mode)
    selected, inc_terms, exc_terms = _normalize_filter_strings(
        df, column, include, exclude, ignore_case, ignore_accents
    )

    filtered = None
    if inc_terms:
        filtered = _apply_filter_transform(selected, inc_terms, inc_mode, na, flags, regex)

    if exc_terms:
        ignore = _apply_filter_transform(selected, exc_terms, exc_mode, na, flags, regex)
        filtered = ~ignore if filtered is None else (filtered & ~ignore)

    if filtered is None:
        filtered = np.repeat(True, len(df.index))

    db = [pd.concat([selected, filtered], axis=1)] if debugger.active else []
    return df[filtered], db


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", pd.DataFrame, index=False, doc="selected table or cell to expand")
@xw.arg("column", doc="column name to sort")
@xw.arg("ascending", doc="True to sort ascending else descending")
@xw.arg("na_position", doc='"last", "first" to place null data accordingly')
@xw.ret(expand="table", index=False)
def xl_sort_by(df, column=None, ascending=True, na_position="first"):
    """sort table by columns and return it"""
    # DONE and TESTED
    col = list(df.columns) if column is None else column
    return df.sort_values(by=col, ascending=ascending, na_position=na_position), [df]


def _normalize_select_ops(expr, op):
    """Validate and align selection operations."""
    valid_ops = ("column", "like", "regex")
    expr_list = [expr] if not isinstance(expr, list) else expr
    if not isinstance(op, (list, tuple)):
        op_list = [op] * len(expr_list)
    elif len(op) != len(expr_list):
        raise ValueError(
            f"operation of length {len(expr_list)} must match "
            f"expression list of length {len(op)}. got {op}"
        )
    else:
        op_list = list(op)

    for i, o in enumerate(op_list):
        cleaned = o.lower().strip()
        if cleaned not in valid_ops:
            raise ValueError(f"op '{cleaned}' not recognized. use {valid_ops}")
        op_list[i] = cleaned

    if "column" in op_list and len(op_list) != op_list.count("column"):
        raise ValueError(f'"column" operation cannot be combined with others. got {op_list}')
    return expr_list, op_list


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", pd.DataFrame, index=False, doc="selected table or cell to expand")
@xw.arg("expr", doc="according to operation filter by expression")
@xw.arg("op", doc='apply operation "column", "like", "regex"')
@xw.ret(expand="table", index=False)
def xl_select_by(df, expr, op="column"):
    """apply a selection to table columns or indexes"""
    # DONE and TESTED
    expr_list, op_list = _normalize_select_ops(expr, op)
    show_root_df = [df]

    dfs = []
    for ex, o in zip(expr_list, op_list):
        if o == "column":
            return df.filter(items=expr_list), show_root_df
        elif o == "regex":
            dfs.append(df.filter(regex=ex, axis=1))
        elif o == "like":
            dfs.append(df.filter(like=ex, axis=0))

    if len(dfs) == 1:
        return dfs[0], show_root_df
    return pd.concat(dfs), show_root_df


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", pd.DataFrame, index=False, doc="selected table or cell to expand")
@xw.arg("in_map", doc="value or index in columns of the table to change")
@xw.arg("out_map", doc="value to rename to corresponding to in_map")
@xw.arg("errors", doc='"ignore" or "raise" in case where names cannot be renamed. Default "raise"')
@xw.ret(expand="table", index=False)
def xl_rename(df, in_map, out_map, errors="raise"):
    """change table names"""
    # DONE and TESTED
    in_list = [in_map] if not isinstance(in_map, list) else in_map
    out_list = [out_map] if not isinstance(out_map, list) else out_map
    mapping = dict(zip([(df.columns[int(i)] if isinstance(i, float) else i) for i in in_list], out_list))

    try:
        renamed_df = df.rename(columns=mapping, errors=errors)
    except TypeError:
        warnings.warn("using previous versions of 'rename' which does not support 'errors' option")
        renamed_df = df.rename(columns=mapping)

    name_mapping = [pd.DataFrame([df.columns, renamed_df.columns])] if debugger.active else []
    return renamed_df, name_mapping
