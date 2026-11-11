"""
Vector and matrix manipulation functions for Excel Enhancer.
"""

import pandas as pd
import xlwings as xw
from enhancer.core import Special, convert_to_string, debugger, general_configuration, strip_accents

_VALID_INDEX_OPS = ("index", "0-index", "1-index")


def _normalize_index_ops(index, op):
    """
    Validate and align index expressions with operations.
    """
    if not isinstance(index, list):
        index = [index]

    if not isinstance(op, (list, tuple)):
        normalized_ops = [op] * len(index)
    elif len(op) != len(index):
        raise ValueError(
            f"operation of length {len(index)} must match "
            f"expression list of length {len(op)}. got {op}"
        )
    else:
        normalized_ops = list(op)

    for i, o in enumerate(normalized_ops):
        cleaned = o.lower().strip()
        if cleaned not in _VALID_INDEX_OPS:
            raise ValueError(f"op '{cleaned}' not recognized. use {_VALID_INDEX_OPS}")
        normalized_ops[i] = cleaned

    return index, normalized_ops


def _resolve_indexed_row(row_values, ex, o):
    """
    Resolve a single row lookup by index and operation.
    """
    try:
        if Special.nan == ex:
            return Special.nan, False
        ex = int(ex)
    except ValueError:
        return Special.err_value, False
    except TypeError:
        return Special.err_null, False

    if o in ("index", "1-index"):
        row = row_values[ex - 1]
    elif o == "0-index":
        row = row_values[ex]
    else:
        row = row_values[ex]
    return row, True


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.ret(expand="table", header=False, index=False)
def xv_concat_h(*tables):
    """concatenate tables horizontally into one"""
    normalized = []
    for i in tables:
        if isinstance(i, list):
            table = pd.DataFrame(i)
        else:
            table = pd.DataFrame([i])
        normalized.append(table)
    return pd.concat(normalized, axis=1), normalized


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.ret(expand="table", header=False, index=False)
def xv_concat_v(*tables):
    """concatenate tables vertically into one"""
    normalized = []
    for i in tables:
        if isinstance(i, list):
            table = pd.DataFrame(i)
        else:
            table = pd.DataFrame([i])
        normalized.append(table)
    return pd.concat(normalized, axis=0), normalized


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("rng", xw.Range, doc="selected range")
@xw.arg("index", doc="according to operation filter by expression")
@xw.arg("op", doc='apply operation "index", "0-index", "1-index"')
@xw.ret(expand="table", index=False)
def xv_index(rng, index, op="index"):
    """apply a selection to table from index"""
    indices, ops = _normalize_index_ops(index, op)
    rows = []
    row_values = rng.value
    expand_len = 0

    for ex, o in zip(indices, ops):
        row, is_valid = _resolve_indexed_row(row_values, ex, o)
        rows.append(row)
        if is_valid and isinstance(row, list) and len(row) > expand_len:
            expand_len = len(row)

    rows_expanded = [i if isinstance(i, list) else [i] * expand_len for i in rows]
    return rows_expanded, [rows_expanded]


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", pd.DataFrame, header=False, index=False, doc="selected data to normalize")
@xw.ret(expand="table", header=False, index=False)
def xv_normalize_accents(df):
    """normalizes removing accents from data"""
    for c in df.columns:
        df[c] = df[c].apply(strip_accents)
    return df, [df]


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", pd.DataFrame, header=False, index=False, doc="selected data to normalize")
@xw.ret(expand="table", header=False, index=False)
def xv_normalize_to_string(df):
    """normalizes data converting everything to string. Numbers like 1 are not
    converted as '1' whereas 1.1 would be '1.1'. Empty or #NA values are converted to ''"""
    for c in df.columns:
        df[c] = df[c].apply(convert_to_string)
    return df, [df]

