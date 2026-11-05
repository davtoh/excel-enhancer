"""
Expression evaluation, mapping, formatting, and string transformation functions for Excel Enhancer.
"""

import pandas as pd
import xlwings as xw
from enhancer.core import (
    Special,
    convert_to_string,
    debugger,
    general_configuration,
    pandas_dataframe_converter,
)


def _build_mapper(dic_mapping, keep_original):
    """
    Build dictionary lookup function for xlv_map_by.
    """
    def mapping(val):
        try:
            return dic_mapping[val]
        except KeyError:
            if keep_original:
                # keeps original value if not found on mapping
                return val
            # returns NaN when values are not in mapping
            return Special.empty  # FIXME not real #N/A
    return mapping


def _build_formatter(template):
    """
    Build string template formatter function for xl_format_by.
    """
    if not isinstance(template, str):
        raise ValueError("to_format must be string")

    def mapping(value):
        try:
            if value is None:
                value = ""
            return template.format(value)
        except (TypeError, ValueError):
            return Special.err_value
    return mapping


def _make_exec_handler(expr, args):
    """
    Construct safe execution closure for xlv_apply_expr.
    """
    def custom_op(x, irow, icol, column):
        # https://stackoverflow.com/a/31593712/5288758
        frame = {"x": x, "args": args, "irow": irow, "icol": icol, "column": column}
        exec(expr, frame, frame)
        return frame.get("x")
    return custom_op


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", ndim=2, doc="selected table or cell to expand")
@xw.arg("in_map", ndim=2, doc="column of input of map")
@xw.arg("out_map", ndim=2, doc="column output of map")
@xw.arg("has_headers", doc="True to treat first row as headers else as data")
@xw.arg("keep_original", doc="True to keep original data if not in mapping else #N/A")
@xw.arg("na_action", doc='action to do with Nan values. #N/A or "ignore"')
@xw.ret(expand="table", index=False)
def xlv_map_by(df, in_map, out_map=None, has_headers=None, keep_original=None, na_action=None):
    """apply a mapping to a table to change values to others"""
    options = {"index": False, "header": bool(has_headers)}
    df = pandas_dataframe_converter(df, options)
    in_map = pandas_dataframe_converter(in_map, options)
    out_map = pandas_dataframe_converter(out_map, options)

    if len(in_map) != len(out_map):
        raise ValueError(
            f"in_map of len {len(in_map)} and out_map of len "
            f"{len(out_map)} must have the same length"
        )
    dic_mapping = dict(zip(in_map[in_map.columns[0]], out_map[out_map.columns[0]]))
    mapping = _build_mapper(dic_mapping, keep_original)

    ret = df.copy()
    for column in df.columns:
        ret[column] = df[column].map(mapping, na_action=na_action)

    show_concat = [pd.concat([df, ret], axis=1)]
    if not has_headers:
        return ret.values.tolist(), show_concat
    return ret, show_concat


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", ndim=2, doc="selected table or cell to expand")
@xw.arg("template", doc="template to format values in df")
@xw.arg("has_headers", doc="True to treat first row as headers else as data")
@xw.arg("na_action", doc='action to do with Nan values. #N/A or "ignore"')
@xw.ret(expand="table", index=False)
def xl_format_by(df, template, has_headers=None, na_action=None):
    """format cell values """
    options = {"index": False, "header": bool(has_headers)}
    df = pandas_dataframe_converter(df, options)
    mapping = _build_formatter(template)

    ret = df.copy()
    for column in df.columns:
        ret[column] = df[column].map(mapping, na_action=na_action)

    show_concat = [pd.concat([df, ret], axis=1)]
    if not has_headers:
        return ret.values.tolist(), show_concat
    return ret, show_concat


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", ndim=2, doc="selected table or cell to expand")
@xw.arg("expresion", doc="expresion to evaluate")
@xw.arg("args", doc="additional arguments")
@xw.arg("has_headers", doc="True to treat first row as headers else as data")
@xw.ret(expand="table", index=False)
def xlv_apply_expr(df, expresion, args=None, has_headers=False):
    """apply a python expresion to table columns"""
    options = {"index": False, "header": bool(has_headers)}
    df = pandas_dataframe_converter(df, options)

    if args is None:
        args = ()
    elif not isinstance(args, (list, tuple)):
        args = [args]

    if not isinstance(expresion, (list, tuple)):
        expresion = [expresion]

    operations = [_make_exec_handler(expr, args) for expr in expresion]
    if len(operations) == 1:
        operations *= len(df.columns)
    elif len(operations) != len(df.columns):
        raise ValueError(
            f"expresion ({len(operations)}) must be of the same number of columns ({len(df.columns)})"
        )

    for icol, (column, custom_op) in enumerate(zip(df.columns, operations)):
        for irow, val in enumerate(df[column]):
            df.iloc[irow, icol] = custom_op(val, irow, icol, column)

    if not has_headers:
        return df.values.tolist(), []
    return df, []


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df", ndim=2, doc="selected table or cell to expand")
@xw.arg("operations", doc="string operation to evaluate")
@xw.arg("args", doc="additional arguments")
@xw.arg("has_headers", doc="True to treat first row as headers else as data")
@xw.ret(expand="table", index=False)
def xlv_apply_str(df, operations, args=None, has_headers=False):
    """apply string opperation to table columns"""
    options = {"index": False, "header": bool(has_headers)}
    df = pandas_dataframe_converter(df, options)

    if not isinstance(operations, (list, tuple)):
        operations = [operations]

    if len(operations) == 1:
        operations *= len(df.columns)
    elif len(operations) != len(df.columns):
        raise ValueError(
            f"operations ({len(operations)}) must be of the same number of columns ({len(df.columns)})"
        )

    if args is None:
        args = ()
    elif not isinstance(args, (list, tuple)):
        args = [args]

    for col, op in zip(df.columns, operations):
        try:
            df[col] = getattr(df[col].str, op)(*args)
        except AttributeError:
            df[col] = getattr(df[col].apply(convert_to_string).str, op)(*args)

    if not has_headers:
        return df.values.tolist(), []
    return df, []

