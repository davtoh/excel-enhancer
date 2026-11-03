"""
Array and table expansion manipulation functions for Excel Enhancer.
"""

import pandas as pd
import xlwings as xw
from enhancer.core import debugger, fit_table, general_configuration


@debugger.get_steps()
@xw.func
@xw.arg("df", pd.DataFrame, header=False, index=False, expand="down", doc="selected table or cell to expand")
@xw.ret(expand="table", header=False, index=False)
def xl_expand_down(df):
    """
    expand selection downward
    """
    # show root data in debugger
    show_root_df = df
    return df, [show_root_df]


@debugger.get_steps()
@xw.func
@xw.arg("df", pd.DataFrame, header=False, index=False, expand="right", doc="selected table or cell to expand")
@xw.ret(expand="table", header=False, index=False)
def xl_expand_right(df):
    """
    expand selection to the right
    """
    # show root data in debugger
    show_root_df = df
    return df, [show_root_df]


@debugger.get_steps()
@xw.func
@xw.arg("df", pd.DataFrame, header=False, index=False, expand="table", doc="selected table or cell to expand")
@xw.ret(expand="table", header=False, index=False)
def xl_expand_table(df):
    """
    expand selection downward and to the right
    """
    # show root data in debugger
    show_root_df = df
    return df, [show_root_df]


@xw.func
@xw.arg("x", xw.Range)
def xf_fit_table(x):
    """
    fit infinite downward selection to a table that has all the values.
    This helps with performance issues as infinite ranges take a long time to process
    even if the are with no values.
    """
    fitted_range = fit_table(x)
    return fitted_range.raw_value


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("df1", pd.DataFrame, index=False, doc="selected table 1")
@xw.arg("df2", pd.DataFrame, index=False, doc="selected table 2")
@xw.arg("axis", doc="axis to apply concatenation")
@xw.ret(expand="table", index=False)
def xl_concat(df1, df2, axis=1):
    """concatenate tables into one"""
    return pd.concat([df1, df2], axis=axis), [df1, df2]

