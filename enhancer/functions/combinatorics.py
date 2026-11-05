"""
Combinatorics and permutations functions for Excel Enhancer.
"""

import itertools
import numpy as np
import xlwings as xw
from enhancer.core import debugger, general_configuration


# TODO xl_combinations lets the user specify a table where columns are the desired combinations, blank cells are ignored
@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("table", ndim=2, doc="table to apply combinations (randomly) in rows")
@xw.arg("columns", doc='specific columns to only apply combinations. Use "*" for all.')
@xw.arg("has_headers", doc="True to treat first row as headers else use in combinations")
@xw.ret(expand="table")
def xl_combinations(table, columns=None, has_headers=None):
    """
    Generate table combinations.
    """
    raise NotImplementedError
    # order does not matter
    # random rows mth1: df.reindex(np.random.permutation(df.index))
    # random rows mth2: df.sample(df.shape[0])
    # random any column:
    # selected = df[column]
    # df[column] = selected[np.random.permutation(len(selected))].values


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("columns", ndim=2, doc="columns of data to permute (orderly)")
@xw.arg("has_headers", doc="True to treat first row as headers else as data")
@xw.arg("ignore_na", doc="True to treat empty data has a permutation candidate")
@xw.ret(expand="table")
def xf_permutations(columns, has_headers=None, ignore_na=None):
    """
    Generate Cartesian permutations of input columns.
    order matters https://www.mathsisfun.com/combinatorics/combinations-permutations.html
    columns is of the form [row_1, row_2, ..., row_N]
    convert to _columns of the form [column_1, column_2, ..., column_N] and get headers
    """
    _columns = [list(i) for i in zip(*columns)]
    headers = [col.pop(0) for col in _columns] if has_headers else None

    # clean up empty rows
    new_columns = [[row for row in col if row] for col in _columns] if ignore_na else _columns
    permuted = list(itertools.product(*new_columns))

    if headers:
        permuted.insert(0, headers)

    return permuted, [columns, new_columns]

