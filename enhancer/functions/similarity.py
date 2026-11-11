"""
Fuzzy string similarity matching and distance validation functions for Excel Enhancer.
"""

from difflib import SequenceMatcher
import numpy as np
import pandas as pd
import xlwings as xw
from enhancer.core import Special, debugger, general_configuration, register

custom_find_similar = {}


def similar_std_instance(a, b):
    """
    helper function for match_text

    :param a:
    :param b:
    :return:
    """
    s = SequenceMatcher(None, str(a).lower(), str(b).lower())
    ratio = s.ratio()
    op_codes = s.get_opcodes()
    return ratio, op_codes


@register(custom_find_similar, "mean")
def mean_validation(row_rates):
    """
    mean validation function for similarity

    :param row_rates:
    :return: overall_rate
    """
    return float(np.mean(row_rates))


@register(custom_find_similar, "rmse")
def rmse_validation(row_rates):
    """
    root mean square error validation function for similarity

    :param row_rates:
    :return: overall_rate
    """
    predictions = np.array(row_rates)
    targets = np.ones_like(predictions)
    rmse = np.sqrt(((predictions - targets) ** 2).mean())
    return float(1 - rmse)


def _parse_similarity_modes(mode, base_modes):
    """
    Parse mode flags into numerical indices and mode names.
    """
    mode_list = [mode] if not isinstance(mode, list) else mode
    parsed_modes = []
    special_cases = (len(base_modes), "all")

    for mo in mode_list:
        try:
            idx = base_modes.index(mo.strip().lower()) if isinstance(mo, str) else int(mo)
            base_modes[idx]
            parsed_modes.append(idx)
        except (IndexError, ValueError):
            if len(mode_list) == 1:
                if mo in special_cases:
                    return list(range(len(base_modes))), base_modes
                modes_with_all = base_modes + ["all"]
            elif mo in special_cases:
                raise ValueError(f"special case '{mo}' cannot be combined with others. got {mode_list}")
            else:
                modes_with_all = base_modes
            raise ValueError(f"mode '{mo}' not supported. use {[(i, j) for i, j in enumerate(modes_with_all)]}")

    return parsed_modes, base_modes


def _compute_vector_matches(value, vector, modes, parsed_modes, min_rate):
    """
    Compute similarity rates across range cells.
    """
    match = []
    mode_address = modes.index("address")
    mode_row = modes.index("row")
    process_all = any(i in parsed_modes for i in (mode_address, mode_row)) or debugger.active

    for row, row_val in zip(vector, vector.options(ndim=2).value):
        for col, cell_val in zip(row, row_val):
            if cell_val:
                rate, _ = similar_std_instance(cell_val, value)
                test = rate >= min_rate
                if process_all:
                    match.append((cell_val, col.address, col.row, rate, test))
                else:
                    match.append((cell_val, None, None, rate, test))

    mode_rate = modes.index("rate")
    match.sort(key=lambda x: x[mode_rate], reverse=True)
    return match


def _format_vector_results(match, parsed_modes, modes, to_slice):
    """
    Slice and format vector similarity results with Special.nan placeholders.
    """
    mode_test = modes.index("test")
    if to_slice is None:
        latest = None
        for i, val in enumerate(match):
            if not val[mode_test] and latest is None:
                latest = i
                break
        to_slice = slice(0, 1 if latest is None else latest)

    values_re = [[i[j] for j in parsed_modes] for i in match[to_slice]]
    values_ex = []
    nanable_modes = {"value", "address", "row"}

    for i, row in enumerate(values_re):
        to_replace_as_nan = not match[i][mode_test]
        row_expansion = []
        for j, val in enumerate(row):
            actual_mode = modes[parsed_modes[j]]
            if to_replace_as_nan and actual_mode in nanable_modes:
                row_expansion.append(Special.nan)
            else:
                row_expansion.append(val)
        values_ex.append(row_expansion)

    return pd.DataFrame(values_ex, columns=[modes[i] for i in parsed_modes])


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("value", doc="value to find similar in vector")
@xw.arg("vector", xw.Range, doc="vector to find value")
@xw.arg("min_rate", doc="minimum acceptable rate which goes from 0 to 1 being 0 the less likely and 1 the closest match")
@xw.arg("num", doc="number of predictions to output. By default 1")
@xw.arg("mode", doc='mode to return values. 0-"value", 1-"address", 2-"row", 3-"rate", 4-"test", 5-"all"')
@xw.ret(expand="table", header=False, index=False, empty="NA")
def xv_find_similar(value, vector, min_rate=0.8, num=1, mode="value"):
    """
    find similar value in vector with a rating tending to 1 when vector values
    are the most similar to value.
    """
    # DONE and TESTED
    to_slice = None if num is True else (slice(0, int(num)) if num else slice(None, None))
    base_modes = ["value", "address", "row", "rate", "test"]
    parsed_modes, modes = _parse_similarity_modes(mode, base_modes)

    norm_value = "" if not value else value
    match = _compute_vector_matches(norm_value, vector, modes, parsed_modes, min_rate)
    db = [pd.DataFrame(match, columns=[i.upper() for i in modes])] if debugger.active else []

    if not match:
        return Special.nan, db

    df_result = _format_vector_results(match, parsed_modes, modes, to_slice)
    return df_result, db


def _compute_matrix_matches(values, matrix, modes, parsed_modes, min_rate, validation_func):
    """
    Evaluate matrix rows against input values vector using selected validation function.
    """
    match = []
    mode_address = modes.index("address")
    mode_row = modes.index("row")
    process_all = any(i in parsed_modes for i in (mode_address, mode_row)) or debugger.active
    base_address = f"'[{matrix.sheet.book.name}]{matrix.sheet.name}'!"
    values_2d = matrix.options(ndim=2).value
    rows = range(matrix.row, len(values_2d) + matrix.row)
    letter = xw.utils.col_name(matrix.column)

    for row_no, row_val in zip(rows, values_2d):
        row_values = []
        row_rates = []
        for cell_val, value in zip(row_val, values):
            str_value = "" if not value else str(value)
            str_cell_val = "" if not cell_val else str(cell_val)
            rate, _ = similar_std_instance(str_cell_val, str_value)
            row_values.append(cell_val)
            row_rates.append(rate)

        overall_rate = validation_func(row_rates)
        test = not min_rate or overall_rate >= min_rate
        if process_all and test:
            address = base_address + f"${letter}${row_no}"
            match.append((row_values, address, row_no, overall_rate, test, row_rates))
        else:
            match.append((row_values, None, None, overall_rate, test, row_rates))

    mode_rate = modes.index("rate")
    match.sort(key=lambda x: x[mode_rate], reverse=True)
    return match


def _build_expansion_columns(columns, sample):
    """Determine expansion widths and expanded column labels."""
    expansions, columns_ex = [], []
    for col_num, col in enumerate(sample):
        if isinstance(col, list):
            extend = [columns[col_num] + str(i) for i, _ in enumerate(col)]
            expansions.append(len(extend))
            columns_ex.extend(extend)
        else:
            expansions.append(0)
            columns_ex.append(columns[col_num])
    return expansions, columns_ex


def _expand_matrix_row(row, expansions, parsed_modes, modes, to_replace_as_nan, row_idx):
    """Expand a single matrix result row handling Special.nan substitutions."""
    nanable_modes = {"value", "address", "row"}
    row_expansion = []
    for j, (expand, value) in enumerate(zip(expansions, row)):
        actual_mode = modes[parsed_modes[j]]
        if expand:
            if len(value) != expand:
                raise ValueError(f"Expansion '{value}' does not correspond to values {value} in row {row_idx}")
            row_expansion.extend([Special.nan] * expand if to_replace_as_nan and actual_mode in nanable_modes else value)
        else:
            row_expansion.append(Special.nan if to_replace_as_nan and actual_mode in nanable_modes else value)
    return row_expansion


def _expand_matrix_results(match, parsed_modes, modes, to_slice, header):
    """Expand sub-columns and format output matrix with optional header."""
    columns = [modes[i] for i in parsed_modes]
    mode_test = modes.index("test")
    if to_slice is None:
        latest = None
        for i, val in enumerate(match):
            if not val[mode_test] and latest is None:
                latest = i
                break
        to_slice = slice(0, 1 if latest is None else latest)

    values_re = [[i[j] for j in parsed_modes] for i in match[to_slice]]
    expansions, columns_ex = _build_expansion_columns(columns, values_re[0])

    values_ex = [
        _expand_matrix_row(row, expansions, parsed_modes, modes, not match[i][mode_test], i)
        for i, row in enumerate(values_re)
    ]

    if header:
        return [columns_ex] + values_ex
    return values_ex


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.arg("values", doc="selected values to find similar in matrix")
@xw.arg("matrix", xw.Range, doc="value or index in columns of the table to change")
@xw.arg("min_rate", doc="minimum acceptable rate which goes from 0 to 1 being 0 the less likely and 1 the closest match")
@xw.arg("num", doc="number of predictions to output. By default 1. Use TRUE to fit to data, #N/A or FALSE to print all")
@xw.arg("mode", doc='mode to return values. 0-"value", 1-"address", 2-"row", 3-"rate", 4-"test", 5-"cell-rate", 6-"all"')
@xw.arg("header", doc="allow printing headers")
@xw.arg("validation", doc=f"validation function to determine overall rates. use: {', '.join(custom_find_similar)}")
@xw.ret(expand="table", header=False, index=False, empty="NA")
def xl_find_similar(values, matrix, min_rate=0.8, num=1, mode="value", header=False, validation="rmse"):
    """
    find similar values in matrix with a rating tending to 1 when matrix values
    are the most similar to value.
    """
    # DONE and TESTED
    if matrix is None:
        raise TypeError("xl_find_similar's matrix argument needs access to actual ranges in spreadsheet")

    to_slice = None if num is True else (slice(0, int(num)) if num else slice(None, None))
    base_modes = ["value", "address", "row", "rate", "test", "cell-rate"]
    parsed_modes, modes = _parse_similarity_modes(mode, base_modes)
    norm_values = [values] if not isinstance(values, list) else values

    try:
        validation_func = custom_find_similar[validation]
    except KeyError:
        raise ValueError(f"requested validation '{validation}' not supported. use: {', '.join(custom_find_similar)}")

    match = _compute_matrix_matches(norm_values, matrix, modes, parsed_modes, min_rate, validation_func)
    db = [pd.DataFrame(match, columns=[i.upper() for i in modes])] if debugger.active else []

    if not match:
        return Special.nan, db

    formatted = _expand_matrix_results(match, parsed_modes, modes, to_slice, header)
    return formatted, db
