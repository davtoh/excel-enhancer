"""
excel_enhancer

Program to provide or facility User Defined Functions (UDFs) in excel so as to quickly create reports,
mange data for analysis or automate repetitive input data into a final product.

Convention:
function starting with the prefix are of the group:
    xl_<name> : specializes in processing tables as DataFrames. always returns a DataFrame with header and no index.
            This always must process Dataframes and return Dataframes with headers and no indexes. No list, no ranges, etc.
    xv_<name> : specializes in processing vectors and matrices which are ranges or headless Dataframes.
            This may return a matrix or a single value. May accept range, lists, arrays, and headless dataframes.
    xs_<name> : specializes in processing single values. Inputs and outputs may be strings, floats, None or dates.
    xf_<name> : special functions with varied functionalities intended for special cases. Often for single value processing.
            This group of funcitons can receive any datatype which cannot fit in the above rules.

UDF Functions:

Special functions:
Special

LINKS:
recalculate books on start https://github.com/xlwings/xlwings/issues/847#issuecomment-379690303
    Private Sub Workbook_Open()
    Application.CalculateFull
    End Sub

dynamic arrays https://www.xlwings.org/blog/dynamic-arrays-in-excel

FIXME
https://github.com/xlwings/xlwings/issues/1010

NOTES:
    * "array" cannot be used as argument to xw.arg as it is reserved
    * if error = -2146826273 use win32api.FormatMessage(error) to describe the error
    * volatile=True allows to update formulas if any value is changed
        even though selected range is not watched by excel
    # use @xw.ret(cache=False) to prevent function to use cache when evaluated
"""
import re
import warnings
from difflib import SequenceMatcher
import itertools
#from xlwings.udfs import import_udfs
from core import *

# TODO improve documentation and add the STANDARD and FRAMEWORK for all the custom UDF functions
# TODO in "udfs.py" implement the memoized or in memory-persistent dictionary for runs to be consistent between sessions
# TODO create the interface to easily register custom functions in sheet for UDF functions with "@register" but for the user
# TODO help tooltip in excel is not useful, maybe consider using a web interface for all the help
# TODO according to locale or lang add help in functions in the desired language much like in Django
# TODO add to each UDF the "category"
# DONE create a submodule like "core" to import from all the basic functionality of the enhancer
# TODO make an example file for each function name
# TODO make tests for each function while using the example file to check for correctness
# TODO consider creating a brother project of a web version that complements this to create custom
#  extensions on the fly and with more useful help messages for functions
# TODO instead of adding header (option if should print or not header) make a function that deheaders results
# TODO change in core a stage in ValueAccessorFitted to allow to print header in another letter color and cell fill

# --------------------------- ARRAY MANIPULATION FUNCTIONS ---------------------------


@debugger.get_steps()
@xw.func#(volatile=True)
@xw.arg('df', pd.DataFrame, header=False, index=False, expand="down", doc='selected table or cell to expand')
@xw.ret(expand='table', header=False, index=False)  #
def xl_expand_down(df):
    """
    expand selection downward
    """
    # show root data in debugger
    show_root_df = df
    return df, [show_root_df]


@debugger.get_steps()
@xw.func#(volatile=True)
@xw.arg('df', pd.DataFrame, header=False, index=False, expand="right", doc='selected table or cell to expand')
@xw.ret(expand='table', header=False, index=False)  #
def xl_expand_right(df):
    """
    expand selection to the right
    """
    # show root data in debugger
    show_root_df = df
    return df, [show_root_df]


@debugger.get_steps()
@xw.func#(volatile=True)
@xw.arg('df', pd.DataFrame, header=False, index=False, expand="table", doc='selected table or cell to expand')
@xw.ret(expand='table', header=False, index=False)  #
def xl_expand_table(df):
    """
    expand selection downward and to the right
    """
    # show root data in debugger
    show_root_df = df
    return df, [show_root_df]


@xw.func
@xw.arg("x", xw.Range)
#@xw.ret(expand='table')
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
@xw.arg('df1', pd.DataFrame, index=False, doc='selected table 1')
@xw.arg('df2', pd.DataFrame, index=False, doc='selected table 2')
@xw.arg('axis', doc='axis to apply concatenation')
@xw.ret(expand='table', index=False)  #
def xl_concat(df1, df2, axis=1):
    """concatenate tables into one"""
    return pd.concat([df1, df2], axis=axis), [df1, df2]


@debugger.get_steps()
@xw.func(**general_configuration)
@xw.ret(expand='table', header=False, index=False)  #
def xv_concat_h(*tables):
    """concatenate tables into one"""
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
@xw.ret(expand='table', header=False, index=False)  #
def xv_concat_v(*tables):
    """concatenate tables into one"""
    normalized = []
    for i in tables:
        if isinstance(i, list):
            table = pd.DataFrame(i)
        else:
            table = pd.DataFrame([i])
        normalized.append(table)
    return pd.concat(normalized, axis=0), normalized


# --------------------------- PRODUCTION FUNCTIONS ---------------------------


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
    Ad, Bd, Cd = [], [], []
    for i, j, k in [([i] * len(j), j, [k] * len(j)) for (i, j), k in zip(selected, operation.values.tolist())]:
        Ad += i
        Bd += list(j)
        Cd += k
    if not isinstance(c1, (list, tuple)):
        arrays = [Ad, Bd]
        index = pd.MultiIndex.from_arrays(arrays, names=[c1]+c2)
    else:
        arrays = list(zip(*zip(*zip(*Ad), Bd)))
        names = [i[0] for i in zip(*zip(*zip(*[c1]), c2))]
        index = pd.MultiIndex.from_arrays(arrays, names=names)
    return pd.DataFrame({c3: Cd}, index=index)


custom_group_by = {}  # group for xl_group_by


@register(custom_group_by, "resume")
def operation_resume(group, sep=", "):
    return sep.join(group)


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', pd.DataFrame, index=False, doc='selected table or cell to expand')
@xw.arg('groups', doc='column name to group')
@xw.arg('selection', doc='column name to apply operation')
@xw.arg('operation', doc='operation to do. E.g. mean, sum, size, count, std, var, sem, '
                         'describe, first, last, nth, min, max. Additional custom operations '
                         f'are: {", ".join(custom_group_by)}')
@xw.arg('mode', doc='1 - to return normal; 2 - to return only result')
@xw.arg('params', doc='custom parameters to pass to operation function')
@xw.ret(expand='table', index=False)  #
#@xw.arg('xl_app', vba='Application')
def xl_group_by(df, groups, selection=None, operation="sum", mode="1", params=None):  # , xl_app=None
    """groups table values applying the desired operations"""
    """agrupa valores en una tabla aplicando las operaciones deseadas"""
    # DONE and TESTED
    # good base https://stackoverflow.com/a/47103408/5288758

    # get caller cell where operation starts, it can be range if dynamic array with actual size
    #wb = xw.Book.caller()
    #addrs = wb.selection.get_address(True, False, external=True)
    # [["#N/A waiting..." * xw_caller.columns.count] * xw_caller.rows.count]
    # target_range.api.FormulaArray = formula

    if not groups:
        raise ValueError("parameter 'groups' not defined")

    if not selection:
        raise ValueError("parameter 'selection' not defined")

    # convert mode
    mode = int(mode)
    if mode not in [1, 2]:
        raise ValueError(F"mode '{mode}' not recognized")

    # normalize groups
    if not isinstance(groups, list):
        groups = [groups]

    # normalize selection
    if selection is None:
        # select all
        selection = df.columns
    elif not isinstance(selection, list):
        selection = [selection]

    # normalize operations
    if operation is not None and not isinstance(operation, list):
        operations = [operation]
    else:
        operations = operation

    if operations is not None and len(operations) < len(selection):
        if len(operations) == 1:
            operations *= len(selection)  # match number of selections
        else:
            raise ValueError(F"operations must match number of selections. "
                             F"Got {len(selection)} selections and {len(operations)} operations with {operations}")

    # normalize params
    if not isinstance(params, list):
        parameters = [params]*len(selection)  # match number of selections
    else:
        parameters = params
        if len(parameters) < len(selection):
            raise ValueError(F"params must match number of selections. "
                         F"Got {len(selection)} selections and {len(parameters)} operations with {parameters}")

    # group on groups
    grouped = df.groupby(groups)

    selected = grouped[selection]

    if operations is None:
        only_groups = selected.head(len(df))
        # routine without operations
        return only_groups, [only_groups]

    col_operations = [F"{op_str} {sel}".upper() for op_str, sel in zip(operations, selection)]
    # DONE allow operation for more than 1 column
    def operation_func(x):
        ops = {}
        for new_col_name, col, operation, param in zip(col_operations, selection, operations, parameters):
            selected = x[col]
            param = param or []
            try:
                # get custom fucntion
                custom_op = custom_group_by[operation]
                ops[new_col_name] = custom_op(selected, *param)
            except KeyError:
                # get buildin fucntion
                try:
                    custom_op_selected = getattr(selected, operation)
                except AttributeError:
                    raise KeyError(f"operation '{operation}' not found. use: mean, sum, size, count, std, var, sem, "
                                   'describe, first, last, nth, min, max. Additional custom operations '
                                   f'are: {", ".join(custom_group_by)}')
                ops[new_col_name] = custom_op_selected(*param)
        return pd.Series(ops, index=col_operations)

    # get opperations
    op = grouped.apply(operation_func)

    # show group selection with operations
    if debugger.active:
        db = [build_group_df(selected, op, groups, selection, col_operations)]
    else:
        db = []

    if mode == 1:
        if len(groups) > 1:
            arrays = [list(i) + j for i, j in zip([i[0] for i in grouped[groups]], op.values.tolist())]
        else:
            # in case i is to subscriptable
            arrays = [list(i) + j for i, j in zip([[i[0]] for i in grouped[groups]], op.values.tolist())]
        names = groups + col_operations
        return pd.DataFrame(arrays, columns=names), db
    elif mode == 2:
        return op, db
    else:
        raise ValueError(F"mode '{mode}' not recognized. use 1 - to return normal; 2 - to return only result")


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', pd.DataFrame, index=False, doc='selected table or cell to expand')
@xw.arg('column', doc='columns to find values to filter containing rows')
@xw.arg('include', doc='values that has to be found')
@xw.arg('exclude', doc='values that must be ignored')
@xw.arg('mode', doc="filter mode for include and exclude. Both can be "
                    "specified as {'mode for include';'mode for exclude'}:\n"
                    '0 - "exact" to match exact values\n'
                    '1 - "startswith" to match cells starting with value\n'
                    '2 - "endswith" to match cells ending with value\n'
                    '3 - "contains" to match cells containing value\n'
                    '4 - "patter" to match a value as re expression')
@xw.arg('ignore_case', doc="True to ignore Case")
@xw.arg('ignore_accents', doc='True to ignore accents in words')
@xw.arg('na', doc='replacement for Nan values e.g. #N/A')
@xw.arg('regex', doc='True to assume values as regular expression else literal string')
@xw.ret(expand='table', index=False)  #
def xl_filter_by(df, column, include, exclude=None, mode=0, ignore_case=None, ignore_accents=None, na=False, regex=True):
    """filter table values in columns"""
    # DONE and TESTED
    # DONE allow n-columns
    # DONE allow filtering modes
    if include is None:
        include = []
    elif not isinstance(include, list):
        include = [include]
    include = [str(i) for i in include]

    if exclude is None:
        exclude = []
    elif not isinstance(exclude, list):
        exclude = [exclude]
    exclude = [str(i) for i in exclude]

    # get flags
    _flags = 0
    flags = None
    if ignore_case:
        flags = "IGNORECASE"
    if flags is None:
        flags = []
    elif not isinstance(flags, list):
        flags = [flags]

    for f in flags:
        _flags &= getattr(re.RegexFlag, f)

    # check mode
    modes = ["exact", "startswith", "endswith", "contains", "match"]  # values
    if not isinstance(mode, list):
        mode = [mode]
    if len(mode) > 2:
        raise ValueError(F"mode cannot be larger than 2 for inclusion and exclusion modes. got {mode}")

    _modes = []  # indexes
    for mo in mode:  # user choices
        try:
            # test
            if isinstance(mo, str):
                mo = modes.index(mo.strip().lower())
            else:
                mo = int(mo)
                modes[mo]
            _modes.append(mo)
        except (IndexError, ValueError):
            raise ValueError(F"mode '{mo}' not supported. use {[(i,j) for i, j in enumerate(modes)]}")
    if len(_modes) == 1:
        _modes *= 2  # modes for either inclusion and exclusion
    include_mode, exclude_mode = [modes[i] for i in _modes]

    selected = df[column].apply(convert_to_string)  # everything works as str

    # get column series where data will be filtered
    if ignore_accents:
        # normalize data to ignore accents
        selected = selected.apply(strip_accents)
        # normalize looked data so that accents are ignored
        include = [strip_accents(x) for x in include]  # without accents
        # normalize looked data so that accents are ignored
        exclude = [strip_accents(x) for x in exclude]  # without accents

    # get column series where data will be filtered
    if ignore_case:
        selected = selected.apply(lambda x: x.lower())
        include = [x.lower() for x in include]
        exclude = [x.lower()for x in exclude]

    # sanitize include and exclude
    #if include_mode in ("match", "contains"):
    #    include = [re.escape(i) for i in include]
    #if exclude_mode in ("match", "contains"):
    #    exclude = [re.escape(i) for i in exclude]

    def get_operations(ops):
        if ops == "exact":
            return ["startswith", "endswith"]
        return [ops]

    def transform(selected, trans, ops, na, flags, regex=True):
        to_operate = selected.str
        operated = None
        for op in ops:
            for t in trans:
                res = getattr(to_operate, op)(t, na=na, **({} if op not in ("match", "contains") else
                                                           {"regex": regex, "flags": flags}))
                if operated is None:
                    operated = res
                else:
                    operated &= res
        return operated

    # find data
    include_ops = get_operations(include_mode)
    exclude_ops = get_operations(exclude_mode)

    # filtered = selected == include
    # filtered = selected.str.contains('|'.join(include), na=na, regex=True, flags=_flags)
    filtered = None
    if include:
        filtered = transform(selected, include, include_ops, na, _flags, regex)

    if exclude:
        # ignore = selected.str.contains('|'.join(exclude), na=na, regex=True, flags=_flags)
        ignore = transform(selected, exclude, exclude_ops, na, _flags, regex)
        if filtered is None:
            filtered = ~ignore
        else:
            filtered &= ~ignore

    if filtered is None:
        filtered = np.repeat(True, len(df.index))  # all true if not filtered

    # show value selection with bool operation
    if debugger.active:
        show_bool_selection = [pd.concat([selected, filtered], axis=1)]
    else:
        show_bool_selection = []
    return df[filtered], show_bool_selection


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', pd.DataFrame, index=False, doc='selected table or cell to expand')
@xw.arg('column', doc='column name to sort')
@xw.arg('ascending', doc='True to sort ascending else descending')
@xw.arg('na_position', doc='"last", "first" to place null data accordingly')
@xw.ret(expand='table', index=False)  #
def xl_sort_by(df, column=None, ascending=True, na_position='first'):
    """sort table by columns and return it"""
    # DONE and TESTED
    # show root data in debugger
    show_root_df = [df]
    if column is None:
        column = list(df.columns)
    return df.sort_values(by=column, ascending=ascending, na_position=na_position), show_root_df


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', pd.DataFrame, index=False, doc='selected table or cell to expand')
@xw.arg('expr', doc='according to operation filter by expression')
@xw.arg('op', doc='apply operation "column", "like", "regex"')
@xw.ret(expand='table', index=False)  #
def xl_select_by(df, expr, op="column"):
    """apply a selection to table columns or indexes"""
    # DONE and TESTED

    opperations = ("column", "like", "regex")
    def check_opps(op):
        op = op.lower().strip()
        if op not in opperations:
            raise ValueError(F"op '{op}' not recognized. use {opperations}")
        return op

    if not isinstance(expr, list):
        expr = [expr]

    if not isinstance(op, (list, tuple)):
        op = [op] * len(expr)
    elif len(op) != len(expr):
        raise ValueError(F"operation of length {len(expr)} must match "
                         F"expression list of length {len(op)}. got {op}")

    for i, o in enumerate(op):
        op[i] = check_opps(o)

    if "column" in op and len(op) != op.count("column"):
        raise ValueError(F'"column" operation cannot be combined with others. got {op}')

    # show root data in debugger
    show_root_df = [df]

    dfs = []
    for ex, o in zip(expr, op):
        if o == "column":
            return df.filter(items=expr), show_root_df
        elif o == "regex":
            dfs.append(df.filter(regex=ex, axis=1))
        elif o == "like":
            dfs.append(df.filter(like=ex, axis=0))

    if len(dfs) == 1:
        return dfs[0], show_root_df
    return pd.concat(dfs), show_root_df


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('rng', xw.Range, doc='selected range')
@xw.arg('index', doc='according to operation filter by expression')
@xw.arg('op', doc='apply operation "column", "like", "regex"')
@xw.ret(expand='table', index=False)  #
def xv_index(rng, index, op="index"):
    """apply a selection to table from index"""
    # DONE and TESTED

    opperations = ("index", "0-index", "1-index")
    def check_opps(op):
        op = op.lower().strip()
        if op not in opperations:
            raise ValueError(F"op '{op}' not recognized. use {opperations}")
        return op

    if not isinstance(index, list):
        index = [index]

    if not isinstance(op, (list, tuple)):
        op = [op] * len(index)
    elif len(op) != len(index):
        raise ValueError(F"operation of length {len(index)} must match "
                         F"expression list of length {len(op)}. got {op}")

    for i, o in enumerate(op):
        op[i] = check_opps(o)

    rows = []
    row_values = rng.value
    expand_len = 0
    for ex, o in zip(index, op):
        try:
            if Special.nan == ex:
                rows.append(Special.nan)
                continue
            ex = int(ex)
        except ValueError:
            rows.append(Special.err_value)
            continue
        except TypeError:
            rows.append(Special.err_null)
            continue
        if o in ("index", "1-index"):
            row = row_values[ex-1]
            rows.append(row)
        elif o == "0-index":
            row = row_values[ex]
            rows.append(row)
        if len(row) > expand_len:
            expand_len = len(row)

    rows_expanded = [i if isinstance(i, list) else [i]*expand_len for i in rows]

    return rows_expanded, [rows_expanded]


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', pd.DataFrame, header=False, index=False, doc='selected data to normalize')
@xw.ret(expand='table', header=False, index=False)  #
def xv_normalize_accents(df):
    """normalizes removing accents from data"""
    for c in df.columns:
        df[c] = df[c].apply(strip_accents)
    return df, [df]


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', pd.DataFrame, header=False, index=False, doc='selected data to normalize')
@xw.ret(expand='table', header=False, index=False)  #
def xv_normalize_to_string(df):
    """normalizes data converting everything to string. Numbers like 1 are not
    converted as '1' whereas 1.1 would be '1.1'. Empty or #NA values are converted to ''"""
    for c in df.columns:
        df[c] = df[c].apply(convert_to_string)
    return df, [df]


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', ndim=2, doc='selected table or cell to expand')
@xw.arg('in_map', ndim=2, doc='column of input of map')
@xw.arg('out_map', ndim=2, doc='column output of map')
@xw.arg('has_headers', doc='True to treat first row as headers else as data')
@xw.arg('keep_original', doc='True to keep original data if not in mapping else #N/A')
@xw.arg('na_action', doc='action to do with Nan values. #N/A or "ignore"')
@xw.ret(expand='table', index=False)  #
def xlv_map_by(df, in_map, out_map=None, has_headers=None, keep_original=None, na_action=None):
    """apply a mapping to a table to change values to others"""

    # convert data
    if has_headers:
        options = {"index": False, "header": True}
    else:
        options = {"index": False, "header": False}

    df = pandas_dataframe_converter(df, options)
    in_map = pandas_dataframe_converter(in_map, options)
    out_map = pandas_dataframe_converter(out_map, options)
    if len(in_map) != len(out_map):
        raise ValueError(F"in_map of len {len(in_map)} and out_map of len "
                         F"{len(out_map)} must have the same length")
    dic_mapping = dict(zip(in_map[in_map.columns[0]], out_map[out_map.columns[0]]))

    def mapping(val):
        try:
            return dic_mapping[val]
        except KeyError:
            if keep_original:
                # keeps original value if not found on mapping
                return val
            else:
                # returns NaN when values are not in mapping
                return Special.empty  # FIXME not real #N/A

    ret = df.copy()
    for column in df.columns:
        # Input NaN values are replaced by na_action, Output NaN are converted to special empty
        ret[column] = df[column].map(mapping, na_action=na_action)

    # show input with output columns
    show_concat = [pd.concat([df, ret], axis=1)]

    if not has_headers:
        return ret.values.tolist(), show_concat
    return ret, show_concat


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', ndim=2, doc='selected table or cell to expand')
@xw.arg('template', doc='template to format values in df')
@xw.arg('has_headers', doc='True to treat first row as headers else as data')
@xw.arg('na_action', doc='action to do with Nan values. #N/A or "ignore"')
@xw.ret(expand='table', index=False)  #
def xl_format_by(df, template, has_headers=None, na_action=None):
    """format cell values """

    # convert data
    if has_headers:
        options = {"index": False, "header": True}
    else:
        options = {"index": False, "header": False}

    df = pandas_dataframe_converter(df, options)
    # DONE and TESTED
    if isinstance(template, str):
        # CASE in_map is string with the from "{}"
        def mapping(value):
            try:
                if value is None:
                    value = ""
                return template.format(value)
            except (TypeError, ValueError):
                return Special.err_value
    else:
        raise ValueError("to_format must be string")

    ret = df.copy()
    for column in df.columns:
        # Input NaN values are replaced by na_action, Output NaN are converted to special empty
        ret[column] = df[column].map(mapping, na_action=na_action)

    # show input with output columns
    show_concat = [pd.concat([df, ret], axis=1)]

    if not has_headers:
        return ret.values.tolist(), show_concat
    return ret, show_concat


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', ndim=2, doc='selected table or cell to expand')
@xw.arg('expresion', doc='expresion to evaluate')
@xw.arg('args', doc='additional arguments')
@xw.arg('has_headers', doc='True to treat first row as headers else as data')
@xw.ret(expand='table', index=False)  #
def xlv_apply_expr(df, expresion, args=None, has_headers=False):
    """apply a python expresion to table columns"""

    # convert data
    if has_headers:
        options = {"index": False, "header": True}
    else:
        options = {"index": False, "header": False}

    df = pandas_dataframe_converter(df, options)

    if args is None:
        args = ()
    elif not isinstance(args, (list, tuple)):
        args = [args]

    if not isinstance(expresion, (list, tuple)):
        expresion = [expresion]

    operations = []
    for expr in expresion:
        # make custom operation for expresion
        def custom_op(x, irow, icol, column):
            # https://stackoverflow.com/a/31593712/5288758
            frame = {'x': x, 'args': args, "irow": irow, "icol": icol, "column": column}
            exec(expr, frame, frame)
            return frame.get("x")
        operations.append(custom_op)

    if len(operations) == 1:
        operations *= len(df.columns)
    elif len(operations) != len(df.columns):
        raise ValueError(f"expresion ({len(operations)}) must be of the same number of columns ({len(df.columns)})")

    for icol, (column, custom_op) in enumerate(zip(df.columns, operations)):
        for irow, val in enumerate(df[column]):
            df.iloc[irow, icol] = custom_op(val, irow, icol, column)

    if not has_headers:
        return df.values.tolist(), []
    return df, []


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', ndim=2, doc='selected table or cell to expand')
@xw.arg('operations', doc='string operation to evaluate')
@xw.arg('args', doc='additional arguments')
@xw.arg('has_headers', doc='True to treat first row as headers else as data')
@xw.ret(expand='table', index=False)  #
def xlv_apply_str(df, operations, args=None, has_headers=False):
    """apply string opperation to table columns"""

    # convert data
    if has_headers:
        options = {"index": False, "header": True}
    else:
        options = {"index": False, "header": False}

    df = pandas_dataframe_converter(df, options)

    if not isinstance(operations, (list, tuple)):
        operations = [operations]

    if len(operations) == 1:
        operations *= len(df.columns)
    elif len(operations) != len(df.columns):
        raise ValueError(f"operations ({len(operations)}) must be of the same number of columns ({len(df.columns)})")

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


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('df', pd.DataFrame, index=False, doc='selected table or cell to expand')
@xw.arg('in_map', doc='value or index in columns of the table to change')
@xw.arg('out_map', doc='value to rename to corresponding to in_map')
@xw.arg('errors', doc='"ignore" or "raise" in case where names cannot be renamed. Default "raise"')
@xw.ret(expand='table', index=False)  #
def xl_rename(df, in_map, out_map, errors="raise"):
    """change table names"""
    # DONE and TESTED
    if not isinstance(in_map, list):
        in_map = [in_map]
    if not isinstance(out_map, list):
        out_map = [out_map]
    mapping = dict(zip([(df.columns[int(i)] if isinstance(i, float) else i) for i in in_map], out_map))

    # rename column names
    try:
        # using recent pandas
        renamed_df = df.rename(columns=mapping, errors=errors)
    except TypeError:
        warnings.warn("using previous versions of 'rename' which does not support 'errors' option")
        # backing up to previous versions
        renamed_df = df.rename(columns=mapping)


    # show column name mapping for debugging
    if debugger.active:
        name_mapping = [pd.DataFrame([df.columns, renamed_df.columns])]
    else:
        name_mapping = []

    return renamed_df, name_mapping


def similar_std_instance(a, b):
    """
    helper function for match_text

    :param a:
    :param b:
    :return:
    """
    s = SequenceMatcher(None, a.lower(), b.lower())
    ratio = s.ratio()
    op_codes = s.get_opcodes()
    return ratio, op_codes


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('value', doc='value to find similar in vector')
@xw.arg('vector', xw.Range, doc='vector to find value')
@xw.arg('min_rate', doc='minimum acceptable rate which goes from 0 to 1 being 0 the less likely and 1 the closest match')
@xw.arg('num', doc='number of predictions to output. By default 1')
@xw.arg('mode', doc='mode to return values. '
                    '0-"value" to return value of found address\n'
                    '1-"address" for found address\n'
                    '2-"row" for the row number\n'
                    '3-"rate" for matching rates\n'
                    '4-"test"\n'
                    '5-"all" to return all values')
@xw.ret(expand='table', header=False, index=False, empty='NA')  #
def xv_find_similar(value, vector, min_rate=0.8, num=1, mode="value"):
    """
    find similar value in vector with a rating tending to 1 when vector values
    are the most similar to value.
    """
    # DONE and TESTED
    # check num
    # allow num=None to calculate it when #N/A values start
    if num is True:
        to_slice = None
    elif num:
        num = int(num)
        to_slice = slice(0, num)
    else:
        to_slice = slice(None, None)  # no rows equivalent to all ie. 0 or None

    # check mode
    modes = ["value", "address", "row", "rate", "test"]  # values
    if not isinstance(mode, list):
        mode = [mode]

    _modes = []  # indexes
    special_cases = (len(modes), "all")
    for mo in mode:  # user choices
        try:
            # test
            if isinstance(mo, str):
                mo = modes.index(mo.strip().lower())
            else:
                mo = int(mo)
                modes[mo]
            _modes.append(mo)
        except (IndexError, ValueError):
            if len(mode) == 1:
                if mo in special_cases:  # special case to get all
                    _modes = list(range(len(modes)))
                    break
                modes += ["all"]  # append especial case in single case for error
            elif mo in special_cases:
                raise ValueError(F"special case '{mo}' cannot be combined with others. got {mode}")

            raise ValueError(F"mode '{mo}' not supported. use {[(i,j) for i, j in enumerate(modes)]}")

    # normalize value
    if not value:
        value = ""

    # order best match
    match = []
    mode_address = modes.index("address")
    mode_row = modes.index("row")
    process_all = any([i in _modes for i in (mode_address, mode_row)]) or debugger.active
    for row, row_val in zip(vector, vector.options(ndim=2).value):
        for col, cell_val in zip(row, row_val):
            if cell_val:
                rate, d = similar_std_instance(cell_val, value)
                test = rate >= min_rate
                if process_all:
                    # additional processed data which takes resources
                    match.append((cell_val,  col.address, col.row, rate, test))
                else:
                    match.append((cell_val, None, None, rate, test))

    mode_rate = modes.index("rate")
    match.sort(key=lambda x: x[mode_rate], reverse=True)

    # debug data
    if debugger.active:
        # rates, values, addresses = list(zip(*match))
        db = [pd.DataFrame(match, columns=[i.upper() for i in modes])]
    else:
        db = []

    if not match:
        return Special.nan, db

    # get slice from the valid tests
    mode_test = modes.index("test")
    if to_slice is None:
        # calculate to get all valid values
        latest = None
        for i, val in enumerate(match):
            if not val[mode_test] and latest is None:
                latest = i
                break
        if latest is None:
            latest = 1
        to_slice = slice(0, latest)

    # replace None values
    values_re = [[i[j] for j in _modes] for i in match[to_slice]]

    values_ex = []
    nanable_modes = {"value", "address", "row"}
    for i, row in enumerate(values_re):

        # replace the ones that does not pass test with #N/A
        to_replace_as_nan = not match[i][mode_test]

        # replace None values
        if to_replace_as_nan:
            row_expansion = []
            for j, value in enumerate(row):
                if modes[j] in nanable_modes:
                    row_expansion.append(Special.nan)
                else:
                    row_expansion.append(value)
        else:
            row_expansion = row
        values_ex.append(row_expansion)

    return pd.DataFrame(values_ex, columns=[modes[i] for i in _modes]), db


custom_find_similar = {}  # custom group form xl_find_similar function


@register(custom_find_similar, "mean")
def mean_validation(row_rates):
    """
    mean validation function for similarity

    :param row_rates:
    :return: overall_rate
    """
    overall_rate = np.mean(row_rates)
    return overall_rate


@register(custom_find_similar, "rmse")
def rmse_validation(row_rates):
    """
    root mean square error validation function for similarity

    :param row_rates:
    :return: overall_rate
    """
    predictions = np.array(row_rates)
    targest = np.ones_like(predictions)
    rmse = np.sqrt(((predictions - targest) ** 2).mean())
    return 1-rmse


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('values', doc='selected values to find similar in matrix')
@xw.arg('matrix', xw.Range, doc='value or index in columns of the table to change')
@xw.arg('min_rate', doc='minimum acceptable rate which goes from 0 to 1 being 0 the less likely and 1 the closest match')
@xw.arg('num', doc='number of predictions to output. By default 1. Use TRUE to fit to data, #N/A or FALSE to print all')
@xw.arg('mode', doc='mode to return values. '
                    '0-"value" to return value of found address\n'
                    '1-"address" for found address\n2-"row" for the row number\n'
                    '2-"row"\n'
                    '3-"rate" for matching rates\n'
                    '4-"test"\n'
                    '5-"cell-rate"\n'
                    '6-"all" to return all values'
        )
@xw.arg('header', doc='allow printing headers')
@xw.arg('validation', doc='validation function to determine overall rates and if test matching is successful. '
                          f'use: {", ".join(custom_find_similar)}')
@xw.ret(expand='table', header=False, index=False, empty='NA')  #
def xl_find_similar(values, matrix, min_rate=0.8, num=1, mode="value", header=False, validation="rmse"):
    """
    find similar values in matrix with a rating tending to 1 when matrix values
    are the most similar to value.
    """
    # DONE and TESTED
    # DONE allow printing of header values
    if matrix is None:
        raise TypeError("xl_find_similar's matrix argument needs access to actual ranges in spreadsheet")

    # check num
    # allow num=None to calculate it when #N/A values start
    if num is True:
        to_slice = None
    elif num:
        num = int(num)
        to_slice = slice(0, num)
    else:
        to_slice = slice(None, None)  # no rows equivalent to all ie. 0 or None

    # check mode
    modes = ["value", "address", "row", "rate", "test", "cell-rate"]  # values
    if not isinstance(mode, list):
        mode = [mode]

    _modes = []  # indexes
    special_cases = (len(modes), "all")
    for mo in mode:  # user choices
        try:
            # test
            if isinstance(mo, str):
                mo = modes.index(mo.strip().lower())
            else:
                mo = int(mo)
                modes[mo]
            _modes.append(mo)
        except (IndexError, ValueError):
            if len(mode) == 1:
                if mo in special_cases:  # special case to get all
                    _modes = list(range(len(modes)))
                    break
                modes += ["all"]  # append especial case in single case for error
            elif mo in special_cases:
                raise ValueError(F"special case '{mo}' cannot be combined with others. got {mode}")

            raise ValueError(F"mode '{mo}' not supported. use {[(i,j) for i, j in enumerate(modes)]}")

    # convert values to list
    if not isinstance(values, list):
        values = [values]

    # TODO allow to pass custom validation functions
    # choose validation function
    try:
        validation_func = custom_find_similar[validation]
    except KeyError:
        raise ValueError(F"requested validation '{validation}' not supported. use: {', '.join(custom_find_similar)}")

    # order best match
    match = []
    mode_address = modes.index("address")
    mode_row = modes.index("row")
    process_all = any([i in _modes for i in (mode_address, mode_row)]) or debugger.active
    base_address = f"'[{matrix.sheet.book.name}]{matrix.sheet.name}'!"  # INDEX does not support full address
    values_2d = matrix.options(ndim=2).value
    rows = range(matrix.row, len(values_2d) + matrix.row)
    letter = xw.utils.col_name(matrix.column)
    for row_no, row_val in zip(rows, values_2d):
        row_values = []
        row_rates = []

        for cell_val, value in zip(row_val, values):

            # normalize values
            # floats which are numbers in excel will have precision points
            # while if a number is a string in excel it won't. It is good
            # practice since types are different i.e. str != float
            # e.g. 1.0 -> '1.0' vs 1 -> '1' will give rate 0.5 instead of 1
            str_value, str_cell_val = "", ""
            if value:
                str_value = str(value)
            if cell_val:
                str_cell_val = str(cell_val)

            rate, d = similar_std_instance(str_cell_val, str_value)
            row_values.append(cell_val)
            row_rates.append(rate)

        # get overall rates from row_matches
        overall_rate = validation_func(row_rates)
        # test rates, if not test allow best match
        test = not min_rate or overall_rate >= min_rate

        if process_all and test:
            # additional processed data which takes resources are calculated if necessary
            address = base_address + f"${letter}${row_no}"
            match.append((row_values, address, row_no, overall_rate, test, row_rates))
        else:
            match.append((row_values, None, None, overall_rate, test, row_rates))

    # sort by most likely
    mode_rate = modes.index("rate")
    match.sort(key=lambda x: x[mode_rate], reverse=True)

    # debug data
    if debugger.active:
        # rates, values, addresses = list(zip(*match))
        # TODO
        db = [pd.DataFrame(match, columns=[i.upper() for i in modes])]
    else:
        db = []

    if not match:  # if not match found
        return Special.nan, db

    # get raw columns and values to be expanded
    columns = [modes[i] for i in _modes]

    # get slice from the valid tests
    mode_test = modes.index("test")
    if to_slice is None:
        # calculate to get all valid values
        latest = None
        for i, val in enumerate(match):
            if not val[mode_test] and latest is None:
                latest = i
                break
        if latest is None:
            latest = 1
        to_slice = slice(0, latest)

    values_re = [[i[j] for j in _modes] for i in match[to_slice]]

    # get expansions and expand columns:
    sample = values_re[0]
    expansions = []
    columns_ex = []
    for col_num, col in enumerate(sample):
        if isinstance(col, list):
            # expand
            extend = [columns[col_num]+str(i) for i, _ in enumerate(col)]
            expansions.append(len(extend))
            columns_ex.extend(extend)
        else:
            expansions.append(0)
            columns_ex.append(columns[col_num])

    # expand values
    values_ex = []
    nanable_modes = {"value", "address", "row"}
    for i, row in enumerate(values_re):

        # replace the ones that does not pass test with #N/A
        to_replace_as_nan = not match[i][mode_test]

        # expand sub-columns to main columns
        row_expansion = []
        for j, (expand, value) in enumerate(zip(expansions, row)):
            if expand:
                # more than 1 column, must be expanded
                if len(value) != expand:
                    raise ValueError("Expansion '{value}' does not correspond to values {value} in row {i}")
                if to_replace_as_nan and modes[j] in nanable_modes:
                    row_expansion.extend([Special.nan]*expand)
                else:
                    row_expansion.extend(value)
            else:
                # single column
                if to_replace_as_nan and modes[j] in nanable_modes:
                    row_expansion.append(Special.nan)
                else:
                    row_expansion.append(value)
        values_ex.append(row_expansion)

    # create dataframe with expanded values
    #return pd.DataFrame(values_ex, columns=columns_ex), db
    if header:
        return [columns_ex] + values_ex, db
    return values_ex, db


#TODO xl_combinations lest de user specify a table were columns are the desired combinations, blank cell are ignored
@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('table', ndim=2, doc='table to apply combinations (randomly) in rows')
@xw.arg('columns', doc='specific columns to only apply combinations. Use "*" for all.')
@xw.arg('has_headers', doc='True to treat first row as headers else use in combinations')
@xw.ret(expand='table')  #
def xl_combinations(table, columns=None, has_headers=None):
    raise  NotImplementedError
    # order does not matters

    # random rows mth1
    df.reindex(np.ramdom.permutations(df.index))  # columns can be added
    # random rows mth2
    df.sample(df.shape[0])

    # random any column
    selected = df[column]
    df[column] = selected[np.ramdom.permutations(len(selected))].values
    pass


@debugger.get_steps()
#@debugger.debug(types=True)
@xw.func(**general_configuration)
@xw.arg('columns', ndim=2, doc='columns of data to permute (orderly)')
@xw.arg('has_headers', doc='True to treat first row as headers else as data')
@xw.arg('ignore_na', doc='True to treat empty data has a permutation candidate')
@xw.ret(expand='table')  #
def xf_permutations(columns, has_headers=None, ignore_na=None):
    # oder matters https://www.mathsisfun.com/combinatorics/combinations-permutations.html
    # columns is of the form [row_1, row_2, ..., row_N]
    # convert to _columns of the form [column_1, column_2, ..., column_N] and get headers
    _columns = [list(i) for i in zip(*columns)]

    headers = [col.pop(0) for col in _columns] if has_headers else None

    # clean up empty rows
    new_columns = [[row for row in col if row] for col in _columns] if ignore_na else _columns

    permuted = list(itertools.product(*new_columns))

    if headers:
        permuted.insert(0, headers)

    return permuted, [columns, new_columns]


# --------------------------- PRODUCTION DEBUGGING ---------------------------


@xw.func
#@xw.arg("data", MyConverter)
def xf_test_data(data):
    """print data type and content to terminal"""
    info = F"{type(data)}\n{data}"
    print(info)
    return info


@xw.func
@xw.arg("test_names", doc="unittest to get and print. '*' to print all")
@xw.ret(expand='table', cache=False)  #
def xf_print_unittest(test_names=None):
    """get the unittest test of a collected function evaluations and print it to terminal"""
    if test_names is None:
        return [f"tests available to print: {', '.join(test_creator.tests.keys())}"]
    if test_names is True or test_names.strip() == "*":
        test_names = list(test_creator.tests.keys())
    if not isinstance(test_names, (list, tuple)):
        test_names = [test_names]

    formatted = [test_creator.print_unittest(i) for i in test_names]
    return formatted


@xw.func
@xw.arg("test_names", doc="unittest name to delete. '*' to delete all")
@xw.ret(cache=False)  #
def xf_delete_unittest(test_names=None):
    """delete unittest test of a collected function evaluation"""
    if test_names is None:
        return f"tests available to delete: {', '.join(test_creator.tests.keys())}"
    if test_names is True or test_names.strip() == "*":
        test_names = list(test_creator.tests.keys())
    if not isinstance(test_names, (list, tuple)):
        test_names = [test_names]
    # check first
    for test_name in test_names:
        try:
            del test_creator.tests[test_name]
            return [f"deleted {test_name}"]
        except KeyError:
            raise KeyError(f"test '{test_name}' does not exists")
    # delete
    deletions = []
    for test_name in test_names:
        try:
            del test_creator.tests[test_name]
            deletions.append(test_name)
        except KeyError:
            raise KeyError(f"test '{test_name}' does not exists")
    return f"tests deleted: {', '.join(deletions)}"


@xw.func
@xw.arg("collection_state", doc="True to start collection. False to stop collection")
@xw.ret(cache=False)  #
def xf_collect_unittest(collection_state=None):
    """set unittest test case collection on or off"""
    if collection_state is not None:
        debugger.collect_tests = collection_state
    new_state = "unittest collection is active" if debugger.collect_tests else "unittest collection is deactivated"
    return new_state


@xw.func
@xw.arg("formula", xw.Range, doc='formula to show')
@xw.ret(expand='table', cache=False)  # prevent cache to be used
def xf_show_formula(formula):
    """show formula from cell or cells"""
    info = formula.formula  # FIXME real equation is not printed i.e. ',' instead of ";" or some values change
    if not isinstance(info, (list, tuple)):
        return [formula.formula_array]  # FIXED partial fix
    else:
        return info  # TODO replace each cell for formula.formula_array of each involved cell


@xw.func
@xw.arg('data', doc='excel array to be interpreted')
@xw.arg('xl_app', vba='Application')
@xw.ret(expand='table')
def xf_evaluate(data, xl_app):
    """mirror array data"""
    # use Evaluate https://docs.microsoft.com/en-us/office/vba/api/excel.application.evaluate
    # Evaluate for arrays https://newtonexcelbach.com/2018/08/24/using-vba-evaluate-as-an-array-function/
    ret = xl_app.Evaluate(data)  # TODO allow to interpret arrays i.e {"A";"B"} it seem that Array exists in App
    if -2146826273 == ret:
        raise ValueError("input is too long to evaluate. it must be 255 characters")
    return ret


@xw.func
@xw.arg('arrays', doc='cells to be converted to array')
@xw.ret(expand='table', header=False, index=False, empty='NA')  #
def xf_make_array(*arrays):
    return list(arrays)


@xw.func
def xf_get_selection_range():
    # TODO use this to get caller address in a thread, then if the address is in a function diagnose it
    # FIXME dangerous as it can be a range if where the selection is even though equation is elsewhere
    # by showing all the steps and source data
    wb = xw.Book.caller()
    return wb.selection.address  # xl_app.Caller.Address


@xw.func
@xw.arg('xl_app', vba='Application')
def xf_get_caller_address(xl_app):
    return xl_app.Caller.Address


@xw.func
@xw.arg('xl_app', vba='Application')
def xf_recalculate_udfs(xl_app):
    """
    equivalent to CalculateFull

    :param xl_app:
    :return:
    """
    if not calculate_all.calculating:
        calculate_all.calculating = True
        try:
            # FIXME this does not returns as it seem recalculation also recalcules this function
            #caller = xl_app.caller
            #xw_caller = xw.Range(impl=xw.xlplatform.Range(xl=caller))
            #range_selection = xw_caller.get_address(True, False, external=True)
            ##equation_address = F"'[{xw_caller.sheet.book.name}]{xw_caller.sheet.name}'!{xw_caller.address}"
            #cell_formula = xw.Range(range_selection)
            #cell_formula.value = ""  # FIXME this seems to never return
            ##xw_caller.value = ""
            try:
                xl_app.CalculateFull()
                return "recalculation done..."
            except com_error as e:
                return "error recalculating..."  # FIXME it seems recalculation fails because it is on UDF, try to set task
        finally:
            calculate_all.calculating = False
    else:
        return "another recalculation ongoing..."


def load_udf_modules(xl_app):
    # FIXME import_udfs adds extra modules with repeated functions. It seems it is because while program is here excel
    #   cannot replace xlwings_udfs creating a duplicate.
    # TODO maybe solution is to set a task to call macro of import functions https://stackoverflow.com/a/30312679/5288758
    # import module to workbook
    #book = xl_app.ActiveWorkbook  #xw.Range(impl=xw.xlplatform.Range(xl=xl_app.caller)).sheet.book  # xl_app.book
    #import_udfs(";".join(udf_modules), book)  # this has to add all the available modules
    return


@xw.func
@xw.arg("library", doc="path or name of library to load")
@xw.arg('xl_app', vba='Application')
def xf_load_library(library, xl_app):

    print(F"adding library reference: '{library}'...")
    register_library(library)

    load_udf_modules(xl_app)
    return f"added reference for load: {library}"


@xw.func
@xw.arg("library", doc="path or name of library to load")
@xw.arg('xl_app', vba='Application')
def xf_delete_library(library, xl_app):
    """delete external library"""

    print(F"deleting library reference: '{library}'...")
    try:
        delete_library(library)
    except KeyError:
        return f"non existent: {library}"

    load_udf_modules(xl_app)
    return f"deleted reference for load: {library}"


def serve(debug=None):

    # TODO load configuration file for server and operations modes

    # to debug
    debugger.active = debug  # deactivate debugger
    watcher.watch()

    # serve
    xw.serve()


if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    #xw.Book('excel_enhancer.xlsm').set_mock_caller()
    
    serve()
    print("excel_enhancer server started...")
else:
    print("excel_enhancer udf module loaded...")
