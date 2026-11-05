"""
Debugging, introspection, testing, and macro utilities for Excel Enhancer.
"""

import xlwings as xw
from enhancer.core import (
    calculate_all,
    debugger,
    delete_library,
    register_library,
    test_creator,
)
from enhancer.server.runner import load_udf_modules

try:
    from pywintypes import com_error
except ImportError:
    com_error = Exception


@xw.func
def xf_test_data(data):
    """print data type and content to terminal"""
    info = f"{type(data)}\n{data}"
    print(info)
    return info


@xw.func
@xw.arg("test_names", doc="unittest to get and print. '*' to print all")
@xw.ret(expand="table", cache=False)
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
@xw.ret(cache=False)
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

    # delete remaining
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
@xw.ret(cache=False)
def xf_collect_unittest(collection_state=None):
    """set unittest test case collection on or off"""
    if collection_state is not None:
        debugger.collect_tests = collection_state
    new_state = (
        "unittest collection is active"
        if debugger.collect_tests
        else "unittest collection is deactivated"
    )
    return new_state


@xw.func
@xw.arg("formula", xw.Range, doc="formula to show")
@xw.ret(expand="table", cache=False)
def xf_show_formula(formula):
    """show formula from cell or cells"""
    info = formula.formula
    if not isinstance(info, (list, tuple)):
        return [formula.formula_array]
    return info


@xw.func
@xw.arg("data", doc="excel array to be interpreted")
@xw.arg("xl_app", vba="Application")
@xw.ret(expand="table")
def xf_evaluate(data, xl_app):
    """mirror array data"""
    ret = xl_app.Evaluate(data)
    if ret == -2146826273:
        raise ValueError("input is too long to evaluate. it must be 255 characters")
    return ret


@xw.func
@xw.arg("arrays", doc="cells to be converted to array")
@xw.ret(expand="table", header=False, index=False, empty="NA")
def xf_make_array(*arrays):
    """convert cells to array list"""
    return list(arrays)


@xw.func
def xf_get_selection_range():
    """get active caller selection address"""
    wb = xw.Book.caller()
    return wb.selection.address


@xw.func
@xw.arg("xl_app", vba="Application")
def xf_get_caller_address(xl_app):
    """get caller address from Excel application object"""
    return xl_app.Caller.Address


@xw.func
@xw.arg("xl_app", vba="Application")
def xf_recalculate_udfs(xl_app):
    """equivalent to CalculateFull"""
    if not calculate_all.calculating:
        calculate_all.calculating = True
        try:
            try:
                xl_app.CalculateFull()
                return "recalculation done..."
            except com_error:
                return "error recalculating..."
        finally:
            calculate_all.calculating = False
    else:
        return "another recalculation ongoing..."


@xw.func
@xw.arg("library", doc="path or name of library to load")
@xw.arg("xl_app", vba="Application")
def xf_load_library(library, xl_app):
    """register external UDF library and reload modules"""
    print(f"adding library reference: '{library}'...")
    register_library(library)
    load_udf_modules(xl_app)
    return f"added reference for load: {library}"


@xw.func
@xw.arg("library", doc="path or name of library to load")
@xw.arg("xl_app", vba="Application")
def xf_delete_library(library, xl_app):
    """delete external library"""
    print(f"deleting library reference: '{library}'...")
    try:
        delete_library(library)
    except KeyError:
        return f"non existent: {library}"

    load_udf_modules(xl_app)
    return f"deleted reference for load: {library}"

