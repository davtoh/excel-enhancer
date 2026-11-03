"""
Server execution runtime and UDF module loading for Excel Enhancer.
"""

import xlwings as xw
from enhancer.core import debugger, watcher


def load_udf_modules(xl_app):
    """
    Import UDF modules into active Excel application.
    """
    # FIXME import_udfs adds extra modules with repeated functions. It seems it is because while program is here excel
    #   cannot replace xlwings_udfs creating a duplicate.
    # TODO maybe solution is to set a task to call macro of import functions https://stackoverflow.com/a/30312679/5288758
    # import module to workbook
    # book = xl_app.ActiveWorkbook
    # import_udfs(";".join(udf_modules), book)
    return


def serve(debug=None):
    """
    Start the Excel Enhancer UDF server loop.
    """
    # TODO load configuration file for server and operations modes
    # to debug
    debugger.active = debug  # deactivate debugger
    watcher.watch()

    # serve
    xw.serve()

