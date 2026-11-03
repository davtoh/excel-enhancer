"""
Dynamic module loading, function publishing, and UDF library registration.
"""

import os
import sys
import importlib
import importlib.util
from xlwings.udfs import udf_modules


def publish_functions(module, destine=None):
    """
    make all the module function available in globals

    :param module: module to get functions from
    :param destine: dictionary to publish functions
    :return: names of the published functions
    """
    # https://stackoverflow.com/a/43059528/5288758
    # get a handle on the module
    mdl = module  # importlib.import_module(module)

    # is there an __all__?  if so respect it
    if "__all__" in mdl.__dict__:
        names = mdl.__dict__["__all__"]
    else:
        # otherwise we import all names that don't begin with _
        names = [x for x in mdl.__dict__ if not x.startswith("_")]

    # now drag them in
    if destine is None:
        destine = globals()
    destine.update({k: getattr(mdl, k) for k in names})
    return names


def get_module_name(module):
    """
    get normalized module name

    :param module: module name, module or path to module file
    :return: module name
    """
    if isinstance(module, str):
        if os.path.isfile(module):
            path = module
            root, ext = os.path.splitext(path)
            root, basename = os.path.split(root)
        else:
            basename = module  # name of the module provided
    else:
        basename = module.__name__
    return basename


def load_module(module, reload=None):
    """
    import and reload from custom external modules

    :param module: module, module name or path to module file
    :param reload:
    :return:
    """
    # FIXED replace with a custom loader for user to externally decide modules to be imported and used dynamically
    if isinstance(module, str):
        if os.path.isfile(module):
            path = module
            root, ext = os.path.splitext(path)
            root, basename = os.path.split(root)
            spec = importlib.util.spec_from_file_location(basename, path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            module_name = module.__name__

            # register module in sys only if reload or no registered before
            if reload or module_name not in sys.modules:
                sys.modules[module_name] = module
        else:
            module_name = module  # name of the module provided
            # load module so that we have it present for reload
            module = importlib.import_module(module_name)
        if reload:
            importlib.reload(module)  # FIX ME, does not work if module was loaded from path and path not in sys.path
    return module


def reload_functions(module):
    """
    load module on demand and update namespace

    # equivalent to: from excel_enhancer_xf import *
    >>> reload_functions("udf_json_support")
    >>> reload_functions("./extensions/udf_json_support.py")

    :param module:
    :return:
    """
    # load module
    module = load_module(module)

    # propagate
    publish_functions(module)


def register_library(library):
    """
    load external library reference to the udf modules that must be imported in excel
    so that next import adds these libraries

    :param library: module, module name or path to module file
    :return:
    """
    # get module from user
    module = load_module(library, reload=False)  # FIXME reload=True not ready

    # register module
    module_name = module.__name__
    filename = os.path.normcase(module.__file__.lower())

    try:  # getmtime fails for zip imports and frozen modules
        mtime = os.path.getmtime(filename)
    except OSError:
        mtime = None

    udf_modules[module_name] = {
        'filename': filename,
        'filetime': mtime,
        'module': module
    }


def delete_library(library):
    """
    delete external library reference to the udf modules that must be imported in excel
    so that next import deletes these libraries

    :param library: module, module name or path to module file
    :return:
    """
    # get module from user
    module_name = get_module_name(library)
    del udf_modules[module_name]

