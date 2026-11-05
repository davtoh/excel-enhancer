"""
VBA wrapper generation and macro registration for Excel UDFs.

Generates VBA boilerplate modules for declared Python UDFs and imports
them into active Excel workbooks with macro descriptions and argument documentation.
"""

import os
import re
import tempfile
from types import ModuleType
from typing import Any

import pywintypes
import win32api
from xlwings.utils import VBAWriter

from enhancer.udf.module_registry import get_udf_module, udf_modules


def _build_vba_signature(fname: str, ftype: str, xlfunc: dict[str, Any]) -> tuple[str, str]:
    """Construct VBA Sub/Function declaration header with argument specifiers."""
    func_sig = f"{ftype} {fname}("
    first = True
    vararg = ''
    for arg in xlfunc['args']:
        if not arg['vba']:
            argname = arg['name']
            if not first:
                func_sig += ', '
            if 'optional' in arg:
                func_sig += 'Optional '
            elif arg['vararg']:
                func_sig += 'ParamArray '
                vararg = argname
            func_sig += argname
            if arg['vararg']:
                func_sig += '()'
            first = False
    func_sig += ')'
    return func_sig, vararg


def _write_vba_args(vba: VBAWriter, xlfunc: dict[str, Any], vararg: str) -> str:
    """Write VBA argument array construction logic and return VBA args expression."""
    if vararg != '':
        vba.writeln("Dim argsArray() As Variant")
        non_varargs = [arg['vba'] or arg['name'] for arg in xlfunc['args'] if not arg['vararg']]
        vba.writeln("argsArray = Array(%s)" % tuple({', '.join(non_varargs)}))
        n_fixed = str(len(non_varargs))
        vba.writeln(f"ReDim Preserve argsArray(0 to UBound({vararg}) - LBound({vararg}) + {n_fixed})")
        vba.writeln(f"For k = LBound({vararg}) To UBound({vararg})")
        vba.writeln(f"argsArray({n_fixed} + k - LBound({vararg})) = {vararg}(k)")
        vba.writeln("Next k")
        return 'argsArray'
    return 'Array(' + ', '.join(arg['vba'] or arg['name'] for arg in xlfunc['args']) + ')'


def _write_vba_sub_invocation(vba: VBAWriter, module_name: str, fname: str, args_vba: str) -> None:
    """Emit Py.CallUDF dispatch lines for a VBA Subroutine."""
    with vba.block('#If App = "Microsoft Excel" Then'):
        vba.writeln(
            'Py.CallUDF "{module_name}", "{fname}", {args_vba}, ThisWorkbook, Application.Caller',
            module_name=module_name, fname=fname, args_vba=args_vba)
    with vba.block("#Else"):
        vba.writeln('Py.CallUDF "{module_name}", "{fname}", {args_vba}',
                    module_name=module_name, fname=fname, args_vba=args_vba)
    vba.writeln("#End If")


def _write_vba_func_invocation(vba: VBAWriter, module_name: str, fname: str, args_vba: str) -> None:
    """Emit Py.CallUDF dispatch lines and error handler for a VBA Function."""
    with vba.block('#If App = "Microsoft Excel" Then'):
        vba.writeln("If TypeOf Application.Caller Is Range Then On Error GoTo failed")
        vba.writeln(
            '{fname} = Py.CallUDF("{module_name}", "{fname}", {args_vba}, ThisWorkbook, Application.Caller)',
            module_name=module_name, fname=fname, args_vba=args_vba)
        vba.writeln("Exit Function")
    with vba.block("#Else"):
        vba.writeln('{fname} = Py.CallUDF("{module_name}", "{fname}", {args_vba})',
                    module_name=module_name, fname=fname, args_vba=args_vba)
        vba.writeln("Exit Function")
    vba.writeln("#End If")
    vba.write_label("failed")
    vba.writeln(f"{fname} = Err.Description")


def _write_single_vba_wrapper(vba: VBAWriter, svar: Any, module_name: str) -> None:
    """Emit full VBA wrapper block for a single __xlfunc__ decorated item."""
    xlfunc = svar.__xlfunc__
    fname = xlfunc['name']
    ftype = 'Sub' if xlfunc['sub'] else 'Function'
    func_sig, vararg = _build_vba_signature(fname, ftype, xlfunc)

    with vba.block(func_sig):
        if ftype == 'Function':
            if not xlfunc['call_in_wizard']:
                vba.writeln('If (Not Application.CommandBars("Standard").Controls(1).Enabled) Then Exit Function')
            if xlfunc['volatile']:
                vba.writeln('Application.Volatile')

        args_vba = _write_vba_args(vba, xlfunc, vararg)
        if ftype == 'Sub':
            _write_vba_sub_invocation(vba, module_name, fname, args_vba)
        else:
            _write_vba_func_invocation(vba, module_name, fname, args_vba)

    vba.writeln(f'End {ftype}\n')


def generate_vba_wrapper(module_name: str, module: ModuleType, f: Any) -> None:
    """Generate VBA wrappers for all decorated functions in a module."""
    vba = VBAWriter(f)
    for attr in dir(module):
        svar = getattr(module, attr)
        if hasattr(svar, '__xlfunc__'):
            _write_single_vba_wrapper(vba, svar, module_name)


def _generate_vba_temp_file(module_names: list[str], xl_workbook: Any) -> str:
    """Write generated VBA wrapper code to a temporary file and return its path."""
    tf = tempfile.NamedTemporaryFile(mode='w', delete=False)
    vba = VBAWriter(tf.file)
    vba.writeln('Attribute VB_Name = "xlwings_udfs"')
    vba.writeln("'Autogenerated code by xlwings - changes will be lost with next import!")
    vba.writeln("""#Const App = "Microsoft Excel" 'Adjust when using outside of Excel""")

    for module_name in module_names:
        mod = get_udf_module(module_name, xl_workbook)
        generate_vba_wrapper(module_name, mod, tf.file)

    tf.close()
    return tf.name


def _register_single_macro(xl_workbook: Any, fname: str, fdoc: str,
                           fcategory: Any, argdocs: list[str] | None, is_modern: bool) -> None:
    """Register macro options for a single UDF in the Excel application."""
    try:
        if is_modern:
            xl_workbook.Application.MacroOptions(
                "'" + xl_workbook.Name + "'!" + fname,
                Description=fdoc,
                HasMenu=False,
                MenuText=None,
                HasShortcutKey=False,
                ShortcutKey=None,
                Category=fcategory,
                StatusBar=None,
                HelpContextID=None,
                HelpFile=None,
                ArgumentDescriptions=argdocs if argdocs else None)
        else:
            xl_workbook.Application.MacroOptions("'" + xl_workbook.Name + "'!" + fname, Description=fdoc)
    except pywintypes.com_error as e:
        try:
            msg = win32api.FormatMessage(e.hresult).strip()
            print(f"something went wrong adding {fname} macro: {msg}...")
        except Exception:
            print(f"something went wrong adding {fname} macro...")


def _register_module_macros(xl_workbook: Any, module_names: list[str]) -> None:
    """Register MacroOptions for all UDFs in the provided module list."""
    excel_version = [int(x) for x in re.split(r"[,.]", xl_workbook.Application.Version)]
    is_modern = excel_version[0] >= 14

    for module_name in module_names:
        mod = get_udf_module(module_name, xl_workbook)
        for attr in dir(mod):
            mvar = getattr(mod, attr)
            if hasattr(mvar, '__xlfunc__'):
                xlfunc = mvar.__xlfunc__
                argdocs = [arg['doc'][:255] for arg in xlfunc['args'] if not arg['vba']]
                _register_single_macro(
                    xl_workbook,
                    xlfunc['name'],
                    xlfunc['ret']['doc'][:255],
                    xlfunc['category'],
                    argdocs,
                    is_modern)


def import_udfs(module_names: str | list[str], xl_workbook: Any) -> None:
    """Import generated UDF VBA wrappers and register documentation in Excel."""
    modules: list[str] = module_names.split(';') if isinstance(module_names, str) else list(module_names)
    for mod_name in udf_modules:
        if mod_name not in modules:
            modules.append(mod_name)

    temp_path = _generate_vba_temp_file(modules, xl_workbook)
    try:
        try:
            xl_workbook.VBProject.VBComponents.Remove(xl_workbook.VBProject.VBComponents("xlwings_udfs"))
        except Exception:
            pass
        xl_workbook.VBProject.VBComponents.Import(temp_path)
        _register_module_macros(xl_workbook, modules)
    finally:
        try:
            os.unlink(temp_path)
        except Exception:
            pass

