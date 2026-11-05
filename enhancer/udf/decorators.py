"""
Function decorators and signature inspection for Excel UDFs.

Provides decorators for declaring Excel User Defined Functions (UDFs),
subs, arguments, and return value converters compatible with xlwings.
"""

import inspect
from typing import Any, Callable


def func_sig(f: Callable) -> dict[str, Any]:
    """Inspect and return the argument signature of a function."""
    s = inspect.signature(f)
    vararg = None
    args = []
    defaults = []
    for p in s.parameters.values():
        if p.kind is inspect.Parameter.POSITIONAL_OR_KEYWORD:
            args.append(p.name)
            if p.default is not inspect.Signature.empty:
                defaults.append(p.default)
        elif p.kind is inspect.Parameter.VAR_POSITIONAL:
            args.append(p.name)
            vararg = p.name
        else:
            raise Exception("xlwings does not support UDFs with keyword arguments")
    return {
        'args': args,
        'defaults': defaults,
        'vararg': vararg
    }


def get_category(**func_kwargs) -> int | str:
    """Validate and extract the Excel function category."""
    if 'category' in func_kwargs:
        category = func_kwargs.pop('category')
        if isinstance(category, int):
            if 1 <= category <= 14:
                return category
            raise Exception(
                'There is only 14 build-in categories available in Excel. '
                'Please use a string value to specify a custom category.')
        if isinstance(category, str):
            return category[:255]
        raise Exception(
            'Category {0} should either be a predefined Excel category (int value) or a custom one (str value).'.format(
                category))
    return "xlwings"  # Default category


def get_async_mode(**func_kwargs) -> str | None:
    """Validate and extract the asynchronous execution mode."""
    if 'async_mode' in func_kwargs:
        value = func_kwargs.pop('async_mode')
        if value in ['threading']:
            return value
        raise Exception('The only supported async_mode mode is currently "threading".')
    return None


def check_bool(kw: str, **func_kwargs) -> bool:
    """Validate and extract a boolean flag from kwargs."""
    if kw in func_kwargs:
        check = func_kwargs.pop(kw)
        if isinstance(check, bool):
            return check
        raise Exception('{0} only takes boolean values. ("{1}" provided).'.format(kw, check))
    return False


def _build_arg_descriptors(sig: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[str, dict[str, Any]]]:
    """Build argument descriptor list and map from a function signature."""
    xlargs: list[dict[str, Any]] = []
    xlargmap: dict[str, dict[str, Any]] = {}
    n_args = len(sig['args'])
    n_defaults = len(sig['defaults'])
    n_required_args = n_args - n_defaults

    if sig['vararg'] and n_defaults > 0:
        raise Exception("xlwings does not support UDFs with both optional and variable length arguments")

    for vpos, vname in enumerate(sig['args']):
        arg_info = {
            "name": vname,
            "pos": vpos,
            "vba": None,
            "doc": "Positional argument " + str(vpos + 1),
            "vararg": vname == sig['vararg'],
            "options": {}
        }
        if vpos >= n_required_args:
            arg_info["optional"] = sig['defaults'][vpos - n_required_args]
        xlargs.append(arg_info)
        xlargmap[vname] = arg_info

    return xlargs, xlargmap


def _init_xlfunc_meta(f: Callable) -> dict[str, Any]:
    """Initialize __xlfunc__ metadata dictionary on function if absent."""
    if not hasattr(f, "__xlfunc__"):
        sig = func_sig(f)
        xlargs, xlargmap = _build_arg_descriptors(sig)
        doc_str = (f.__doc__ if f.__doc__ is not None
                   else "Python function '" + f.__name__ + "' defined in '" + str(f.__code__.co_filename) + "'.")
        f.__xlfunc__ = {
            "name": f.__name__,
            "sub": False,
            "args": xlargs,
            "argmap": xlargmap,
            "ret": {
                "doc": doc_str,
                "options": {}
            }
        }
    return f.__xlfunc__


def xlfunc(f=None, **kwargs):
    """Decorator for Excel UDF functions."""
    def inner(func):
        xlf = _init_xlfunc_meta(func)
        xlf["category"] = get_category(**kwargs)
        xlf['call_in_wizard'] = check_bool('call_in_wizard', **kwargs)
        xlf['volatile'] = check_bool('volatile', **kwargs)
        xlf['async_mode'] = get_async_mode(**kwargs)
        return func

    if f is None:
        return inner
    return inner(f)


def xlsub(f=None, **kwargs):
    """Decorator for Excel VBA subroutines."""
    def inner(func):
        func = xlfunc(**kwargs)(func)
        func.__xlfunc__["sub"] = True
        return func

    if f is None:
        return inner
    return inner(f)


def xlret(convert=None, **kwargs):
    """Decorator for return value conversion options."""
    if convert is not None:
        kwargs['convert'] = convert

    def inner(f):
        xlf = xlfunc(f).__xlfunc__
        xlr = xlf["ret"]
        xlr['options'].update(kwargs)
        return f

    return inner


def xlarg(arg, convert=None, **kwargs):
    """Decorator for parameter conversion options and metadata."""
    if convert is not None:
        kwargs['convert'] = convert

    def inner(f):
        xlf = xlfunc(f).__xlfunc__
        if arg not in xlf["argmap"]:
            raise Exception("Invalid argument name '" + arg + "'.")
        xla = xlf["argmap"][arg]
        for special in ('vba', 'doc'):
            if special in kwargs:
                xla[special] = kwargs.pop(special)
        xla['options'].update(kwargs)
        return f

    return inner

