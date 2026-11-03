"""
Text manipulation, reflection, and registration utilities.
"""

import unicodedata

general_configuration = {}  # configuration to modify function modes in @xw.func i.e. {"async_mode": 'threading'}
custom_group_ops = {}  # custom functions used in VBAs


def find_attribute_in_object_tree(att, obj, addrs="", max_re=None):
    """
    helper function to find attributes in an object

    >>> find_attribute_in_object_tree("calculate", xw, max_re= 10)

    :param att: attribute to find
    :param obj: object to look on
    :param addrs: starting address. Leave as "" if search is starting.
    :param max_re: maximun deep in tree
    :return:
    """
    if max_re is not None and addrs.count(".") > max_re:
        return
    try:
        got = getattr(obj, att, None)
        if got is not None:
            print(f"found {att} in address {addrs}")
    except Exception:
        pass

    for i in dir(obj):
        if not i.startswith("__"):  # ignore special methods
            if att.lower() in i.lower():
                print(f"found {i} in address {addrs}")
            try:
                got = getattr(obj, i, None)
                if got is not None:
                    find_attribute_in_object_tree(att, got, addrs + "." + i, max_re)
            except Exception:
                continue


def find_similar_attribute_in_object_tree(att, obj, addrs="", max_re=None):
    """
    helper function to find similar attributes in an object

    >>> find_similar_attribute_in_object_tree("calculate", xw, max_re= 10)

    :param att: attribute to find
    :param obj: object to look on
    :param addrs: starting address. Leave as "" if search is starting.
    :param max_re: maximun deep in tree
    :return:
    """
    if max_re is not None and addrs.count(".") > max_re:
        return

    for i in dir(obj):
        if not i.startswith("__"):  # ignore special methods
            if att.lower() in i.lower():
                print(f"found {i} in address {addrs}")
            try:
                got = getattr(obj, i, None)
                if got is not None:
                    find_attribute_in_object_tree(att, got, addrs + "." + i, max_re)
            except Exception:
                continue


def register(register_in=None, name=None):
    """
    register function in group. This is useful to make function available for a process or VBA that uses the group
    that is registered to.

    :param register_in: dictionary group. If None then register in default custom_group_ops in the module
    :param name: name to give in group
    :return: registered function
    """
    if register_in is None:
        register_in = custom_group_ops

    def to_register(f):
        """register function to """
        if name is None:
            register_in[f.__name__] = f
        else:
            register_in[name] = f
        return f

    return to_register


def strip_accents(s, throw=False):
    """helper function used to normalize text so that accents are ignored"""
    # https://stackoverflow.com/a/518232/5288758
    try:
        return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    except TypeError:
        if throw:
            raise
        return s


def col_name(n):
    # https://stackoverflow.com/a/23862195/5288758
    # replaces from xlwings.utils import col_name
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def convert_to_string(value):
    """
    convert value to string in the most simple way to be human readable
    """
    if value is None:
        return ""  # should not return anything

    # float case to handle e.g 0 instead of 0.0
    if isinstance(value, float) and value == int(value):
        return str(int(value))

    return str(value)


def helper_correct_to_str(values):
    """
    ensure lists hast all values as str

    :param values: 2D list
    :return: whether invalid values were found
    """
    found_invalid = False
    for row in values:
        for j, col in enumerate(row):
            if not isinstance(col, str):
                row[j] = str(col)
                found_invalid = True
    return found_invalid

