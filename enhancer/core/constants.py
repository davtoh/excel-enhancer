"""
Excel Enhancer special constants and error codes.
"""

import numpy as np


# class keeping special values to standardize them on VBA functions results or interpretations
class Special:
    # https://github.com/xlwings/xlwings/issues/776#issuecomment-476183843
    # this needs to be something like sheet.range("A5").formula = '#N/A'
    # all errors https://support.microsoft.com/en-us/office/formulas-and-functions-294d9486-b332-48ed-b489-abe7d0f9eda9?ui=en-US&rs=en-US&ad=US#id0eaabaaa=errors

    # The #N/A error generally indicates that a formula can’t find what it’s been asked to look for.
    nan = empty = '#N/A'

    err_div = '#DIV/0!'

    # The top reason why the #NAME? error appears in your formula is because there is a typo in the formula name
    err_name = '#NAME?'

    err_null = '#NULL!'

    # if function contains numeric values that aren't valid
    err_num = '#NUM!'

    # The #REF! error shows when a formula refers to a cell that’s not valid
    err_ref = '#REF!'

    # #VALUE! There's something wrong with the way your formula is typed. Or,
    # there's something wrong with the cells you are referencing.
    err_value = '#VALUE!'

    error_dict = {-2146826281: np.inf, -2146826246: np.nan}
    row_sep = ';'  # separator for rows in excel equation

