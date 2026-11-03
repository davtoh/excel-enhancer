"""
Excel validation and recalculation macros.
"""

import xlwings as xw


def create_validation(cell, formula1):
    """
    function used to create drop down lists in excel

    :param cell: range cell to create validation
    :param formula1: string of address or a vector e.g. "A1:A5", "data 1; data 2"
    :return:
    """
    validation = cell.api.Validation
    validation.delete()  # delete any previous drop down list
    # add a drop down list
    validation.Add(Type=3, AlertStyle=1, Operator=1, Formula1=formula1)
    validation.IgnoreBlank = -1
    validation.InCellDropdown = -1
    validation.InputTitle = ""
    validation.ErrorTitle = ""
    validation.InputMessage = ""
    validation.ErrorMessage = ""
    validation.ShowInput = -1
    validation.ShowError = -1


@xw.sub
def calculate_all():
    """
    macro to calculate all. Equivalent to :

    Private Sub Workbook_Open()
    Application.CalculateFull
    End Sub
    """
    if not calculate_all.calculating:
        calculate_all.calculating = True
        try:
            wb = xw.Book.caller()
            wb.app.api.CalculateFull()
            return "recalculation done..."
        finally:
            calculate_all.calculating = False
    else:
        return "another recalculation ongoing..."


# global flag to indicate if recalculating
calculate_all.calculating = False

