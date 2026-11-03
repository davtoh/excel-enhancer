"""
Range accessors, bounding box adjusters, and xlwings converters.
"""

import pandas as pd
import xlwings as xw
from xlwings.conversion.standard import (
    BaseAccessor, ValueAccessor, ReadValueFromRangeStage, Ensure2DStage,
    CleanDataFromReadStage, TransposeStage, AdjustDimensionsStage,
    Pipeline, WriteValueToRangeStage, CleanDataForWriteStage
)
from xlwings.conversion import Converter, PandasDataFrameConverter

from enhancer.core.text_utils import helper_correct_to_str


def is_range_infinite(rng):
    """determine if a range is infinite e.g A:C"""
    address = rng.address
    is_finite = any([str(i) in address for i in range(10)])
    return not is_finite


def range_bounding_box(rng):
    """
    get bounding box to fit contents of table

    :param rng: range
    :return: first_row, last_row, first_col, last_col
    """
    # https://github.com/xlwings/xlwings/pull/1177#issuecomment-558322707
    # x.sheet[row_range,column_range]
    # https://stackoverflow.com/a/53102362/5288758
    # x.sheet.range("B"+str(x.last_cell.row)).end("up").row  # fit the data
    last_row = rng.last_cell.row  # get last row
    sampled_rows = []
    last_column = 0  # counter for last column
    for col in (i.column for i in rng.columns):  # see columns
        last_column = col
        last_range_str = rng.sheet[0, last_column - 1].address.replace("1", str(last_row))
        sampled_rows.append(rng.sheet.range(last_range_str).end("up").row)
    last_row_sampled = max(sampled_rows)
    return rng.row - 1, last_row_sampled, rng.column - 1, last_column


def fit_table(rng):
    """
    adjust range to fit contents of table in the rows so that infinite ranges like A:B
    don't go to infinite but the actual last row with data e.g. A1:B500

    :param rng: range
    :return: fitted_range
    """
    first_row, last_row, first_col, last_col = range_bounding_box(rng)
    return rng.sheet[first_row:last_row, first_col:last_col]


class ReadValueFromRangeStageInfiniteAdjusted:
    """Converter to fit values to table data from infinite range. Finite ranges are untouched"""

    def __call__(self, c):
        if c.range:
            if is_range_infinite(c.range):
                fitted_range = fit_table(c.range)
                # raw_range changes None to '' automatically
                c.value = fitted_range.raw_value
            else:
                c.value = c.range.raw_value


class ConverterAdjustInfiniteRange:
    """Converter to fit infinite range to finite range using table data. Finite ranges are untouched"""

    def __call__(self, c):
        if c.range:
            if is_range_infinite(c.range):
                c.range = fit_table(c.range)


class EnsureDataIsWritable:

    def __call__(self, c):
        # value must be a list of lists
        if isinstance(c.value, (list, tuple)) and isinstance(c.value[0], (list, tuple)):
            found_invalid = helper_correct_to_str(c.value)
            if found_invalid:
                print("values on invalid output corrected")


class ValueAccessorFitted(ValueAccessor):
    """
    Accessor that ensures values of tables fitting infinite rows to table data
    so that data result data don't have many None values and take time to convert
    """

    @staticmethod
    def reader(options):
        return (
            BaseAccessor.reader(options)
            .append_stage(ConverterAdjustInfiniteRange())
            .append_stage(ReadValueFromRangeStage())
            .append_stage(Ensure2DStage())
            .append_stage(CleanDataFromReadStage(options))
            .append_stage(TransposeStage(), only_if=options.get('transpose', False))
            .append_stage(AdjustDimensionsStage(options))
        )

    @staticmethod
    def writer(options):
        return (
            Pipeline()
            #.append_stage(EnsureDataIsWritable())
            .prepend_stage(WriteValueToRangeStage(options))
            .prepend_stage(CleanDataForWriteStage())
            .prepend_stage(TransposeStage(), only_if=options.get('transpose', False))
            .prepend_stage(Ensure2DStage())
        )


ValueAccessorFitted.register(None)  # register as default


class MyConverter(Converter):
    # TODO, this is a test

    @staticmethod
    def read_value(value, options):
        myoption = options.get('myoption', None)
        return_value = value  # Implement your conversion here
        return return_value

    @staticmethod
    def write_value(value, options):
        myoption = options.get('myoption', None)
        return_value = value  # Implement your conversion here
        return return_value

    @classmethod
    def reader(self, options):
        base_reader = self.base_reader(options)
        # modify stage of reading default
        base_reader.insert(0, ConverterAdjustInfiniteRange())
        return (
            base_reader.append_stage(self.FromValueStage(self.read_value, options))
        )


def pandas_dataframe_converter(value, *args, **kwargs):
    """ pandas converter from xlwings for particular conversions when needed"""
    if isinstance(value, pd.DataFrame):
        return value
    if args and not kwargs:
        kwargs = args[0]
    elif args:
        raise ValueError(f"cannot deal with options args {args} and kwargs {kwargs}")
    return PandasDataFrameConverter.read_value(value, kwargs)

