# -*- coding: utf-8 -*-

import six
from datetime import datetime, date

try:
    from itertools import zip_longest
except ImportError:
    from itertools import izip_longest as zip_longest

from gspread_formatting import format_cell_ranges, cellFormat, numberFormat, Color, textFormat, set_frozen
from gspread.utils import rowcol_to_a1

__all__ = (
    'format_with_dataframe', 
    'DataFrameFormatter', 
    'BasicFormatter', 
    'DEFAULT_FORMATTER', 
    'DEFAULT_HEADER_BACKGROUND_COLOR'
)

DEFAULT_HEADER_BACKGROUND_COLOR = Color(0.8980392, 0.8980392, 0.8980392)

def format_with_dataframe(worksheet,
                          dataframe,
                          formatter,
                          row=1,
                          col=1,
                          include_index=False,
                          include_column_header=True):
    """
    Sets the values of a given DataFrame, anchoring its upper-left corner
    at (row, col). (Default is row 1, column 1.)

    :param worksheet: the gspread worksheet to set with content of DataFrame.
    :param dataframe: the DataFrame.
    :param include_index: if True, include the DataFrame's index as an
            additional column. Defaults to False.
    :param include_column_header: if True, add a header row before data with
            column names. (If include_index is True, the index's name will be
            used as its column's header.) Defaults to True.
    :param formatter: an optional instance of WorksheetFormatter, which
                      will examine the contents of the DataFrame and
                      assemble a set of `gspread_formatter` operations
                      to be performed after the DataFrame contents 
                      are written to the given Worksheet. The formatting
                      operations are performed after the contents are written
                      and before this function returns.
    """
    formatting_ranges = []

    columns = [ dataframe[c] for c in dataframe.columns ]
    if include_index:
        columns = [ dataframe.index ] + columns

    for idx, column in enumerate(columns):
        column_fmt = formatter.format_for_column(column, col + idx, dataframe)
        if not column_fmt or not column_fmt.to_props():
            continue
        range = '{}:{}'.format(
            rowcol_to_a1(row, col + idx), 
            rowcol_to_a1(row + dataframe.shape[0], col + idx)
        )
        formatting_ranges.append( (range, column_fmt) )

    if include_column_header:
        elts = list(dataframe.columns)
        if include_index:
            elts = [ dataframe.index.name ] + elts
            header_fmt = formatter.format_for_header(dataframe.index, dataframe)
            if header_fmt:
                formatting_ranges.append( 
                    (
                        '{}:{}'.format(
                            rowcol_to_a1(row, col), 
                            rowcol_to_a1(row + dataframe.shape[0], col)
                        ), 
                        header_fmt
                    )
                )

        header_fmt = formatter.format_for_header(elts, dataframe)
        if header_fmt:
            formatting_ranges.append( 
                (
                    '{}:{}'.format(
                        rowcol_to_a1(row, col), 
                        rowcol_to_a1(row, col + dataframe.shape[1])
                    ), 
                    header_fmt
                )
            )

        freeze_args = {}

        if row == 1 and formatter.should_freeze_header(elts, dataframe):
            freeze_args['rows'] = 1

        if include_index and col == 1 and formatter.should_freeze_header(dataframe.index, dataframe):
            freeze_args['cols'] = 1

        row += 1


    values = []
    for value_row, index_value in zip_longest(dataframe.values, dataframe.index):
        if include_index:
            value_row = [index_value] + list(value_row)
        values.append(value_row)
    for y_idx, value_row in enumerate(values):
        for x_idx, cell_value in enumerate(value_row):
            cell_fmt = formatter.format_for_cell(cell_value, y_idx+row, x_idx+col, dataframe)
            if cell_fmt:
                formatting_ranges.append((rowcol_to_a1(y_idx+row, x_idx+col), cell_fmt))
        row_fmt = formatter.format_for_row(values, y_idx+row, dataframe)
        if row_fmt:
            formatting_ranges.append(
                (
                    '{}:{}'.format(
                        rowcol_to_a1(y_idx+row, col), 
                        rowcol_to_a1(y_idx+row, col+dataframe.shape[1])
                    ), 
                row_fmt
                )
            )

    if formatting_ranges:
        formatting_ranges = [ r for r in formatting_ranges if r[1] and r[1].to_props() ]
        format_cell_ranges(worksheet, formatting_ranges)

    if freeze_args:
        set_frozen(worksheet, **freeze_args)

class DataFrameFormatter():
    @classmethod
    def resolve_number_format(cls, value, type=None):
        if value is None:
            return None
        elif isinstance(value, numberFormat):
            return value
        elif isinstance(value, str):
            return numberFormat(type, value) 
        else:
            raise ValueError(value)

    def format_with_dataframe(self, worksheet, dataframe, row=1, col=1, include_index=False, include_column_header=True):
        return format_with_dataframe(
            worksheet, 
            dataframe, 
            self, 
            row=row, 
            col=col, 
            include_index=include_index, 
            include_column_header=include_column_header
        )

    def format_for_header(self, series, dataframe):
        raise NotImplementedError()

    def format_for_column(self, column, col_number, dataframe):
        raise NotImplementedError()

    def format_for_cell(self, values, row_number, col_number, dataframe):
        raise NotImplementedError()

    def format_for_row(self, values, row_number, dataframe):
        raise NotImplementedError()

    def should_freeze_header(self, series, dataframe):
        raise NotImplementedError()

class BasicFormatter(DataFrameFormatter):

    @classmethod
    def with_defaults(cls,
        header_background_color=None, 
        header_text_color=None,
        date_format=None,
        decimal_format=None,
        integer_format=None,
        currency_format=None,
        freeze_headers=None,
        column_formats=None):
        return cls(
            (header_background_color or DEFAULT_HEADER_BACKGROUND_COLOR),
            header_text_color,
            date_format,
            decimal_format,
            integer_format,
            currency_format,
            freeze_headers,
            column_formats
        )

    def __init__(self, 
        header_background_color=None, 
        header_text_color=None,
        date_format=None,
        decimal_format=None,
        integer_format=None,
        currency_format=None,
        freeze_headers=None,
        column_formats=None):
        self.header_background_color = header_background_color
        self.header_text_color = header_text_color
        self.date_format = BasicFormatter.resolve_number_format(date_format or '', 'DATE')
        self.decimal_format = BasicFormatter.resolve_number_format(decimal_format or '', 'NUMBER')
        self.integer_format = BasicFormatter.resolve_number_format(integer_format or '', 'NUMBER')
        self.currency_format = BasicFormatter.resolve_number_format(currency_format or '', 'CURRENCY')
        self.freeze_headers = bool(freeze_headers)
        self.column_formats = column_formats or {}

    def format_for_header(self, series, dataframe):
        return cellFormat(
            backgroundColor=self.header_background_color, 
            textFormat=textFormat(bold=True, foregroundColor=self.header_text_color)
        )

    def format_for_column(self, column, col_number, dataframe):
        if column.name in self.column_formats:
            return self.column_formats[column.name]
        dtype = column.dtype
        if dtype.kind == 'O' and hasattr(column, 'infer_objects'):
            dtype = column.infer_objects().dtype
        if dtype.kind == 'f':
            return cellFormat(numberFormat=self.decimal_format, horizontalAlignment='RIGHT')
        elif dtype.kind == 'i':
            return cellFormat(numberFormat=self.integer_format, horizontalAlignment='RIGHT')
        elif dtype.kind == 'M':
            return cellFormat(numberFormat=self.date_format, horizontalAlignment='CENTER')
        else:
            return cellFormat(horizontalAlignment=('LEFT' if col_number == 1 else 'CENTER'))

    def format_for_cell(self, value, row_number, col_number, dataframe):
        return None

    def format_for_row(self, values, row_number, dataframe):
        return None

    def should_freeze_header(self, series, dataframe):
        return self.freeze_headers

DEFAULT_FORMATTER = BasicFormatter.with_defaults()
