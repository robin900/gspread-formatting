# -*- coding: utf-8 -*-

try:
    from itertools import zip_longest
except ImportError:
    from itertools import izip_longest as zip_longest

from gspread_formatting.batch_update_requests import format_cell_ranges, set_frozen
from gspread_formatting.models import cellFormat, numberFormat, Color, textFormat

from gspread.utils import rowcol_to_a1

from functools import wraps

__all__ = (
    'format_with_dataframe', 
    'DataFrameFormatter', 
    'BasicFormatter', 
    'DEFAULT_FORMATTER', 
    'DEFAULT_HEADER_BACKGROUND_COLOR'
)

DEFAULT_HEADER_BACKGROUND_COLOR = Color(0.8980392, 0.8980392, 0.8980392)

def _determine_index_or_columns_size(obj):
    if hasattr(obj, 'levshape'):
        return len(obj.levshape)
    return 1
        
def _format_with_dataframe(worksheet,
                          dataframe,
                          formatter=None,
                          row=1,
                          col=1,
                          include_index=False,
                          include_column_header=True):
    """
    Modifies the cell formatting of an area of the provided Worksheet, using
    the provided DataFrame to determine the area to be formatted and the formats
    to be used.

    :param worksheet: the gspread worksheet to set with content of DataFrame.
    :param dataframe: the DataFrame.
    :param formatter: an optional instance of ``DataFrameFormatter`` class, which
                      will examine the contents of the DataFrame and
                      assemble a set of ``gspread_formatter`` operations
                      to be performed after the DataFrame contents 
                      are written to the given Worksheet. The formatting
                      operations are performed after the contents are written
                      and before this function returns. Defaults to 
                      ``DEFAULT_FORMATTER``.
    :param row: number of row at which to begin formatting. Defaults to 1.
    :param col: number of column at which to begin formatting. Defaults to 1.
    :param include_index: if True, include the DataFrame's index as an
            additional column when performing formatting. Defaults to False.
    :param include_column_header: if True, format a header row before data.
            Defaults to True.
    """
    if not formatter:
        formatter = DEFAULT_FORMATTER

    formatting_ranges = []

    columns = [ dataframe[c] for c in dataframe.columns ]
    index_column_size = _determine_index_or_columns_size(dataframe.index)
    column_header_size = _determine_index_or_columns_size(dataframe.columns)

    if include_index:
        # allow for multi-index index
        if index_column_size > 1:
            reset_df = dataframe.reset_index()
            index_elts = [ reset_df[c] for c in list(reset_df.columns)[:index_column_size] ]
        else:
            index_elts = [ dataframe.index ]
        columns = index_elts + columns

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
        # TODO allow for multi-index columns object
        elts = list(dataframe.columns)
        if include_index:
            # allow for multi-index index
            if index_column_size > 1:
                index_names = list(dataframe.index.names)
            else:
                index_names = [ dataframe.index.name ]
            elts = index_names + elts
            header_fmt = formatter.format_for_header(dataframe.index, dataframe)
            if header_fmt:
                formatting_ranges.append( 
                    (
                        '{}:{}'.format(
                            rowcol_to_a1(row, col), 
                            rowcol_to_a1(row + dataframe.shape[0], col + index_column_size - 1)
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
                        rowcol_to_a1(row + column_header_size - 1, col + len(elts) - 1)
                    ), 
                    header_fmt
                )
            )

        freeze_args = {}

        if row == 1 and formatter.should_freeze_header(elts, dataframe):
            freeze_args['rows'] = column_header_size

        if include_index and col == 1 and formatter.should_freeze_header(dataframe.index, dataframe):
            freeze_args['cols'] = index_column_size

        row += column_header_size


    values = []
    for value_row, index_value in zip_longest(dataframe.values, dataframe.index):
        if include_index:
            if index_column_size > 1:
                index_values = list(index_value)
            else:
                index_values = [index_value]
            value_row = index_values + list(value_row)
        values.append(value_row)
    for y_idx, value_row in enumerate(values):
        for x_idx, cell_value in enumerate(value_row):
            cell_fmt = formatter.format_for_cell(cell_value, y_idx+row, x_idx+col, dataframe)
            if cell_fmt:
                formatting_ranges.append((rowcol_to_a1(y_idx+row, x_idx+col), cell_fmt))
        row_fmt = formatter.format_for_data_row(values, y_idx+row, dataframe)
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

    requests = []

    if formatting_ranges:
        formatting_ranges = [ r for r in formatting_ranges if r[1] and r[1].to_props() ]
        requests.extend(format_cell_ranges(worksheet, formatting_ranges))

    if freeze_args:
        requests.extend(set_frozen(worksheet, **freeze_args))

    return requests

@wraps(_format_with_dataframe)
def format_with_dataframe(worksheet, *args, **kwargs):
    return worksheet.spreadsheet.batch_update(
        {'requests': _format_with_dataframe(worksheet, *args, **kwargs)}
    )

class DataFrameFormatter(object):
    """
    An abstract base class defining the interface for producing formats
    for a worksheet based on a given DataFrame.
    """
    @classmethod
    def resolve_number_format(cls, value, type=None):
        """
        A utility class method that resolves a value to a ``NumberFormat`` object,
        whether that value is a ``NumberFormat`` object or a pattern string.
        Optional ``type`` parameter is to specify ``NumberFormat`` enum value.
        """
        if value is None:
            return None
        elif isinstance(value, numberFormat):
            return value
        elif isinstance(value, str):
            return numberFormat(type, value) 
        else:
            raise ValueError(value)

    def format_with_dataframe(self, worksheet, dataframe, row=1, col=1, include_index=False, include_column_header=True):
        """
        Convenience method that will call this module's ``format_with_dataframe`` function with
        this ``DataFrameFormatter`` object as the formatter.
        """
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
        """
        Called by ``format_with_dataframe`` once for each header row (if ``include_column_header``
        parameter is ``True``) or column (if ``include_index`` parameter is also ``True``)..

        :param series: A sequence of elements representing the values in the row or column.
                 Can be a simple list, or a ``pandas.Series`` or ``pandas.Index`` object.
        :param dataframe: The ``pandas.DataFrame`` object, as additional context.

        :return: Either a ``CellFormat`` object or ``None``.
        """
        raise NotImplementedError()

    def format_for_column(self, column, col_number, dataframe):
        """
        Called by ``format_with_dataframe`` once for each column in the dataframe.

        :param column: A ``pandas.Series`` object representing the column.
        :param col_number: The index (starting with 1) of the column in the worksheet.
        :param dataframe: The ``pandas.DataFrame`` object, as additional context.

        :return: Either a ``CellFormat`` object or ``None``.
        """
        raise NotImplementedError()

    def format_for_data_row(self, values, row_number, dataframe):
        """
        Called by ``format_with_dataframe`` once for each data row in the dataframe.
        Allows for row-specific additional formatting to complement any
        column-based formatting.

        :param values: The values in the row, obtained directly from the ``DataFrame``.
                 If ``include_index`` parameter to ``format_with_dataframe`` is ``True``,
                 then the first element in this sequence is the index value for the row.
        :param row_number: The index (starting with 1) of the row in the worksheet.
        :param dataframe: The ``pandas.DataFrame`` object, as additional context.

        :return: Either a ``CellFormat`` object or ``None``.
        """
        raise NotImplementedError()

    def format_for_cell(self, value, row_number, col_number, dataframe):
        """
        Called by ``format_with_dataframe`` once for each cell in the dataframe.
        Allows for cell-specific additional formatting to complement any column
        or row formatting.

        :param value: The value of the cell, obtained directly from the ``DataFrame``.
        :param row_number: The index (starting with 1) of the row in the worksheet.
        :param col_number: The index (starting with 1) of the column in the worksheet.
        :param dataframe: The ``pandas.DataFrame`` object, as additional context.

        :return: Either a ``CellFormat`` object or ``None``.
        """
        raise NotImplementedError()

    def should_freeze_header(self, series, dataframe):
        """
        Called by ``format_with_dataframe`` once for each header row or column.

        :param series: A sequence of elements representing the values in the row or column.
                 Can be a simple list, or a ``pandas.Series`` or ``pandas.Index`` object.
        :param dataframe: The ``pandas.DataFrame`` object, as additional context.

        :return: boolean value
        """
        raise NotImplementedError()

class BasicFormatter(DataFrameFormatter):
    """
    A basic formatter class that offers: selection of format based on
    inferred data type of each column; bold headers with a custom background color;
    frozen header row (and column if index is included); and column-specific
    override formats.
    """

    @classmethod
    def with_defaults(cls,
        header_background_color=None, 
        header_text_color=None,
        date_format=None,
        decimal_format=None,
        integer_format=None,
        freeze_headers=None,
        column_formats=None):
        """
        Returns an instance of this class, with any unspecified parameters
        being substituted with this package's default values for the parameters.
        Instantiate the class directly if you want unspecified parameters to be ``None``
        and thus always be omitted from formatting operations.
        """
        return cls(
            (header_background_color or DEFAULT_HEADER_BACKGROUND_COLOR),
            header_text_color,
            date_format,
            decimal_format,
            integer_format,
            freeze_headers,
            column_formats
        )

    def __init__(self, 
        header_background_color=None, 
        header_text_color=None,
        date_format=None,
        decimal_format=None,
        integer_format=None,
        freeze_headers=None,
        column_formats=None):
        self.header_background_color = header_background_color
        self.header_text_color = header_text_color
        self.date_format = BasicFormatter.resolve_number_format(date_format or '', 'DATE')
        self.decimal_format = BasicFormatter.resolve_number_format(decimal_format or '', 'NUMBER')
        self.integer_format = BasicFormatter.resolve_number_format(integer_format or '', 'NUMBER')
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

    def format_for_data_row(self, values, row_number, dataframe):
        return None

    def should_freeze_header(self, series, dataframe):
        return self.freeze_headers

DEFAULT_FORMATTER = BasicFormatter.with_defaults()
