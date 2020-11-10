# -*- coding: utf-8 -*-
"""
This module provides functions that generate request objects compatible with the
"batchUpdate" API call in Google Sheets. Both the ``.functions`` and
``.batch`` modules make use of these request functions, wrapping them
in functions that make the API call or calls using the generated request objects.
"""

from .util import _build_repeat_cell_request, _range_to_dimensionrange_object

from functools import wraps

__all__ = (
    'format_cell_ranges', 'format_cell_range', 'set_frozen',
    'set_data_validation_for_cell_range', 'set_data_validation_for_cell_ranges',
    'set_row_height', 'set_row_heights',
    'set_column_width', 'set_column_widths'
)


def set_row_heights(worksheet, ranges):
    """Update a row or range of rows in the given ``Worksheet`` 
    to have the specified height in pixels.

    :param worksheet: The ``Worksheet`` object.
    :param ranges: An iterable whose elements are pairs of:
        a string with row range value in A1 notation, e.g. '1' or '1:50',
        and a integer specifying height in pixels.
    """

    return [
        { 
            'updateDimensionProperties': { 
                'range': _range_to_dimensionrange_object(range, worksheet.id), 
                'properties': { 'pixelSize': height }, 
                'fields': 'pixelSize' 
            } 
        }
        for range, height in ranges
    ]


def set_row_height(worksheet, label, height):
    """Update a row or range of rows in the given ``Worksheet`` 
    to have the specified height in pixels.

    :param worksheet: The ``Worksheet`` object.
    :param label: string representing a single row or range of rows, e.g. ``1`` or ``3:400``.
    :param height: An integer greater than or equal to 0.

    """
    return set_row_heights(worksheet, [(label, height)])
 

def set_column_widths(worksheet, ranges):
    """Update a column or range of columns in the given ``Worksheet`` 
    to have the specified width in pixels.

    :param worksheet: The ``Worksheet`` object.
    :param ranges: An iterable whose elements are pairs of:
                   a string with column range value in A1 notation, e.g. 'A:C',
                   and a integer specifying width in pixels.

    """

    return [
        { 
            'updateDimensionProperties': { 
                'range': _range_to_dimensionrange_object(range, worksheet.id), 
                'properties': { 'pixelSize': width }, 
                'fields': 'pixelSize' 
            } 
        }
        for range, width in ranges
    ]


def set_column_width(worksheet, label, width):
    """Update a column or range of columns in the given ``Worksheet`` 
    to have the specified width in pixels.

    :param worksheet: The ``Worksheet`` object.
    :param label: string representing a single column or range of columns, e.g. ``A`` or ``B:D``.
    :param height: An integer greater than or equal to 0.

    """

    return set_column_widths(worksheet, [(label, width)])

      
def format_cell_ranges(worksheet, ranges):
    """Update a list of Cell object ranges of :class:`Cell` objects
    in the given ``Worksheet`` to have the accompanying ``CellFormat``.

    :param worksheet: The ``Worksheet`` object.
    :param ranges: An iterable whose elements are pairs of:
        a string with range value in A1 notation, e.g. 'A1:A5',
        and a ``CellFormat`` object).

    """

    return [
        _build_repeat_cell_request(worksheet, range, cell_format)
        for range, cell_format in ranges
    ]


def format_cell_range(worksheet, name, cell_format):
    """Update a range of :class:`Cell` objects in the given Worksheet
    to have the specified ``CellFormat``.

    :param worksheet: The ``Worksheet`` object.
    :param name: A string with range value in A1 notation, e.g. 'A1:A5'.
    :param cell_format: A ``CellFormat`` object.

    """

    return format_cell_ranges(worksheet, [(name, cell_format)])


def set_data_validation_for_cell_ranges(worksheet, ranges):
    """Update a list of Cell object ranges of :class:`Cell` objects
    in the given ``Worksheet`` to have the accompanying ``DataValidationRule``.

    :param worksheet: The ``Worksheet`` object.
    :param ranges: An iterable whose elements are pairs of:
                   a string with range value in A1 notation, e.g. 'A1:A5',
                   and a ``DataValidationRule`` object or None to clear the data
                   validation rule).

    """

    return [
        _build_repeat_cell_request(worksheet, range, data_validation_rule, 'dataValidation')
        for range, data_validation_rule in ranges
    ]


def set_data_validation_for_cell_range(worksheet, range, rule):
    """Update a Cell range in the given ``Worksheet``
    to have the accompanying ``DataValidationRule`` (or no rule).

    :param worksheet: The ``Worksheet`` object.
    :param range: A string with range value in A1 notation, e.g. 'A1:A5'.
    :param rule: A DataValidationRule object, or None to remove data validation rule for cells..

    """

    return set_data_validation_for_cell_ranges(worksheet, [(range, rule)])


def set_frozen(worksheet, rows=None, cols=None):
    if rows is None and cols is None:
        raise ValueError("Must specify at least one of rows and cols")
    grid_properties = {}
    if rows is not None:
        grid_properties['frozenRowCount'] = rows
    if cols is not None:
        grid_properties['frozenColumnCount'] = cols
    fields = ','.join(
        'gridProperties.%s' % p for p in grid_properties.keys()
    )
    return [{
        'updateSheetProperties': {
            'properties': {
                'sheetId': worksheet.id,
                'gridProperties': grid_properties
            },
            'fields': fields
        }
    }]

