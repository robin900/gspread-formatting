# -*- coding: utf-8 -*-

from .util import _build_repeat_cell_request, _fetch_with_updated_properties
from .models import CellFormat
from .conditionals import DataValidationRule
from gspread.utils import a1_to_rowcol, rowcol_to_a1, finditem
from gspread.models import Spreadsheet
from gspread.urls import SPREADSHEET_URL

__all__ = (
    'format_cell_ranges', 'format_cell_range', 
    'get_default_format', 'get_effective_format', 'get_user_entered_format',
    'get_frozen_row_count', 'get_frozen_column_count', 'set_frozen',
    'set_data_validation_for_cell_range', 'set_data_validation_for_cell_ranges',
    'get_data_validation_rule'
)

def format_cell_ranges(worksheet, ranges):
    """Update a list of Cell object ranges of :class:`Cell` objects
    in the given ``Worksheet`` to have the accompanying ``CellFormat``.

    :param worksheet: The ``Worksheet`` object.
    :param ranges: An iterable whose elements are pairs of:
        a string with range value in A1 notation, e.g. 'A1:A5',
        and a ``CellFormat`` object).
    """

    body = {
        'requests': [
            _build_repeat_cell_request(worksheet, range, cell_format)
            for range, cell_format in ranges
        ]
    }

    return worksheet.spreadsheet.batch_update(body)

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
        and a ``DataValidationRule`` object).
    """

    body = {
        'requests': [
            _build_repeat_cell_request(worksheet, range, data_validation_rule, 'dataValidation')
            for range, data_validation_rule in ranges
        ]
    }

    return worksheet.spreadsheet.batch_update(body)

def set_data_validation_for_cell_range(worksheet, range, rule):
    """Update a list of Cell object ranges in the given ``Worksheet``
    to have the accompanying ``DataValidationRule``.
    :param worksheet: The ``Worksheet`` object.
    :param range: A string with range value in A1 notation, e.g. 'A1:A5'.
    :param rule: A DataValidationRule object.
    """

    return set_data_validation_for_cell_ranges(worksheet, [(range, rule)])

def get_data_validation_rule(worksheet, label):
    """Returns a DataValidationRule object or None representing the
    data validation in effect for the cell identified by ``label``.
    :param worksheet: Worksheet object containing the cell whose data
                      validation rule is desired.
    :param label: String with cell label in common format, e.g. 'B1'.
                  Letter case is ignored.
    Example:
    >>> get_data_validation_rule(worksheet, 'A1')
    <DataValidationRule condition=(bold=True)>
    >>> get_data_validation_rule(worksheet, 'A2')
    None
    """
    label = '%s!%s' % (worksheet.title, rowcol_to_a1(*a1_to_rowcol(label)))

    resp = worksheet.spreadsheet.fetch_sheet_metadata({
        'includeGridData': True,
        'ranges': [label],
        'fields': 'sheets.data.rowData.values.effectiveFormat,sheets.data.rowData.values.dataValidation'
    })
    props = resp['sheets'][0]['data'][0]['rowData'][0]['values'][0].get('dataValidation')
    return DataValidationRule.from_props(props) if props else None

def get_default_format(spreadsheet):
    """Return Default CellFormat for spreadsheet, or None if no default formatting was specified."""
    fmt = _fetch_with_updated_properties(spreadsheet, 'defaultFormat')
    return CellFormat.from_props(fmt) if fmt else None

def get_effective_format(worksheet, label):
    """Returns a CellFormat object or None representing the effective formatting directives,
    if any, for the cell; that is a combination of default formatting, user-entered formatting,
    and conditional formatting.

    :param worksheet: Worksheet object containing the cell whose format is desired.
    :param label: String with cell label in common format, e.g. 'B1'.
                  Letter case is ignored.
    Example:

    >>> get_effective_format(worksheet, 'A1')
    <CellFormat textFormat=(bold=True)>
    >>> get_effective_format(worksheet, 'A2')
    None
    """
    label = '%s!%s' % (worksheet.title, rowcol_to_a1(*a1_to_rowcol(label)))

    resp = worksheet.spreadsheet.fetch_sheet_metadata({
        'includeGridData': True,
        'ranges': [label],
        'fields': 'sheets.data.rowData.values.effectiveFormat'
    })
    props = resp['sheets'][0]['data'][0]['rowData'][0]['values'][0].get('effectiveFormat')
    return CellFormat.from_props(props) if props else None

def get_user_entered_format(worksheet, label):
    """Returns a CellFormat object or None representing the user-entered formatting directives,
    if any, for the cell.

    :param worksheet: Worksheet object containing the cell whose format is desired.
    :param label: String with cell label in common format, e.g. 'B1'.
                  Letter case is ignored.
    Example:

    >>> get_user_entered_format(worksheet, 'A1')
    <CellFormat textFormat=(bold=True)>
    >>> get_user_entered_format(worksheet, 'A2')
    None
    """
    label = '%s!%s' % (worksheet.title, rowcol_to_a1(*a1_to_rowcol(label)))

    resp = worksheet.spreadsheet.fetch_sheet_metadata({
        'includeGridData': True,
        'ranges': [label],
        'fields': 'sheets.data.rowData.values.userEnteredFormat'
    })
    props = resp['sheets'][0]['data'][0]['rowData'][0]['values'][0].get('userEnteredFormat')
    return CellFormat.from_props(props) if props else None

def get_frozen_row_count(worksheet):
    md = worksheet.spreadsheet.fetch_sheet_metadata({'includeGridData': True})
    sheet_data = finditem(lambda i: i['properties']['title'] == worksheet.title, md['sheets'])
    grid_props = sheet_data['properties']['gridProperties']
    return grid_props.get('frozenRowCount')

def get_frozen_column_count(worksheet):
    md = worksheet.spreadsheet.fetch_sheet_metadata({'includeGridData': True})
    sheet_data = finditem(lambda i: i['properties']['title'] == worksheet.title, md['sheets'])
    grid_props = sheet_data['properties']['gridProperties']
    return grid_props.get('frozenColumnCount')

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
    body = {
        'requests': [{
            'updateSheetProperties': {
                'properties': {
                    'sheetId': worksheet.id,
                    'gridProperties': grid_properties
                },
                'fields': fields
            }
        }]
    }
    return worksheet.spreadsheet.batch_update(body)

# monkey-patch Spreadsheet class

def fetch_sheet_metadata(self, params=None):
    if params is None:
        params = {'includeGridData': 'false'}
    url = SPREADSHEET_URL % self.id
    r = self.client.request('get', url, params=params)
    return r.json()

Spreadsheet.fetch_sheet_metadata = fetch_sheet_metadata
del fetch_sheet_metadata

