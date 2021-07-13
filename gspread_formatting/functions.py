# -*- coding: utf-8 -*-

from .util import _build_repeat_cell_request, _fetch_with_updated_properties, _range_to_dimensionrange_object
from .models import CellFormat
from .conditionals import DataValidationRule
# These imports allow IDEs like PyCharm to verify the existence of these functions, 
# even though we will rebind the names below with wrapped versions of the functions
from gspread_formatting.batch_update_requests import * 
import gspread_formatting.batch_update_requests

from gspread.utils import a1_to_rowcol, rowcol_to_a1, finditem
from gspread.models import Spreadsheet
from gspread.urls import SPREADSHEET_URL

from functools import wraps

__all__ = (
    'get_default_format', 'get_effective_format', 'get_user_entered_format',
    'get_frozen_row_count', 'get_frozen_column_count', 
    'get_data_validation_rule',
) + gspread_formatting.batch_update_requests.__all__


def _wrap_as_standalone_function(func):
    @wraps(func)
    def f(worksheet, *args, **kwargs):
        return worksheet.spreadsheet.batch_update({'requests': func(worksheet, *args, **kwargs)})
    return f

for _fname in gspread_formatting.batch_update_requests.__all__:
    locals()[_fname] = _wrap_as_standalone_function(locals()[_fname])


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
    data = resp['sheets'][0]['data'][0]
    props = data.get('rowData', [{}])[0].get('values', [{}])[0].get('dataValidation')
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
    data = resp['sheets'][0]['data'][0]
    props = data.get('rowData', [{}])[0].get('values', [{}])[0].get('effectiveFormat')
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
    data = resp['sheets'][0]['data'][0]
    props = data.get('rowData', [{}])[0].get('values', [{}])[0].get('userEnteredFormat')
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


# monkey-patch Spreadsheet class

def fetch_sheet_metadata(self, params=None):
    if params is None:
        params = {'includeGridData': 'false'}
    url = SPREADSHEET_URL % self.id
    r = self.client.request('get', url, params=params)
    return r.json()

Spreadsheet.fetch_sheet_metadata = fetch_sheet_metadata
del fetch_sheet_metadata

