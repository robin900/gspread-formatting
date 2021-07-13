# -*- coding: utf-8 -*-

import os
import re
import random
import unittest
import itertools
import uuid
from datetime import datetime, date
import pandas as pd
from gspread_dataframe import set_with_dataframe

try:
    from StringIO import StringIO
except ImportError:
    from io import StringIO

try:
    import ConfigParser
except ImportError:
    import configparser as ConfigParser

from oauth2client.service_account import ServiceAccountCredentials

import gspread
from gspread import utils
from gspread_formatting import *
from gspread_formatting.dataframe import *
from gspread_formatting.util import _range_to_gridrange_object, _range_to_dimensionrange_object

try:
    unicode
except NameError:
    basestring = unicode = str


CONFIG_FILENAME = os.path.join(os.path.dirname(__file__), 'tests.config')
CREDS_FILENAME = os.path.join(os.path.dirname(__file__), 'creds.json')
SCOPE = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive.file'
]

I18N_STR = u'Iñtërnâtiônàlizætiøn'  # .encode('utf8')


def read_config():
    config = ConfigParser.ConfigParser()
    envconfig = os.environ.get('GSPREAD_FORMATTING_CONFIG')
    if envconfig:
        fp = StringIO(envconfig)
    else:
        fp = open(CONFIG_FILENAME)
    if hasattr(config, 'read_file'):
       read_func = config.read_file
    else:
       read_func = config.readfp
    try:
        read_func(fp)
    finally:
        fp.close()
    return config

def read_credentials():
    credjson = os.environ.get('GSPREAD_FORMATTING_CREDENTIALS')
    if credjson:
        return ServiceAccountCredentials.from_json_keyfile_dict(json.loads(credjson), SCOPE)
    else:
        return ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILENAME, SCOPE)


def gen_value(prefix=None):
    if prefix:
        return u'%s %s' % (prefix, gen_value())
    else:
        return unicode(uuid.uuid4())


class RangeConversionTest(unittest.TestCase):
    RANGES = {
        'A': {'startColumnIndex': 0, 'endColumnIndex': 1},
        'A:C': {'startColumnIndex': 0, 'endColumnIndex': 3},
        'A5:B': {'startRowIndex': 4, 'startColumnIndex': 0, 'endColumnIndex': 2},
        '3': {'startRowIndex': 2, 'endRowIndex': 3},
        '3:100': {'startRowIndex': 2, 'endRowIndex': 100}
    }

    ILLEGAL_RANGES = (
        'B:A',
        'A100:A1',
        'C1:A20',
        'AA1:A1',
        ''
    )

    DIMENSION_RANGES = {
        'A': {'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 1},
        'A:C': {'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 3},
        '3': {'dimension': 'ROWS', 'startIndex': 2, 'endIndex': 3},
        '3:100': {'dimension': 'ROWS', 'startIndex': 2, 'endIndex': 100}
    }

    ILLEGAL_DIMENSION_RANGES = ( 'A5:B', '1:C3', 'A1:D5' )

    def test_ranges(self):
        worksheet_id = 0
        for range, gridrange_obj in self.RANGES.items():
            gridrange_obj['sheetId'] = worksheet_id
            self.assertEqual(gridrange_obj, _range_to_gridrange_object(range, worksheet_id))
        pass

    def test_illegal_ranges(self):
        for range in self.ILLEGAL_RANGES:
            exc = None
            try:
                _range_to_gridrange_object(range, 0)
            except Exception as e:
                exc = e
            self.assertTrue(isinstance(exc, ValueError))

    def test_dimension_ranges(self):
        worksheet_id = 0
        for range, range_obj in self.DIMENSION_RANGES.items():
            range_obj['sheetId'] = worksheet_id
            self.assertEqual(range_obj, _range_to_dimensionrange_object(range, worksheet_id))
        pass

    def test_illegal_dimension_ranges(self):
        for range in self.ILLEGAL_DIMENSION_RANGES:
            exc = None
            try:
                _range_to_dimensionrange_object(range, 0)
            except Exception as e:
                exc = e
            self.assertTrue(isinstance(exc, ValueError))

class GspreadTest(unittest.TestCase):
    maxDiff = None
    config = None
    gc = None

    @classmethod
    def setUpClass(cls):
        try:
            cls.config = read_config()
            credentials = read_credentials()
            cls.gc = gspread.authorize(credentials)
        except IOError as e:
            msg = "Can't find %s for reading test configuration. "
            raise Exception(msg % e.filename)

    def setUp(self):
        if self.__class__.gc is None:
            self.__class__.setUpClass()
        self.assertTrue(isinstance(self.gc, gspread.client.Client))

class WorksheetTest(GspreadTest):
    """Test for gspread.Worksheet."""
    spreadsheet = None

    @classmethod
    def setUpClass(cls):
        super(WorksheetTest, cls).setUpClass()
        ss_id = cls.config.get('Spreadsheet', 'id')
        cls.spreadsheet = cls.gc.open_by_key(ss_id)
        cls.spreadsheet.batch_update(
            {
                "requests": [
                    {
                        "updateSpreadsheetProperties": {
                            "properties": {"locale": "en_US"},
                            "fields": "locale",
                        }
                    }
                ]
            }
        )
        try:
            test_sheet = cls.spreadsheet.worksheet('wksht_test')
            if test_sheet:
                # somehow left over from interrupted test, remove.
                cls.spreadsheet.del_worksheet(test_sheet)
        except gspread.exceptions.WorksheetNotFound:
            pass # expected

    def setUp(self):
        super(WorksheetTest, self).setUp()
        if self.__class__.spreadsheet is None:
            self.__class__.setUpClass()
        try:
            test_sheet = self.spreadsheet.worksheet('wksht_test')
            if test_sheet:
                # somehow left over from interrupted test, remove.
                self.spreadsheet.del_worksheet(test_sheet)
        except gspread.exceptions.WorksheetNotFound:
            pass # expected
        self.sheet = self.spreadsheet.add_worksheet('wksht_test', 20, 20)

    def tearDown(self):
        self.spreadsheet.del_worksheet(self.sheet)

    def test_some_format_constructors(self):
        f = numberFormat('TEXT', '###0')
        f = border('DOTTED', color(0.2, 0.2, 0.2))

    def test_bottom_attribute(self):
        f = padding(bottom=1.1)
        f = borders(bottom=border('SOLID'))

    def test_format_range(self):
        rows = [["", "", "", ""],
                ["", "", "", ""],
                ["A1", "B1", "", "D1"],
                [1, "b2", 1.45, ""],
                ["", "", "", ""],
                ["A4", 0.4, "", 4]]

        def_fmt = get_default_format(self.spreadsheet)
        cell_list = self.sheet.range('A1:D6')
        for cell, value in zip(cell_list, itertools.chain(*rows)):
            cell.value = value
        self.sheet.update_cells(cell_list)

        fmt = cellFormat(textFormat=textFormat(bold=True), backgroundColorStyle=ColorStyle(rgbColor=Color(1,0,0)))
        format_cell_ranges(self.sheet, [('A:A', fmt), ('B1:B6', fmt), ('C1:D6', fmt), ('2', fmt)])
        ue_fmt = get_user_entered_format(self.sheet, 'A1')
        self.assertEqual(ue_fmt.textFormat.bold, True)
        # userEnteredFormat will not have backgroundColorStyle...
        eff_fmt = get_effective_format(self.sheet, 'A1')
        self.assertEqual(eff_fmt.textFormat.bold, True)
        # effectiveFormat will have backgroundColorStyle...
        self.assertEqual(eff_fmt.backgroundColorStyle.rgbColor.red, 1)
        self.assertEqual(eff_fmt.textFormat.bold, True)
        fmt2 = cellFormat(textFormat=textFormat(italic=True))
        format_cell_range(self.sheet, 'A:D', fmt2)
        ue_fmt = get_user_entered_format(self.sheet, 'A1')
        self.assertEqual(ue_fmt.textFormat.italic, True)
        eff_fmt = get_effective_format(self.sheet, 'A1')
        self.assertEqual(eff_fmt.textFormat.italic, True)

    def test_bottom_formatting(self):
        rows = [["", "", "", ""],
                ["", "", "", ""],
                ["A1", "B1", "", "D1"],
                [1, "b2", 1.45, ""],
                ["", "", "", ""],
                ["A4", 0.4, "", 4]]

        def_fmt = get_default_format(self.spreadsheet)
        cell_list = self.sheet.range('A1:D6')
        for cell, value in zip(cell_list, itertools.chain(*rows)):
            cell.value = value
        self.sheet.update_cells(cell_list)
        fmt = cellFormat(textFormat=textFormat(bold=True))
        format_cell_ranges(self.sheet, [('A1:B6', fmt), ('C1:D6', fmt)])

        orig_fmt = get_user_entered_format(self.sheet, 'A1')
        new_fmt = cellFormat(borders=borders(bottom=border('SOLID')), padding=padding(bottom=3))
        format_cell_range(self.sheet, 'A1:A1', new_fmt)
        # Sheets API bug: user entered format will now contain default color and colorStyle.rgbColor
        ue_fmt = get_user_entered_format(self.sheet, 'A1')
        self.assertEqual(new_fmt.borders.bottom.style, ue_fmt.borders.bottom.style)
        self.assertEqual(new_fmt.padding.bottom, ue_fmt.padding.bottom)
        eff_fmt = get_effective_format(self.sheet, 'A1')
        self.assertEqual(new_fmt.borders.bottom.style, eff_fmt.borders.bottom.style)
        self.assertEqual(new_fmt.padding.bottom, eff_fmt.padding.bottom)

    def test_frozen_rows_cols(self):
        set_frozen(self.sheet, rows=1, cols=1)
        fresh = self.sheet.spreadsheet.fetch_sheet_metadata({'includeGridData': True})
        item = utils.finditem(lambda x: x['properties']['title'] == self.sheet.title, fresh['sheets'])
        pr = item['properties']['gridProperties']
        self.assertEqual(pr.get('frozenRowCount'), 1)
        self.assertEqual(pr.get('frozenColumnCount'), 1)
        self.assertEqual(get_frozen_row_count(self.sheet), 1)
        self.assertEqual(get_frozen_column_count(self.sheet), 1)

    def test_format_props_roundtrip(self):
        fmt = cellFormat(backgroundColor=Color(1,0,1),textFormat=textFormat(italic=False))
        fmt_roundtrip = CellFormat.from_props(fmt.to_props())
        self.assertEqual(fmt, fmt_roundtrip)

    def test_formats_equality_and_arithmetic(self):
        def_fmt = cellFormat(backgroundColor=Color(1,0,1),textFormat=textFormat(italic=False))
        fmt = cellFormat(textFormat=textFormat(bold=True))
        effective_format = def_fmt + fmt
        self.assertEqual(effective_format.textFormat.bold, True)
        effective_format2 = def_fmt + fmt
        self.assertEqual(effective_format, effective_format2)
        self.assertEqual(effective_format - fmt, def_fmt)
        self.assertEqual(effective_format.difference(fmt), def_fmt)
        self.assertEqual(effective_format.intersection(effective_format), effective_format)
        self.assertEqual(effective_format & effective_format, effective_format)
        self.assertEqual(effective_format - effective_format, None)

    def test_date_formatting_roundtrip(self):
        rows = [
            ["9/1/2018", "1/2/2017", "4/4/2014", "4/4/2019"],
            ["10/2/2019", "2/4/2000", "5/5/1994", "7/7/1979"]
        ]
        def_fmt = get_default_format(self.spreadsheet)
        cell_list = self.sheet.range('A1:D2')
        for cell, value in zip(cell_list, itertools.chain(*rows)):
            cell.value = value
        self.sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
        fmt = cellFormat(
                numberFormat=numberFormat('DATE', ' DD MM YYYY'),
                backgroundColor=color(0.8, 0.9, 1),
                horizontalAlignment='RIGHT',
                textFormat=textFormat(bold=False))
        format_cell_range(self.sheet, 'A1:D2', fmt)
        ue_fmt = get_user_entered_format(self.sheet, 'A1')
        self.assertEqual(ue_fmt.numberFormat.type, 'DATE')
        self.assertEqual(ue_fmt.numberFormat.pattern, ' DD MM YYYY')
        eff_fmt = get_effective_format(self.sheet, 'A1')
        self.assertEqual(eff_fmt.numberFormat.type, 'DATE')
        self.assertEqual(eff_fmt.numberFormat.pattern, ' DD MM YYYY')
        dt = self.sheet.acell('A1').value
        self.assertEqual(dt, ' 01 09 2018')

    def test_blank_color_as_black(self):
        rows = [
            ["A", "B", "C", "D"],
            ["1", "2", "3", "4"],
            ["A", "B", "C", "D"],
            ["TRUE", "FALSE", "FALSE", "TRUE"],
        ]
        cell_list = self.sheet.range('A1:D4')
        for cell, value in zip(cell_list, itertools.chain(*rows)):
            cell.value = value
        self.sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
        fmt = CellFormat(backgroundColor=Color(0,0,0,1))
        format_cell_range(self.sheet, '1:1', fmt)
        ue_fmt = get_user_entered_format(self.sheet, 'A1')
        self.assertEqual(ue_fmt.backgroundColor, Color(0,0,0,1))
        self.assertEqual(ue_fmt.backgroundColor, Color())
        fmt = CellFormat(backgroundColor=Color(red=1))
        format_cell_range(self.sheet, '1:1', fmt)
        ue_fmt = get_user_entered_format(self.sheet, 'A1')
        self.assertEqual(ue_fmt.backgroundColor, Color(1,0,0,1))
        self.assertEqual(ue_fmt.backgroundColor, Color(red=1))
        fmt = CellFormat(backgroundColor=Color())
        format_cell_range(self.sheet, '1:1', fmt)
        ue_fmt = get_user_entered_format(self.sheet, 'A1')
        eff_fmt = get_effective_format(self.sheet, 'A1')
        self.assertEqual(ue_fmt.backgroundColor, Color(0,0,0,1))
        self.assertEqual(ue_fmt.backgroundColor, Color())


    def test_empty_cell_formatting(self):
        self.assertEqual(get_user_entered_format(self.sheet, 'A1'), None)
        self.assertEqual(get_effective_format(self.sheet, 'A1'), None)
        self.assertEqual(get_data_validation_rule(self.sheet, 'A1'), None)
        

    def test_data_validation_rule(self):
        rows = [
            ["A", "B", "C", "D"],
            ["1", "2", "3", "4"],
            ["A", "B", "C", "D"],
            ["TRUE", "FALSE", "FALSE", "TRUE"],
        ]
        cell_list = self.sheet.range('A1:D4')
        for cell, value in zip(cell_list, itertools.chain(*rows)):
            cell.value = value
        self.sheet.update_cells(cell_list, value_input_option='USER_ENTERED')
        validation_rule = DataValidationRule(
            BooleanCondition('ONE_OF_LIST', ['1', '2', '3', '4']), 
            showCustomUi=True
        )
        set_data_validation_for_cell_range(self.sheet, 'A2:D2', validation_rule)
        # No data validation for A1
        eff_rule = get_data_validation_rule(self.sheet, 'A1')
        self.assertEqual(eff_rule, None)
        # data validation for A2 should be equal to validation_rule
        eff_rule = get_data_validation_rule(self.sheet, 'A2')
        self.assertEqual(eff_rule.condition.type, 'ONE_OF_LIST')
        self.assertEqual([ x.userEnteredValue for x in eff_rule.condition.values ], ['1', '2', '3', '4'])
        self.assertEqual(eff_rule.showCustomUi, True)
        self.assertEqual(eff_rule.strict, None)
        self.assertEqual(eff_rule, validation_rule)

        boolean_validation_rule = DataValidationRule(
            BooleanCondition('BOOLEAN', [])
        )
        set_data_validation_for_cell_range(self.sheet, 'A4:D4', boolean_validation_rule)
        eff_rule = get_data_validation_rule(self.sheet, 'A4')
        self.assertEqual([ x.userEnteredValue for x in eff_rule.condition.values ], [])
        self.assertEqual(eff_rule.showCustomUi, None)
        self.assertEqual(eff_rule.strict, None)
        self.assertEqual(eff_rule, boolean_validation_rule)

        set_data_validation_for_cell_range(self.sheet, 'A4:D4', None)
        eff_rule = get_data_validation_rule(self.sheet, 'A4')
        self.assertEqual(eff_rule, None)

    def test_boolean_condition(self):
        with self.assertRaises(ValueError):
            BooleanCondition('TEXT_EQ', 'foo')
        with self.assertRaises(ValueError):
            BooleanCondition('ONE_OF_LIST', 'foo')

    def test_conditional_format_rules(self):
        rows = [
            ["A", "B", "C", "D"],
            ["1", "2", "3", "4"]
        ]
        cell_list = self.sheet.range('A1:D2')
        for cell, value in zip(cell_list, itertools.chain(*rows)):
            cell.value = value
        self.sheet.update_cells(cell_list, value_input_option='USER_ENTERED')

        current_rules = get_conditional_format_rules(self.sheet)
        self.assertEqual(list(current_rules), [])
        new_rule = ConditionalFormatRule(
            ranges=[GridRange.from_a1_range('A1:D1', self.sheet)],
            booleanRule=BooleanRule(
                condition=BooleanCondition('TEXT_CONTAINS', ['A']), 
                format=CellFormat(textFormat=TextFormat(bold=True))
            )
        )
        new_rule_2 = ConditionalFormatRule(
            ranges=[GridRange.from_a1_range('A2:D2', self.sheet)],
            gradientRule=GradientRule(
                maxpoint=InterpolationPoint(colorStyle=ColorStyle(themeColor='BACKGROUND'), type='MAX'),
                minpoint=InterpolationPoint(colorStyle=ColorStyle(themeColor='TEXT'), type='NUMBER', value='1')
            )
        )
        current_rules.append(new_rule)
        current_rules.append(new_rule_2)
        self.assertNotEqual(current_rules.save(), None)
        # re-saving _always_ sends a request to API, even if no local changes made
        self.assertNotEqual(current_rules.save(), None)
        current_rules = get_conditional_format_rules(self.sheet)
        self.assertEqual(
            current_rules.rules[0].booleanRule.format.textFormat.bold, 
            new_rule.booleanRule.format.textFormat.bold
        )
        self.assertEqual(
            current_rules.rules[1].gradientRule.maxpoint.colorStyle.themeColor, 
            new_rule_2.gradientRule.maxpoint.colorStyle.themeColor, 
        )
        current_rules[0] = new_rule_2
        del current_rules[1]
        current_rules.append(new_rule)
        self.assertNotEqual(current_rules.save(), None)
        current_rules = get_conditional_format_rules(self.sheet)
        self.assertEqual(
            current_rules.rules[1].booleanRule.format.textFormat.bold, 
            new_rule.booleanRule.format.textFormat.bold
        )
        self.assertEqual(
            current_rules.rules[0].gradientRule.maxpoint.colorStyle.themeColor, 
            new_rule_2.gradientRule.maxpoint.colorStyle.themeColor, 
        )

        bold_fmt = get_effective_format(self.sheet, 'A1')
        normal_fmt = get_effective_format(self.sheet, 'C1')
        self.assertEqual(bold_fmt.textFormat.bold, True)
        self.assertEqual(bool(normal_fmt.textFormat.bold), False)
        self.assertEqual(bool(normal_fmt.textFormat.italic), False)

        current_rules.clear()
        current_rules.save()
        current_rules = get_conditional_format_rules(self.sheet)
        self.assertEqual(list(current_rules), [])

    def test_conditionals_issue_31(self):
        rules = [
            ConditionalFormatRule(
                ranges=[GridRange(self.sheet.id, 1, 1, 1, 2)],
                booleanRule=BooleanRule(
                    BooleanCondition('NUMBER_EQ', ['1']),
                    CellFormat(textFormat=TextFormat(foregroundColor=Color.fromHex("#000000")))
                )
            ),
            ConditionalFormatRule(
                ranges=[GridRange(self.sheet.id, 2, 3, 2, 3)],
                booleanRule=BooleanRule(
                    BooleanCondition('NUMBER_EQ', ['1']),
                    CellFormat(textFormat=TextFormat(foregroundColor=Color.fromHex("#00FFFF")))
                )
           ),
            ConditionalFormatRule(
                ranges=[GridRange(self.sheet.id, 1, 2, 1, 3)],
                booleanRule=BooleanRule(
                    BooleanCondition('NUMBER_EQ', ['1']),
                    CellFormat(textFormat=TextFormat(foregroundColor=Color.fromHex("#FFFF00")))
                )
            )
        ]
        rows = [
            ["A", "B", "C", "D"],
            ["1", "2", "3", "4"]
        ]
        cell_list = self.sheet.range('A1:D2')
        for cell, value in zip(cell_list, itertools.chain(*rows)):
            cell.value = value
        current_rules = get_conditional_format_rules(self.sheet)
        current_rules.extend(rules)
        self.assertNotEqual(current_rules.save(), None)
        current_rules_fetched = get_conditional_format_rules(self.sheet)
        self.assertEqual(
            current_rules_fetched.rules[0].booleanRule.format.textFormat.foregroundColor,
            current_rules.rules[0].booleanRule.format.textFormat.foregroundColor
        )
        self.assertEqual(
            current_rules_fetched.rules[1].booleanRule.format.textFormat.foregroundColor,
            current_rules.rules[1].booleanRule.format.textFormat.foregroundColor
        )
        self.assertEqual(
            current_rules_fetched.rules[2].booleanRule.format.textFormat.foregroundColor,
            current_rules.rules[2].booleanRule.format.textFormat.foregroundColor
        )


    def test_dataframe_formatter(self):
        rows = [  
            {
                'i': i,
                'j': i * 2,
                'A': 'Label ' + str(i), 
                'B': i * 100 + 2.34, 
                'C': date(2019, 3, i % 31 + 1), 
                'D': datetime(2019, 3, i % 31 + 1, i % 24, i % 60, i % 60),
                'E': i * 1000 + 7.8001, 
            } 
            for i in range(200) 
        ]
        df = pd.DataFrame.from_records(rows, index=['i', 'j'])
        set_with_dataframe(self.sheet, df, include_index=True)
        format_with_dataframe(
            self.sheet, 
            df, 
            formatter=BasicFormatter.with_defaults(
                freeze_headers=True, 
                column_formats={
                    'C': cellFormat(
                            numberFormat=numberFormat(type='DATE', pattern='yyyy mmmmmm dd'), 
                            horizontalAlignment='CENTER'
                        ),
                    'E': cellFormat(
                            numberFormat=numberFormat(type='NUMBER', pattern='[Color23][>40000]"HIGH";[Color43][<=10000]"LOW";0000'), 
                            horizontalAlignment='CENTER'
                        )
                }
            ), 
            include_index=True,
        )
        for cell_range, expected_uef in [
            ('A2:A201', cellFormat(numberFormat=numberFormat(type='NUMBER'), horizontalAlignment='RIGHT')), 
            ('B2:B201', cellFormat(numberFormat=numberFormat(type='NUMBER'), horizontalAlignment='RIGHT')), 
            ('C2:C201', cellFormat(horizontalAlignment='CENTER')), 
            ('D2:D201', cellFormat(numberFormat=numberFormat(type='NUMBER'), horizontalAlignment='RIGHT')), 
            ('E2:E201', 
                cellFormat(
                    numberFormat=numberFormat(type='DATE', pattern='yyyy mmmmmm dd'), 
                    horizontalAlignment='CENTER'
                )
            ), 
            ('F2:F201', cellFormat(numberFormat=numberFormat(type='DATE'), horizontalAlignment='CENTER')), 
            ('G2:G201', 
                cellFormat(
                    numberFormat=numberFormat(
                        type='NUMBER', 
                        pattern='[Color23][>40000]"HIGH";[Color43][<=10000]"LOW";0000'
                    ), 
                    horizontalAlignment='CENTER'
                )
            ), 
            ('A1:B201', 
                cellFormat(
                    backgroundColor=DEFAULT_HEADER_BACKGROUND_COLOR,
                    textFormat=textFormat(bold=True)
                )
            ), 
            ('A1:G1', 
                cellFormat(
                    backgroundColor=DEFAULT_HEADER_BACKGROUND_COLOR,
                    textFormat=textFormat(bold=True)
                )
            )
            ]:
            start_cell, end_cell = cell_range.split(':')
            for cell in (start_cell, end_cell):
                actual_uef = get_user_entered_format(self.sheet, cell)
                # actual_uef must be a superset of expected_uef
                self.assertTrue(
                    actual_uef & expected_uef == expected_uef, 
                    "%s range expected format %s, got %s" % (cell_range, expected_uef, actual_uef)
                )
        self.assertEqual(1, get_frozen_row_count(self.sheet))
        self.assertEqual(2, get_frozen_column_count(self.sheet))

    def test_row_height_and_column_width(self):
        set_row_height(self.sheet, '1:5', 42)
        set_column_width(self.sheet, 'A', 187)
        metadata = self.sheet.spreadsheet.fetch_sheet_metadata({'includeGridData': 'true'})
        sheet_md = [ s for s in metadata['sheets'] if s['properties']['sheetId'] == self.sheet.id ][0]
        row_md = sheet_md['data'][0]['rowMetadata']
        col_md = sheet_md['data'][0]['columnMetadata']
        for row in row_md[0:4]:
            self.assertEqual(42, row['pixelSize'])
        for col in col_md[0:1]:
            self.assertEqual(187, col['pixelSize'])

    def test_row_height_and_column_width_batch(self):
        with batch_updater(self.sheet.spreadsheet) as batch:
            batch.set_row_height(self.sheet, '1:5', 42)
            batch.set_column_width(self.sheet, 'A', 187)
        metadata = self.sheet.spreadsheet.fetch_sheet_metadata({'includeGridData': 'true'})
        sheet_md = [ s for s in metadata['sheets'] if s['properties']['sheetId'] == self.sheet.id ][0]
        row_md = sheet_md['data'][0]['rowMetadata']
        col_md = sheet_md['data'][0]['columnMetadata']
        for row in row_md[0:4]:
            self.assertEqual(42, row['pixelSize'])
        for col in col_md[0:1]:
            self.assertEqual(187, col['pixelSize'])

    def test_batch_updater_different_spreadsheet(self):
        batch = batch_updater(self.sheet.spreadsheet)
        other_spread = gspread.Spreadsheet(None, {'id': 'blech', 'title': 'Other sheet'})
        other_sheet = gspread.Worksheet(other_spread, {'sheetId': 4, 'title': 'Bleh'})
        batch.set_row_height(self.sheet, '1:5', 42)
        with self.assertRaises(ValueError):
            batch.set_row_height(other_sheet, '1:5', 42)

    def test_batch_updater_context(self):
        batch = batch_updater(self.sheet.spreadsheet)
        batch.set_row_height(self.sheet, '1:5', 42)
        batch.set_column_width(self.sheet, 'A', 187)
        self.assertEqual(2, len(batch.requests))
        try:
            with batch:
                batch.set_row_height(self.sheet, '1:5', 40)
        except Exception as e:
            self.assertIsInstance(e, IOError)
        self.assertEqual(2, len(batch.requests))
        batch.execute()
        self.assertEqual(0, len(batch.requests))


class ColorTest(unittest.TestCase):

    SAMPLE_HEXSTRINGS_NOALPHA = ['#230ac7','#9ec08b','#037778','#134d70','#f1f974','#0997b6','#42da14','#be5ee8']
    SAMPLE_HEXSTRINGS_ALPHA = ['#b7d90600','#0a29f321','#33db6a48','#4134a467','#7d172388','#58fe5fa1','#2ea14ecc','#c18de9f8']
    SAMPLE_HEXSTRING_CAPS = ['#DDEEFF','#EEFFAABB','#1A2B3C4E','#A1F2B3']
    # [NO_POUND_SIGN, NO_POUND_SIGN_ALPHA, INVALID_HEX_CHAR, INVALID_HEX_CHAR_ALPHA, SPECIAL_INVALID_CHAR, TOO_FEW_CHARS, TOO_MANY_CHARS]
    SAMPLE_HEXSTRINGS_BAD = ['230ac7','9ec08b9b','#Adbeye','#1122ccgg','#11$100FF', '#11678','#867530910']

    def test_color_roundtrip(self):
        for hexstring in self.SAMPLE_HEXSTRINGS_NOALPHA:
            self.assertEqual(hexstring, Color.fromHex(hexstring).toHex())
        for hexstring in self.SAMPLE_HEXSTRINGS_ALPHA:
            self.assertEqual(hexstring, Color.fromHex(hexstring).toHex())
        for hexstring in self.SAMPLE_HEXSTRING_CAPS:
            # Check equality with lowercase version of string
            self.assertEqual(hexstring.lower(), Color.fromHex(hexstring).toHex())

    def test_color_malformed(self):
        for hexstring in self.SAMPLE_HEXSTRINGS_BAD:
            with self.assertRaises(ValueError):
                Color.fromHex(hexstring)


class FormattingComponentTest(unittest.TestCase):

    def test_repr_and_equality(self):
        comp = TextFormat(bold=True, italic=True)
        comp2 = TextFormat(bold=True)
        self.assertEqual('<TextFormat bold=True;italic=True>', repr(comp))
        self.assertNotEqual(comp, comp2)
        self.assertNotEqual(comp, None)
        self.assertEqual(comp, comp)

    def test_number_format_types(self):
        for type_ in NumberFormat.TYPES:
            f = NumberFormat(type_)
            self.assertEqual(f, f)
            self.assertEqual(type_, f.type)
        with self.assertRaises(ValueError):
            NumberFormat('BAD_TYPE')

    def test_border_styles(self):
        for style in Border.STYLES:
            f = Border(style)
            self.assertEqual(f, f)
            self.assertEqual(style, f.style)
        with self.assertRaises(ValueError):
            Border('BAD_STYLE')

    def test_text_format_link(self):
        TextFormat(link=None)
        TextFormat(link=Link("https://foo.com/"))
        tf = TextFormat(link=Link(uri="https://foo.com/"))
        self.assertEqual("https://foo.com/", tf.link.uri)
        tf2 = TextFormat.from_props(tf.to_props())
        self.assertEqual(tf, tf2)

    def test_text_rotation_exclusion(self):
        TextRotation(angle=1)
        TextRotation(vertical=True)
        with self.assertRaises(ValueError):
            TextRotation(angle=1, vertical=True)
        with self.assertRaises(ValueError):
            TextRotation()

class GridRangeTest(unittest.TestCase):

    def test_absent_sheet_id(self):
        gr = GridRange.from_props({'startRowIndex': 1})
        self.assertEqual(0, gr.sheetId)
        self.assertEqual(1, gr.startRowIndex)
