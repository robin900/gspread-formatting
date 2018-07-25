# -*- coding: utf-8 -*-

import os
import re
import random
import unittest
import itertools
import uuid

try:
    import ConfigParser
except ImportError:
    import configparser as ConfigParser

from oauth2client.service_account import ServiceAccountCredentials

import gspread
from gspread import utils
from gspread_formatting import *

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


def read_config(filename):
    config = ConfigParser.ConfigParser()
    config.readfp(open(filename))
    return config


def read_credentials(filename):
    return ServiceAccountCredentials.from_json_keyfile_name(filename, SCOPE)


def gen_value(prefix=None):
    if prefix:
        return u'%s %s' % (prefix, gen_value())
    else:
        return unicode(uuid.uuid4())


class GspreadTest(unittest.TestCase):
    config = None
    gc = None

    @classmethod
    def setUpClass(cls):
        try:
            cls.config = read_config(CONFIG_FILENAME)
            credentials = read_credentials(CREDS_FILENAME)
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
        title = cls.config.get('Spreadsheet', 'title')
        cls.spreadsheet = cls.gc.open(title)
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
        # NOTE(msuozzo): Here, a new worksheet is created for each test.
        # This was determined to be faster than reusing a single sheet and
        # having to clear its contents after each test.
        # Basically: Time(add_wks + del_wks) < Time(range + update_cells)
        self.sheet = self.spreadsheet.add_worksheet('wksht_test', 20, 20)

    def tearDown(self):
        self.spreadsheet.del_worksheet(self.sheet)

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

        fmt = cellFormat(textFormat=textFormat(bold=True))
        format_cell_ranges(self.sheet, [('A1:B6', fmt), ('C1:D6', fmt)])
        ue_fmt = get_user_entered_format(self.sheet, 'A1')
        self.assertEqual(ue_fmt.textFormat.bold, True)
        eff_fmt = get_effective_format(self.sheet, 'A1')
        self.assertEqual(eff_fmt.textFormat.bold, True)
        fmt2 = cellFormat(textFormat=textFormat(italic=True))
        format_cell_range(self.sheet, 'A1:D6', fmt2)
        ue_fmt = get_user_entered_format(self.sheet, 'A1')
        self.assertEqual(ue_fmt.textFormat.italic, True)
        eff_fmt = get_effective_format(self.sheet, 'A1')
        self.assertEqual(eff_fmt.textFormat.italic, True)

    def test_frozen_rows_cols(self):
        set_frozen(self.sheet, rows=1, cols=1)
        fresh = self.sheet.spreadsheet.fetch_sheet_metadata({'includeGridData': True})
        item = utils.finditem(lambda x: x['properties']['title'] == self.sheet.title, fresh['sheets'])
        pr = item['properties']['gridProperties']
        self.assertEqual(pr.get('frozenRowCount'), 1)
        self.assertEqual(pr.get('frozenColumnCount'), 1)
        self.assertEqual(get_frozen_row_count(self.sheet), 1)
        self.assertEqual(get_frozen_column_count(self.sheet), 1)

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
