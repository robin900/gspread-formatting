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


class UtilsTest(unittest.TestCase):

    def test_extract_id_from_url(self):
        url_id_list = [
            # New-style url
            ('https://docs.google.com/spreadsheets/d/'
             '1qpyC0X3A0MwQoFDE8p-Bll4hps/edit#gid=0',
             '1qpyC0X3A0MwQoFDE8p-Bll4hps'),

            ('https://docs.google.com/spreadsheets/d/'
             '1qpyC0X3A0MwQoFDE8p-Bll4hps/edit',
             '1qpyC0X3A0MwQoFDE8p-Bll4hps'),

            ('https://docs.google.com/spreadsheets/d/'
             '1qpyC0X3A0MwQoFDE8p-Bll4hps',
             '1qpyC0X3A0MwQoFDE8p-Bll4hps'),

            # Old-style url
            ('https://docs.google.com/spreadsheet/'
             'ccc?key=1qpyC0X3A0MwQoFDE8p-Bll4hps&usp=drive_web#gid=0',
             '1qpyC0X3A0MwQoFDE8p-Bll4hps')
        ]

        for url, id in url_id_list:
            self.assertEqual(id, utils.extract_id_from_url(url))

    def test_no_extract_id_from_url(self):
        self.assertRaises(
            gspread.NoValidUrlKeyFound,
            utils.extract_id_from_url,
            'http://example.org'
        )

    def test_a1_to_rowcol(self):
        self.assertEqual(utils.a1_to_rowcol('ABC3'), (3, 731))

    def test_rowcol_to_a1(self):
        self.assertEqual(utils.rowcol_to_a1(3, 731), 'ABC3')
        self.assertEqual(utils.rowcol_to_a1(1, 104), 'CZ1')

    def test_addr_converters(self):
        for row in range(1, 257):
            for col in range(1, 512):
                addr = utils.rowcol_to_a1(row, col)
                (r, c) = utils.a1_to_rowcol(addr)
                self.assertEqual((row, col), (r, c))

    def test_get_gid(self):
        gid = 'od6'
        self.assertEqual(utils.wid_to_gid(gid), '0')
        gid = 'osyqnsz'
        self.assertEqual(utils.wid_to_gid(gid), '1751403737')
        gid = 'ogsrar0'
        self.assertEqual(utils.wid_to_gid(gid), '1015761654')


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


class ClientTest(GspreadTest):

    """Test for gspread.client."""

    def test_open(self):
        title = self.config.get('Spreadsheet', 'title')
        spreadsheet = self.gc.open(title)
        self.assertTrue(isinstance(spreadsheet, gspread.models.Spreadsheet))

    def test_no_found_exeption(self):
        noexistent_title = "Please don't use this phrase as a name of a sheet."
        self.assertRaises(gspread.SpreadsheetNotFound,
                          self.gc.open,
                          noexistent_title)

    def test_open_by_key(self):
        key = self.config.get('Spreadsheet', 'key')
        spreadsheet = self.gc.open_by_key(key)
        self.assertTrue(isinstance(spreadsheet, gspread.models.Spreadsheet))

    def test_open_by_url(self):
        url = self.config.get('Spreadsheet', 'url')
        spreadsheet = self.gc.open_by_url(url)
        self.assertTrue(isinstance(spreadsheet, gspread.models.Spreadsheet))

    def test_openall(self):
        spreadsheet_list = self.gc.openall()
        for s in spreadsheet_list:
            self.assertTrue(isinstance(s, gspread.models.Spreadsheet))

    def test_create(self):
        title = gen_value('TestSpreadsheet')
        new_spreadsheet = self.gc.create(title)
        self.assertTrue(
            isinstance(new_spreadsheet, gspread.models.Spreadsheet))

    def test_import_csv(self):
        title = gen_value('TestImportSpreadsheet')
        new_spreadsheet = self.gc.create(title)

        csv_rows = 4
        csv_cols = 4

        rows = [[
            gen_value('%s-%s' % (i, j))
            for j in range(csv_cols)]
            for i in range(csv_rows)
        ]

        simple_csv_data = '\n'.join([','.join(row) for row in rows])

        self.gc.import_csv(new_spreadsheet.id, simple_csv_data)

        sh = self.gc.open_by_key(new_spreadsheet.id)
        self.assertEqual(sh.sheet1.get_all_values(), rows)

        self.gc.del_spreadsheet(new_spreadsheet.id)


class SpreadsheetTest(GspreadTest):

    """Test for gspread.Spreadsheet."""

    def setUp(self):
        super(SpreadsheetTest, self).setUp()
        title = self.config.get('Spreadsheet', 'title')
        self.spreadsheet = self.gc.open(title)

    def test_properties(self):
        self.assertEqual(self.config.get('Spreadsheet', 'id'),
                         self.spreadsheet.id)
        self.assertEqual(self.config.get('Spreadsheet', 'title'),
                         self.spreadsheet.title)

    def test_sheet1(self):
        sheet1 = self.spreadsheet.sheet1
        self.assertTrue(isinstance(sheet1, gspread.Worksheet))

    def test_get_worksheet(self):
        sheet1 = self.spreadsheet.get_worksheet(0)
        self.assertTrue(isinstance(sheet1, gspread.Worksheet))

    def test_worksheet(self):
        sheet_title = self.config.get('Spreadsheet', 'sheet1_title')
        sheet = self.spreadsheet.worksheet(sheet_title)
        self.assertTrue(isinstance(sheet, gspread.Worksheet))

    def test_worksheet_iteration(self):
        self.assertEqual(
            [x.id for x in self.spreadsheet.worksheets()],
            [sheet.id for sheet in self.spreadsheet]
        )


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
        format_cell_range(self.sheet, 'A1:D6', fmt)
        ue_fmt = get_user_entered_format(self.sheet, 'A1')
        self.assertEqual(ue_fmt.textFormat.bold, True)
        eff_fmt = get_effective_format(self.sheet, 'A1')
        self.assertEqual(eff_fmt.textFormat.bold, True)

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
