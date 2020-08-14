# -*- coding: utf-8 -*-

import gspread_formatting.functions
import gspread_formatting.dataframe

from functools import wraps

__all__ = ('batch_updater', 'SpreadsheetBatchUpdater')

def batch_updater(spreadsheet):
    return SpreadsheetBatchUpdater(spreadsheet)

class SpreadsheetBatchUpdater(object):
    def __init__(self, spreadsheet):
        self.spreadsheet = spreadsheet
        self.requests = []

    def __enter__(self):
        if self.requests:
            raise IOError("BatchUpdater has un-executed requests pending, cannot be __enter__ed")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.execute()
        return False

    def execute(self):
        resps = self.spreadsheet.batch_update({'requests': self.requests})
        del self.requests[:]
        return resps

def _wrap_for_batch_updater(func):
    @wraps(func)
    def f(self, worksheet, *args, **kwargs):
        if worksheet.spreadsheet != self.spreadsheet:
            raise ValueError(
                "Worksheet %r belongs to spreadsheet %r, not batch updater's spreadsheet %r" 
                % (worksheet, worksheet.spreadsheet, self.spreadsheet)
            )
        self.requests.append( func(worksheet, *args, **kwargs) )
        return self
    return f

for fname in gspread_formatting.batch_update_requests.__all__:
    func = getattr(gspread_formatting.batch_update_requests, fname)
    setattr(SpreadsheetBatchUpdater, fname, _wrap_for_batch_updater(func))

setattr(
    SpreadsheetBatchUpdater, 
    'format_with_dataframe', 
    _wrap_for_batch_updater(gspread_formatting.dataframe._format_with_dataframe)
)

