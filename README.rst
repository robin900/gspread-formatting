gspread-formatting
------------------

.. image:: https://badge.fury.io/py/gspread-formatting.svg
    :target: https://badge.fury.io/py/gspread-formatting

.. image:: https://travis-ci.org/robin900/gspread-formatting.svg?branch=master
    :target: https://travis-ci.org/robin900/gspread-formatting

This package provides complete support of basic cell formatting for Google spreadsheets
to the popular ``gspread`` package, along with a few related features such as setting
"frozen" rows and columns in a worksheet.

The package also offers graceful formatting of Google spreadsheets using a Pandas DataFrame.
See the section below for usage and details.

Usage
~~~~~

Basic formatting of a range of cells in a worksheet is offered by the ``format_cell_range`` function. 
All basic formatting components of the v4 Sheets API's ``CellFormat`` are present as classes 
in the ``gspread_formatting`` module, available both by ``InitialCaps`` names and ``camelCase`` names: 
for example, the background color class is ``BackgroundColor`` but is also available as 
``backgroundColor``, while the color class is ``Color`` but available also as ``color``. 
Attributes of formatting components are best specified as keyword arguments using ``camelCase`` 
naming, e.g. ``backgroundColor=...``. Complex formats may be composed easily, by nesting the calls to the classes.  

See `the CellFormat page of the Sheets API documentation <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#CellFormat>`_
to learn more about each formatting component.::

    from gspread_formatting import *

    fmt = cellFormat(
        backgroundColor=color(1, 0.9, 0.9),
        textFormat=textFormat(bold=True, foregroundColor=color(1, 0, 1)),
        horizontalAlignment='CENTER'
        )

    format_cell_range(worksheet, 'A1:J1', fmt)

The ``format_cell_ranges`` function allows for formatting multiple ranges with corresponding formats,
all in one function call and Sheets API operation::

    fmt = cellFormat(
        backgroundColor=color(1, 0.9, 0.9),
        textFormat=textFormat(bold=True, foregroundColor=color(1, 0, 1)),
        horizontalAlignment='CENTER'
        )

    fmt2 = cellFormat(
        backgroundColor=color(0.9, 0.9, 0.9),
        horizontalAlignment='RIGHT'
        )

    format_cell_ranges(worksheet, [('A1:J1', fmt), ('K1:K200', fmt2)])

Retrieving, Comparing, and Composing CellFormats
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

A Google spreadsheet's own default format, as a CellFormat object, is available via ``get_default_format(spreadsheet)``.
``get_effective_format(worksheet, label)`` and ``get_user_entered_format(worksheet, label)`` also will return
for any provided cell label either a CellFormat object (if any formatting is present) or None.

``CellFormat`` objects are comparable with ``==`` and ``!=``, and are mutable at all times; 
they can be safely copied with Python's ``copy.deepcopy`` function. ``CellFormat`` objects can be combined
into a new ``CellFormat`` object using the ``add`` method (or ``+`` operator). ``CellFormat`` objects also offer 
``difference`` and ``intersection`` methods, as well as the corresponding
operators ``-`` (for difference) and ``&`` (for intersection).::

    >>> default_format = CellFormat(backgroundColor=color(1,1,1), textFormat=textFormat(bold=True))
    >>> user_format = CellFormat(textFormat=textFormat(italic=True))
    >>> effective_format = default_format + user_format
    >>> effective_format
    CellFormat(backgroundColor=color(1,1,1), textFormat=textFormat(bold=True, italic=True))
    >>> effective_format - user_format 
    CellFormat(backgroundColor=color(1,1,1), textFormat=textFormat(bold=True))
    >>> effective_format - user_format == default_format
    True

Frozen Rows and Columns
~~~~~~~~~~~~~~~~~~~~~~~

The following functions get or set "frozen" row or column counts for a worksheet::

    get_frozen_row_count(worksheet)
    get_frozen_column_count(worksheet)
    set_frozen(worksheet, rows=1)
    set_frozen(worksheet, cols=1)
    set_frozen(worksheet, rows=1, cols=0)

Getting and Setting Data Validation Rules for Cells and Cell Ranges
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

The following functions get or set the "data validation rule" for a cell or cell range::

    get_data_validation_rule(worksheet, label)
    set_data_validation_for_cell_range(worksheet, range, rule)
    set_data_validation_for_cell_ranges(worksheet, ranges)

The full functionality of data validation rules is supported: all of ``BooleanCondition``. 
See `the API documentation <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#DataValidationRule>`_
for more information. Here's a short example::

    validation_rule = DataValidationRule(
        BooleanCondition('ONE_OF_LIST', ['1', '2', '3', '4']),
        showCustomUi=True
    )
    set_data_validation_for_cell_range(worksheet, 'A2:D2', validation_rule)
    # data validation for A2
    eff_rule = get_data_validation_rule(worksheet, 'A2')
    eff_rule.condition.type
    >>> 'ONE_OF_LIST'
    eff_rule.showCustomUi
    >>> True
    # No data validation for A1
    eff_rule = get_data_validation_rule(worksheet, 'A1')
    eff_rule
    >>> None

Conditional Formatting Rules
~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Conditional format rules are supported by this package! See the `Conditional Format Rules docs <./CONDITIONALS.rst>`_.

Formatting a Worksheet Using a Pandas DataFrame
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

If you are using Pandas DataFrames to provide data to a Google spreadsheet -- using perhaps
the ``gspread-dataframe`` package `available on PyPI <https://pypi.org/project/gspread-dataframe/>`_ --
the ``format_with_dataframe`` function in ``gspread_formatting.dataframe`` allows you to use that same 
DataFrame object and specify formatting for a worksheet. There is a ``DEFAULT_FORMATTER`` in the module,
which will be used if no formatter object is provided to ``format_with_dataframe``::

    from gspread_formatting.dataframe import format_with_dataframe, BasicFormatter
    from gspread_formatting import Color

    # uses DEFAULT_FORMATTER
    format_with_dataframe(worksheet, dataframe, include_index=True, include_column_header=True)

    formatter = BasicFormatter(
        header_background_color=Color(0,0,0), 
        header_text_color=Color(1,1,1),
        decimal_format='#,##0.00'
    )

    format_with_dataframe(worksheet, dataframe, formatter, include_index=False, include_column_header=True)

Installation
------------

Requirements
~~~~~~~~~~~~

* Python 2.6+ or Python 3+
* gspread >= 3.0.0

From PyPI
~~~~~~~~~

::

    pip install gspread-formatting

From GitHub
~~~~~~~~~~~

::

    git clone https://github.com/robin900/gspread-formatting.git
    cd gspread-formatting
    python setup.py install

Development and Testing
-----------------------

Install packages listed in ``requirements-dev.txt``. To run the test suite
in ``test.py`` you will need to:

* Authorize as the Google account you wish to use as a test, and download
  a JSON file containing the credentials. Name the file ``creds.json``
  and locate it in the top-level folder of the repository.
* Set up a ``tests.config`` file using the ``tests.config.example`` file as a template.
  Specify the ID of a spreadsheet that the Google account you are using
  can access with write privileges.
