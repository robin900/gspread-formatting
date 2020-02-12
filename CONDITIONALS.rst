Conditional Format Rules
~~~~~~~~~~~~~~~~~~~~~~~~

A conditional format rule allows you to specify a cell format that (additively) applies to cells in certain ranges
only when the value of the cell meets a certain condition. 
The `ConditionalFormatRule documentation <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/sheets#ConditionalFormatRule>`_ for the Sheets API describes the two kinds of rules allowed:
a ``BooleanRule`` in which the `CellFormat` is applied to the cell if the value meets the specified boolean
condition; or a ``GradientRule`` in which the ``Color`` or ``ColorStyle`` of the cell varies depending on the numeric
value of the cell or cells. 

You can specify multiple rules for each worksheet present in a Google spreadsheet. To add or remove rules,
use the ``get_conditional_format_rules(worksheet)`` function, which returns a list-like object which you can
modify as you would modify a list, and then call ``.save()`` to store the rule changes you've made.

Here is an example that applies bold text and a bright red color to cells in column A if the cell value
is numeric and greater than 100::

    from gspread_formatting import *

    worksheet = some_spreadsheet.worksheet('My Worksheet')

    rule = ConditionalFormatRule(
        ranges=[GridRange.from_a1_range('A1:A2000', worksheet)],
        booleanRule=BooleanRule(
            condition=BooleanCondition('NUMBER_GREATER', '100'), 
            format=CellFormat(textFormat=textFormat(bold=True), color=Color(1,0,0))
        )
    )

    rules = get_conditional_format_rules(worksheet)
    rules.append(rule)
    rules.save()

    # or, to replace any existing rules with just your single rule:
    rules.clear()
    rules.append(rule)
    rules.save()

An important note: A ``ConditionalFormatRule`` is, like all other objects provided by this package,
mutable in all of its fields. Mutating a ``ConditionalFormatRule`` object in place will not automatically
store the changes via the Sheets API; but calling `.save()` on the list-like rules object will store
the mutated rule as expected.
