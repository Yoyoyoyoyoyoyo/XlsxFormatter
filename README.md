# XlsxFormatter
A way to deal with the formatting challenges of Python's XlsxWriter


Formatting cells with Xlsxwriter can be a pain. Updating a cell's format means
    any 'memory' of the cell's previous format is lost instead of being augmented.
    This file is my quick & dirty approach to fixing that problem: aggregate all
    format requests before writing any cells. Format requests are given a
    hierarchy: requests across an entire sheet are lowest, then row/column
    formats, and cell formats trump all others. The hierarchy only applies when
    formats are exclusive: the worksheet background is grey but the cell
    should be green. Otherwise all formats are applied.
  Use these classes in place of Workbook and Worksheet. There are 5 main changes:
	First, instead of worksheet.write(...), use sheet.cell_writer(...). 
		The arguments are interchangeable between the two.
	Second, don't use worksheet.set_row(...) or worksheet.set_column(...).
		Instead use sheet.write_row(...) and sheet.write_column(...).
		These two also take an optional kwarg of "override_cell_format".
		If True, any row/column formatting will replace any clashing cell's
		format. The default is False.
	Third, apply a format to an entire worksheet with sheet.write_sheet().
	  The function's input is a dict of the formats you want to apply.
	Fourth, don't use workbook.add_format(). Just pass dicts of the format.
	Fifth, I threw in a function to make boxes. It gives you formatting
		control over border styles & colors, as well as cell patterns & colors.
