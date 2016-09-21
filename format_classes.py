"""
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
"""

## I don't plan on doing much more with this, but there are some
## easy improvements to make. Rows & columns need to be processed
## similarly to cells, as they overwrite each other right now
## and don't maintain a memory of old row/column formats.
## There are some terrible names, like override_cell_format.
## Book.add_book_sheet() in particular is bad. It's mainly
## copy/pasted from xlsxwriter, so an update there could break it.
## These functions need tests. I'm sure there are errors that
## my limited use-cases haven't seen.


import xlsxwriter
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet, convert_cell_args, \
	convert_range_args, convert_column_args
from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol


class Book(Workbook):

	def write_cells(self, sheet):
		"""Finally write values & formats to cells"""
		for cell in sheet.cells_to_write:
			# If there's a formatting applied to the whole sheet:
			if sheet.worksheet_format:
				format_copy = {key: sheet.worksheet_format[key] for key in \
					sheet.worksheet_format}
				format_copy.update(sheet.cells_to_write[cell][1])
				sheet.write(cell, sheet.cells_to_write[cell][0],
					self.add_format(format_copy))
			else:
				sheet.write(cell, sheet.cells_to_write[cell][0],
					self.add_format(sheet.cells_to_write[cell][1]))

	def format_rows(self, sheet):
		# row_formats = [(row, height, format, options, override_cell_format), ...]
		for item in sheet.row_formats:
			if sheet.worksheet_format:
				temp_dict = {key: sheet.worksheet_format[key] for key in \
					sheet.worksheet_format}
				temp_dict.update(item[2])
				sheet.set_row(item[0], item[1], self.add_format(temp_dict),
					item[3])
			else:
				sheet.set_row(item[0], item[1], self.add_format(item[2]),
					item[3])
		
			if item[2] and item[4]:
				# Look for any cells that need their formatting augmented
				for key in sheet.cells_to_write:
					row, col = xl_cell_to_rowcol(key)
					if row == item[0]:
						if sheet.cells_to_write[key][1]:
							sheet.cells_to_write[key][1].update(item[2])
						else:
							sheet.cells_to_write[key] = (
								sheet.cells_to_write[key][0], item[2])

	def format_columns(self, sheet):
		"""Writes columns from sheet.column_formats[]"""
		for item in sheet.column_formats:
			# Add any worksheet-wide formatting, if it exists
			if sheet.worksheet_format:
				temp_dict = {key: sheet.worksheet_format[key] for key in \
					sheet.worksheet_format}
				temp_dict.update(item[3])
				sheet.set_column(item[0], item[1], item[2],
					self.add_format(temp_dict), item[4])
			else:
				sheet.set_column(item[0], item[1], item[2],
					self.add_format(item[3]), item[4])

			if item[3] and item[5]:
				# If there's a format, and you should override other formats:
				for key in sheet.cells_to_write:
					row, col = xl_cell_to_rowcol(key)
					if item[0] <= col <= item[1]:
						# Priority of formats: cell > row
						updated_format = {key: item[3][key] for key in item[3]}
						grrr = {d_key: sheet.cells_to_write[key][1][d_key] \
							for d_key in sheet.cells_to_write[key][1]}
						sheet.cells_to_write[key] = (
							sheet.cells_to_write[key][0],
							updated_format.update(grrr))

	def add_book_sheet(self, name):
		"""I don't see a good way to interject the type of object I want
			created instead of the default type, so I'm largely copying
			the code from xlsxwriter.workbook._add_sheet()"""
		sheet_index = len(self.worksheets_objs)
		# name = self._check_sheetname(name, False)

		# Initialization data to pass to the worksheet.
		init_data = {
			'name': name,
			'index': sheet_index,
			'str_table': self.str_table,
			'worksheet_meta': self.worksheet_meta,
			'optimization': self.optimization,
			'tmpdir': self.tmpdir,
			'date_1904': self.date_1904,
			'strings_to_numbers': self.strings_to_numbers,
			'strings_to_formulas': self.strings_to_formulas,
			'strings_to_urls': self.strings_to_urls,
			'nan_inf_to_errors': self.nan_inf_to_errors,
			'default_date_format': self.default_date_format,
			'default_url_format': self.default_url_format,
			'excel2003_style': self.excel2003_style,
		}

		worksheet = Sheet()

		worksheet._initialize(init_data)

		self.worksheets_objs.append(worksheet)
		self.sheetnames[name] = worksheet

		return worksheet

	def close(self, *args, **kwargs):
		for sheet in self.worksheets():
			if sheet.worksheet_format:
				sheet.set_column("A:XFD", None,
					self.add_format(sheet.worksheet_format))
			self.format_columns(sheet)
			self.format_rows(sheet)
			self.write_cells(sheet)
		super(Book, self).close(*args, **kwargs)


class Sheet(Worksheet):

	def __init__(self, *args, **kwargs):
		super(Sheet, self).__init__(*args, **kwargs)
		self.cells_to_write = {}  # = {location: (what_to_write, format), ...}
		self.row_formats = []
		self.column_formats = []
		self.worksheet_format = {}
	
	def write_sheet(self, format):
		"""Apply a dictionary of formats to the entire worksheet"""
		self.worksheet_format.update(format)

	@convert_cell_args
	def cell_writer(self, row, col, write_this, format=False):
	# def cell_writer(self, location, write_this, format=False):
		"""Call this instead of worksheet.write(). Stores information about
		what should be written, and adds to the formatting of cells instead
		of replacing old formatting when a new format is overlayed"""
		# The @convert_cell_args decorator converts any "A1" style input into
		# the row, col format so that you can use the formats interchangeably.
		# But cells_to_write dict is easier to deal with in the "A1" style,
		# so convert everything back to that.
		location = xl_rowcol_to_cell(row, col)

		if format:
			# Augment the format if it already exists, instead of overwriting it
			if location in self.cells_to_write:
				if self.cells_to_write[location][1]:
					# Copy the dict since .update() doesn't work
					grrr = {key: self.cells_to_write[location][1][key] for key in \
						self.cells_to_write[location][1]}
					grrr.update(format)
					self.cells_to_write[location] = (write_this, grrr)
				else:
					self.cells_to_write[location] = (write_this, format)
			# Otherwise just add the format
			else:
				self.cells_to_write[location] = (write_this, format)
		else:
			# If the cell already exists and has a format:
			if location in self.cells_to_write:
				self.cells_to_write[location] = (write_this, 
					self.cells_to_write[location][1])
			else:
				self.cells_to_write[location] = (write_this, {})

	@convert_column_args
	def write_column(self, first_col, last_col, width=None, cell_format={},
		options={}, override_cell_format=False):
		"""This is an analogue for set_column(). It stores the arguments
		for set_column() in self.column_formats to be written later by the
		workbook in format_columns()"""
		### This isn't ideal. Formatting still breaks when you write one
		### column on top of a previously-written column. Should fix
		### this the same way as cells, where there's a dictionary keeping
		### track of what columns have previously been formatted & updates
		### that formatting instead of replacing it.
		self.column_formats += [(first_col, last_col, width, cell_format,
			options, override_cell_format)]

	def write_row(self, row, height=None, cell_format=None, options={},
		override_cell_format=False):
		"""This is an analogue for set_row(). It stores the arguments
		for set_row() in self.row_formats to be written later by the
		workbook in format_rows()"""
		self.row_formats += [(row, height, cell_format, options,
			override_cell_format)]

	@convert_range_args
	def box(self, row_1, col_1, row_2, col_2, border_style=1,
		border_color='black', pattern=0, bg_color=0, fg_color=0):
		"""Makes an RxC box. Use integers, not the 'A1' format"""

		rows = row_2 - row_1 + 1
		cols = col_2 - col_1 + 1

		for x in range((rows) * (cols)): # Total number of cells in the rectangle
			box_form = {}   # The format resets each loop
			row = row_1 + (x // cols)
			column = col_1 + (x % cols)

			if x < (cols):                     # If it's on the top row
				box_form['top'] = border_style
			if x >= ((rows * cols) - cols):    # If it's on the bottom row
				box_form['bottom'] = border_style
			if x % cols == 0:                  # If it's on the left column
				box_form['left'] = border_style
			if x % cols == (cols - 1):         # If it's on the right column
				box_form['right'] = border_style

			if box_form != {}:
				if border_color:
					box_form['border_color'] = border_color
			for item in [pattern, bg_color, fg_color]:
				if item:
					box_form[item] = item

			self.cell_writer(row, column, "", box_form)
