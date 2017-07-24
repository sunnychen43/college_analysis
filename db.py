from openpyxl import load_workbook
from openpyxl import workbook
import classifier_lists
import excel


# Create headers for database
def create_headers(wb):
	ws = wb.active
	for key, value in classifier_lists.headers.items():
		ws[value + '1'].value = key
	print('Create header done')


# Moves raw data into db
def raw_to_db(raw_wb, db_wb, college, year, type):
	raw_ws = raw_wb.active
	db_ws = db_wb.active
	max_db_row = db_ws.max_row
	num_entries = raw_ws.max_row - 1

	for entry in range(num_entries):
		raw_row = entry + 2  # Start scanning on the second line

		entry_id = raw_ws['A' + str(raw_row)].value
		if entry_id is None:  # Blank line
			continue

		db_row = max_db_row + int(entry_id)  # Start at first free line

		if db_ws['A' + str(db_row)].value is None:  # Create entry if it doesn't exist
			db_ws['A' + str(db_row)].value = db_row - 1
			db_ws['B' + str(db_row)].value = college
			db_ws['C' + str(db_row)].value = year
			db_ws['D' + str(db_row)].value = type

		entry_class = raw_ws['B' + str(raw_row)].value
		if entry_class is not None:
			try:
				db_column = classifier_lists.headers[entry_class]  # Sets column under matching header
			except KeyError:  # If class is not appropriate header
				continue
			db_ws[db_column + str(db_row)].value = raw_ws['C' + str(raw_row)].value  # Copies raw data to db
	print('Convert to db done')


# Main method, creates db
def create_db(url, file_name, school, year, type):
	raw_wb = workbook.Workbook()
	excel.get_comments(url, raw_wb)
	excel.classify(raw_wb)
	excel.collate(raw_wb)

	try:
		db_wb = load_workbook(file_name)
	except FileNotFoundError:
		db_wb = workbook.Workbook()
	create_headers(db_wb)
	raw_to_db(raw_wb, db_wb, school, year, type)
	db_wb.save(file_name)


def clean_blank_lines(wb):
	old_wb = wb
	old_ws = old_wb.active
	new_wb = workbook.Workbook()
	new_ws = new_wb.active
	max_row = old_ws.max_row
	max_column = old_ws.max_column

	for i in range(max_row):
		row = i + 1
		print(row, '/', max_row)
		if old_ws['F' + str(row)].value is None:
			if old_ws['H' + str(row)].value is None:
				if old_ws['I' + str(row)].value is None:
					continue
		for i in range(max_column):
			column = i + 1
			new_ws.cell(row=row, column=column).value = old_ws.cell(row=row, column=column).value
	wb = new_wb