from openpyxl import load_workbook
from openpyxl import workbook
import expand_entry
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

		entry_id = raw_ws[classifier_lists.headers['ID:'] + str(raw_row)].value
		if entry_id is None:  # Blank line
			continue

		db_row = max_db_row + int(entry_id)  # Start at first free line

		if db_ws[classifier_lists.headers['ID:'] + str(db_row)].value is None:  # Create entry if it doesn't exist
			db_ws[classifier_lists.headers['ID:'] + str(db_row)].value = db_row - 1
			db_ws[classifier_lists.headers['College:'] + str(db_row)].value = college
			db_ws[classifier_lists.headers['Year:'] + str(db_row)].value = year
			db_ws[classifier_lists.headers['Type:'] + str(db_row)].value = type

		entry_class = raw_ws[classifier_lists.headers['College:'] + str(raw_row)].value
		if entry_class is not None:
			try:
				db_column = classifier_lists.headers[entry_class]  # Sets column under matching header
			except KeyError:  # If class is not appropriate header
				continue
			db_ws[db_column + str(db_row)].value = raw_ws[classifier_lists.headers['Year:'] + str(raw_row)].value  # Copies raw data to db
	print('Convert to db done')


# Main method, creates db
def create(url, file_name, school, year, type):
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


def clean(wb):
	old_wb = wb
	old_ws = old_wb.active
	new_wb = workbook.Workbook()
	new_ws = new_wb.active
	max_row = old_ws.max_row
	max_column = old_ws.max_column

	for i in range(max_row):
		row = i + 1
		if old_ws[classifier_lists.headers['SAT I:'] + str(row)].value is None:
			if old_ws[classifier_lists.headers['SAT II:'] + str(row)].value is None:
				if old_ws[classifier_lists.headers['GPA:'] + str(row)].value is None:
					continue
		for i in range(max_column):
			column = i + 1
			new_ws.cell(row=row, column=column).value = old_ws.cell(row=row, column=column).value
	print('Clean done')
	return new_wb


def additional_decisions(wb):
	ws = wb.active
	max_row = ws.max_row
	for i in range(max_row):
		row = i + 2
		decisions = expand_entry.get_decisions(ws[classifier_lists.headers['Where else?'] + str(row)].value)
		if decisions is None:
			continue
		for decision in decisions:
			column = classifier_lists.headers[decision[0]]
			ws[column + str(row)].value = decision[1]
	print('Additional decisions done')

