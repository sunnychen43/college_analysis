from openpyxl import load_workbook
import sys
import db

if __name__ == '__main__':
	if(len(sys.argv) != 3):
		print('Invalid arguments')
		quit()
	try:
		refrence_wb = load_workbook(sys.argv[1])
	except FileNotFoundError:
		print('Refrence File not Found')
		quit()

	refrence_ws = refrence_wb.active
	max_row = refrence_ws.max_row - 1
	for i in range(max_row):
		row = i + 2
		school = refrence_ws['A' + str(row)].value
		year = refrence_ws['B' + str(row)].value
		if refrence_ws['C' + str(row)].value is not None:
			try:
				db.create_db(refrence_ws['C' + str(row)].value, sys.argv[2], school, year, 'ED')
			except IndexError:
				continue
			print(school, year, type, 'ED', 'Done')
		if refrence_ws['D' + str(row)].value is not None:
			try:
				db.create_db(refrence_ws['D' + str(row)].value, sys.argv[2], school, year, 'RD')
			except IndexError:
				continue
			print(school, year, type, 'RD', 'Done')
		print(i + 1, '/', max_row, 'row done')
