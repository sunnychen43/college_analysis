import comment
from openpyxl import Workbook

wb = Workbook()
ws = wb.active

url = 'http://talk.collegeconfidential.com/cornell-university/1940785-cornell-ed-class-of-2021-results.html'
data = comment.format(comment.scrape_all_comments(url)[1])

id = 1
row = 1
for data_point in data:
    id_cell = ws["A"+str(row)]
    id_cell.value = id

    data_cell = ws["B"+str(row)]
    data_cell.value = data_point
    row += 1

wb.save("test.xlsx")

