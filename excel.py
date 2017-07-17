from lxml import html
from openpyxl import workbook
import requests
import re
import comment


# Takes a comment_set, workbook, and imports the comment_set into the workbook starting at id and row
def save_comment_set(comment_set, wb, id, row):
    ws = wb.active
    comment_count = 1  # Keeps comment count for progress tracking
    for single_comment in comment_set:  # Loops through all comments in the set
        single_comment = comment.format(single_comment)  # Formats comment

        if single_comment == False:  # If comment format is invalid, skip it
            print("Invalid Comment")
            comment_count += 1
            continue

        for data_point in single_comment:  # Saves each data point in the workbook
            id_cell = ws["A"+str(row)]
            id_cell.value = id

            data_cell = ws["B"+str(row)]
            data_cell.value = data_point
            row += 1

        print('Comment', str(comment_count) + '/' + str(len(comment_set)), "done")
        id += 1  # Increment counters
        row += 1
        comment_count += 1
    return wb, id, row  # Return edited workbook, next id, and next row


# Saves an excel file with data from the given link
def save_from_url(url, file_name):
    page = requests.get(url)
    tree = html.fromstring(page.content)

    nav_container = tree.xpath('//*[@id="PagerBefore"]/node()')  # Locates link of last page in thread
    last_page = nav_container[len(nav_container) - 4]
    last_page_url = last_page.attrib['href']

    p = re.compile('(?:-p)(\d+)')  # Extracts the last page number (-p12345)
    m = re.findall(p, last_page_url)
    num_pages = int(m[0])

    wb = workbook.Workbook()
    wb.active['A1'].value = "ID:"  # Create column headers
    wb.active['B1'].value = "Data:"

    id = 1
    row = 2
    for page_id in range(num_pages):  # Loop through all pages in the thread
        page_id += 1
        if page_id == 1:
            new_url = url  # Base url is the link for the first page
        else:
            new_url = url[:-5] + "-p" + str(page_id) + ".html"  # Modifies base url to include page number

        comment_set = comment.scrape_all_comments(new_url)  # Get comment set for current page
        wb, id, row = save_comment_set(comment_set, wb, id, row)  # Save comment set in the workbook
        print("Page", page_id, "done")
    wb.save(file_name)  # Save workbook as file_name


save_from_url('http://talk.collegeconfidential.com/cornell-university/1940785-cornell-ed-class-of-2021-results.html', "cornell.xlsx")






