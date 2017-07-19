from lxml import html
from openpyxl import workbook
from openpyxl import load_workbook
import requests
import re
import comment
import classifier_lists


# Imports comment_set into workbook starting at id and row
def save_comment_set(comment_set, wb, id, row):
    ws = wb.active
    comment_count = 1  # Keeps comment count for progress tracking
    for single_comment in comment_set:  # Loops through all comments in the set
        single_comment = comment.format(single_comment)  # Formats comment

        if not single_comment:  # If comment format is invalid, skip it
            print("Invalid Comment Format")
            print('Save', 'Comment', str(comment_count) + '/' + str(len(comment_set)), "done")
            comment_count += 1
            continue

        for data_point in single_comment:  # Saves each data point in the workbook
            id_cell = ws["A" + str(row)]
            id_cell.value = id

            data_cell = ws["C" + str(row)]
            data_cell.value = data_point
            row += 1

        print('Save', 'Comment', str(comment_count) + '/' + str(len(comment_set)), "done")
        id += 1  # Increment counters
        row += 1
        comment_count += 1
    return wb, id, row  # Return edited workbook, next id, and next row


# Saves an excel file with data from the given link
def save_page(url, file_name):
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
    wb.active['B1'].value = "Class:"
    wb.active['C1'].value = "Data:"

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
        print('Save:', "Page", str(page_id) + '/' + str(num_pages), "done")
    wb.save(file_name)  # Save workbook as file_name


def classify(file_path):
    wb = load_workbook(file_path)
    ws = wb.active
    local_regex_dict = classifier_lists.regex_dict

    max_row = ws.max_row
    for row in range(max_row):
        row += 1
        if row == 1:
            continue

        data_cell = ws['C' + str(row)]
        data = data_cell.value
        if data is None:

            continue
        for key, value in local_regex_dict.items():
            p = re.compile(value)
            m = p.search(data)

            if m:
                ws['B' + str(row)].value = key
                data_cell.value = re.sub(p, '', data)  # Removes identifier from every cell in a category
                break
        print('Classify:', str(row) + '/' + str(max_row), 'done')
    wb.save(file_path)


def collate_lists(wb):

    ws = wb.active
    max_row = ws.max_row
    in_list = False
    list_class = None
    list_data_cell = None
    str_list = []
    for row in range(max_row):
        row += 1
        current_class = ws['B' + str(row)].value
        current_data = ws['C' + str(row)].value

        if in_list:
            if current_class is not None and current_class != list_class:
                for string in str_list:
                    if string is None:
                        continue
                    try:
                        list_data_cell.value += '\n' + string
                    except TypeError:
                        list_data_cell.value = string
                    print(list_data_cell.value)
                str_list = []
                in_list = False
            elif current_class is None:
                str_list.append(str(current_data))

        if not in_list:
            if current_class in classifier_lists.list_categories:
                list_class = current_class
                list_data_cell = ws['C' + str(row)]
                in_list = True
        print('Collate:', str(row) + '/' + str(max_row), 'done')
    return wb


url = 'http://talk.collegeconfidential.com/cornell-university/1835971-cornell-class-of-2020-ed-results-only.html'
fp = 'cornell 2020 ed.xlsx'
save_page(url, fp)
classify(fp)
wb = load_workbook(fp)
collate_lists(wb)
wb.save(fp)