from lxml import html
from lxml.etree import _ElementUnicodeResult
import re
import requests


# Returns list of comments in html structure
def scrape_all_comments(url):
    page = requests.get(url)
    tree = html.fromstring(page.content)
    comments = tree.xpath('//*[@id="Content"]/div[5]/div[1]/ul/li')
    return comments


# Formats and converts html comment to str[]
def format(comment):
    body = comment.xpath('div/div[3]/div/div[1]/node()')  # Gets body of comment
    body = remove_junk(flatten([to_string(e) for e in body]))  # Converts body into a list of strings per element
    for line in body[:5]:  # Checks if "decision:" is in the first 5 lines case-insensitive
        if "decision:" in line.lower():
            return body
    return False


# Removes junk text in a list of strings
def remove_junk(list):
    new_list = []
    for e in list:
        e = re.sub('\n *', '', e)  # Remove all line breaks and extra spaces
        e = re.sub('\[(.*?)\]', '', e)  # Remove all [tags]
        if e == "":  # If element is blank, skip adding it to the output list
            continue
        new_list.append(e)  # Add modified element to new list
    return new_list  # Returns new body without junk text


# Converts single html element to string
def to_string(element):
    if isinstance(element, html.HtmlElement):
        if element.tag == 'ul':  # If element is list of bullet points
            separated_nodes = element.xpath("li/text()")
            return [str(node) for node in separated_nodes]  # Returns string of each node/bullet point
        else:
            return str(element.text_content())
    elif isinstance(element, _ElementUnicodeResult):  # If element is raw text
        return str(element)


# Flattens list
def flatten(l):
    new_list = []
    for e in l:
        if hasattr(e, "__iter__") and not isinstance(e, str):  # If e is a list, add its elements to list
            for se in e:
                new_list.append(se)
        else:  # If e is a string, add e to list
            new_list.append(e)
    return new_list
