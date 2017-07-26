from openpyxl import load_workbook
from operator import itemgetter
import re


# Returns string of schools in given text
def get_schools(text):
	p = re.compile('(?=(?:,|@|/|-|:|and|to|into|at|by|from|the) ?([A-Z]\w*(?: (?:of )?[A-Z]\w+)?)(?:,| and|.|;|$))')  # regex to match all school names
	matches = p.findall(text)
	final_string = ''
	for match in matches:
		if final_string == '':  # First match
			final_string = match
		else:
			final_string += ', ' + match
	return final_string


# Returns a list of (decision, [schools])
def get_decisions(text):
	if text is None:
		return
	decision_list = []  # List of tuples for decision and schools
	decisions = ['Accepted', 'Rejected', 'Deferred', 'Applied', 'Waitlisted']

	for decision in decisions:  # Finds index of each decision string in text
		if decision == 'Waitlisted':  # Scan for variants of waitlisted
			alternate_forms = ['Waiting', 'Waitlist', 'Waitlisted']
			for alt in alternate_forms:
				index = text.lower().find(alt.lower())
				if index != -1:
					break
		else:
			index = text.lower().find(decision.lower())
		if index is not -1:  # If decision is found in the text
			decision_list.append((decision, index))
	decision_list = sorted(decision_list, key=itemgetter(1))  # Sort by index of decision locations
	if len(decision_list) == 0:  # If no decisions were found, return nothing
		return
	if decision_list[0][1] != 0:  # If first word is not a decision, return nothing
		return

	decision_strings = []
	for i, decision_index in enumerate(decision_list):
		if decision_index[0] == 'Applied':  # Skip applied decisions
			continue
		start_index = decision_index[1]
		if i + 1 == len(decision_list):  # If decision is last in list
			end_index = len(text)
		else:
			end_index = decision_list[i + 1][1]
		decision_strings.append((decision_index[0], get_schools(text[start_index:end_index])))  # Finds list of schools for each decision splice
	return decision_strings
