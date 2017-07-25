from openpyxl import load_workbook
from operator import itemgetter
import re


def get_schools(text):
	p = re.compile('(?=(?:, ?|and |to |into |: ?|at |- ?|by |from |the )([A-Z]\w*(?: (?:of )?[A-Z]\w+)?)(?:,| and|.|;|$))')
	matches = p.findall(text)
	final_string = ''
	for match in matches:
		if final_string == '':
			final_string = match
		else:
			final_string += ', ' + match
	return final_string


def get_decisions(text):
	if text is None:
		return
	decision_dict = []
	decisions = ['Accepted', 'Rejected', 'Deferred', 'Applied', 'Waitlisted']

	for decision in decisions:
		if decision == 'Waitlisted':
			alternate_forms = ['Waiting', 'Waitlist', 'Waitlisted']
			for alt in alternate_forms:
				index = text.lower().find(alt.lower())
				if index != -1:
					break
		else:
			index = text.lower().find(decision.lower())
		if index is not -1:
			decision_dict.append((decision, index))
	decision_dict = sorted(decision_dict, key=itemgetter(1))
	if len(decision_dict) == 0:
		return
	if decision_dict[0][1] != 0:
		return

	decision_strings = []
	for i, decision_index in enumerate(decision_dict):
		if decision_index[0] == 'Applied':
			continue
		start_index = decision_index[1]
		if i + 1 == len(decision_dict):
			end_index = len(text)
		else:
			end_index = decision_dict[i + 1][1]
		decision_strings.append((decision_index[0], get_schools(text[start_index:end_index])))
	return decision_strings

