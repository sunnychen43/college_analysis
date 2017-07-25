regex_dict = {
    # Objective
    'Decision:': '^.*?Decision: ?',
    'SAT I:': '^.*?SAT(?! [Il1]{2}).*?: ?',
    'ACT:': '^.*?ACT.*?: ?',
    'SAT II:': '^.*?SAT ?([Il]{2}|2).*?: ?',
    'GPA:': '^.*?GPA.*?: ?',
    'Rank:': '^.*?Rank.*?: ?',
    'AP:': '^.*?AP.*?: ?',
    'IB:': '^.*?IB.*?: ?',
    'Senior Year Course Load:': '^.*?Senior Year Course Load: ?',
    'Major Awards:': '^.*?Major Awards.*?: ?',
    'Gender:': '^.*?(Gender|Sex).*?: ?',

    # Subjective
    'Extracirriculars:': '^.*?(Extracurriculars|EC).*?: ?',
    'Work Experience:': '^.*?Job/Work Experience: ?',
    'Voulunteer/Community Service:': '^.*?Volunteer/Community service: ?',
    'Summer Activities:': '^.*?Summer Activities: ?',
    'Essays:': '^.*?Essays.*?: ?',
    'Recommendations:': '^.*?Recommendation.*?: ?',
    'Teacher Rec #1:': '^.*?Teacher Rec #?1: ?',
    'Teacher Rec #2:': '^.*?Teacher Rec #?2: ?',
    'Counselor Rec:': '^.*?Counselor Rec: ?',
    'Additional Rec:': '^.*?Additional Rec: ?',
    'Interview:': '^.*?Interview: ?',

    # Other
    'Financial Aid:': '^.*?Financial Aid\??: ?',
    'Intended Major:': '^.*?Intended Major: ?',
    'State:': '^.*?State.*?: ?',
    'Country:': '^.*?Country.*?: ?',
    'School Type:': '^.*?School Type: ?',
    'Ethnicity:': '^.*?Ethnicity: ?',
    'Gender:': '^.*?Gender: ?',
    'Income Bracket:': '^.*?Income Bracket: ?',
    'Hooks:': '^.*?Hook.*?: ?',

    # Reflection
    'Strengths:': '^.*?Strengths: ?',
    'Weaknesses:': '^.*?Weaknesses: ?',
    'Why?': '^.*?Why.*?: ?',
    'Where else?': '^.*?Where.*?: ?',
    'Objective:': '^Objective[:-]? ?$',
    'Subjective:': '^Subjective[:-]? ?$',
    'Other:': '^Other[:-]? ?$',
    'Reflection:': '^Reflection[:-]? ?$'}

list_categories = ['Major Awards:', 'Extracirriculars:',
                   'Work Experience:', 'Voulunteer/Community Service:',
                   'Summer Activities:', 'Senior Year Course Load:', 'Essays:',
                   'Recommendations:', 'AP', 'IB']

exclude_regex = ['^Objective[:-]? ?$', '^Subjective[:-]? ?$', '^Other[:-]? ?$', '^Reflection[:-]? ?$']

headers = {
    'ID:': 'A',
    'College:': 'B',
    'Year:': 'C',
    'Type:': 'D',
    'Decision:': 'E',
    'SAT I:': 'F',
    'ACT:': 'G',
    'SAT II:': 'H',
    'GPA:': 'I',
    'Rank:': 'J',
    'AP:': 'K',
    'IB:': 'L',
    'Senior Year Course Load:': 'M',
    'Major Awards:': 'N',
    'Extracirriculars:': 'O',
    'Work Experience:': 'P',
    'Voulunteer/Community Service:': 'Q',
    'Summer Activities:': 'R',
    'Essays:': 'S',
    'Recommendations:': 'T',
    'Teacher Rec #1:': 'U',
    'Teacher Rec #2:': 'V',
    'Counselor Rec::': 'W',
    'Additional Rec::': 'X',
    'Interview::': 'Y',
    'Financial Aid:': 'Z',
    'Intended Major:': 'AA',
    'State:': 'AB',
    'Country:': 'AC',
    'School Type:': 'AD',
    'Ethnicity:': 'AE',
    'Gender:': 'AF',
    'Income Bracket:': 'AG',
    'Hooks:': 'AH',
    'Strengths:': 'AI',
    'Weaknesses:': 'AJ',
    'Why?': 'AK',
    'Where else?': 'AL',
    'Accepted': 'AM',
    'Waitlisted': 'AN',
    'Deferred': 'AO',
    'Rejected': 'AP',
    'Applied': 'AQ'}
