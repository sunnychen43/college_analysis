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

headers = [
    'ID:',
    'College:',
    'Year:',
    'Type:',
    'Decision:',
    'SAT I:',
    'Total Score',
    'Reading:',
    'Math:',
    'Writing:',
    'ACT:',
    'ACT Composite:',
    'SAT II:',
    'GPA:',
    'Rank:',
    'AP:',
    'IB:',
    'Senior Year Course Load:',
    'Major Awards:',
    'Extracirriculars:',
    'Work Experience:',
    'Voulunteer/Community Service:',
    'Summer Activities:',
    'Essays:',
    'Recommendations:',
    'Teacher Rec #1:',
    'Teacher Rec #2:',
    'Counselor Rec:',
    'Additional Rec:',
    'Interview::',
    'Financial Aid:',
    'Intended Major:',
    'State:',
    'Country:',
    'School Type:',
    'Ethnicity:',
    'Gender:',
    'Income Bracket:',
    'Hooks:',
    'Strengths:',
    'Weaknesses:',
    'Why?',
    'Where else?',
    'Accepted',
    'Waitlisted',
    'Deferred',
    'Rejected']
