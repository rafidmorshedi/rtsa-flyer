import pandas as pd
from mailmerge import MailMerge
import datetime
from pprint import pprint

talks_df = pd.read_excel('data/RTSA NSW Technical Presentation Proposal (Responses).xlsx')

pprint(talks_df['Title of talk'])

talk_num = int(input('Talk index:'))

date_entry = input('Enter a date in YYYY-MM-DD format: ')
year, month, day = map(int, date_entry.split('-'))
talk_date = datetime.date(year, month, day)
date_str = talk_date.strftime("%A %d %B %Y")

talk_details = talks_df.to_dict(orient='records')[talk_num]

def clean_name(name):
    name = name.upper()
    name = name.replace(" ","_")
    return name

template = 'RTSA-meeting-flyer-template.docx'
document = MailMerge(template)
document.merge(
    Presenter_background = talk_details['Short Speaker Biography'],
    Description = talk_details['Talk summary (keep this brief)'],
    date = date_str,
    Organisation = talk_details['Organisation'],
    presentation_title = talk_details['Title of talk'],
    presenter_name = talk_details['Name of speaker'],
    date_short = date_str)

output_path = 'outputs/{}-{}-flyer-DRAFT'.format(
    talk_date.strftime("%Y-%m-%d"),
    clean_name(talk_details['Name of speaker']))
document.write('{}.docx'.format(output_path))
print('Word document exported.')
