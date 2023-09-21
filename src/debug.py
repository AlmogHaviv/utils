import pandas as pd

import mission_planning

jira_data_filename = 'jira-missions-yearly.xlsx'

# Read the Excel file into a DataFrame
df = mission_planning.read_excel_file(jira_data_filename)
    
# Group the data by months
grouped_data_by_months = mission_planning.divide_by_months(df)[0]


for month in grouped_data_by_months:
    grouped_data_by_months[month] = grouped_data_by_months[month][grouped_data_by_months[month]['Team Name'] != 'Irrelevant']

