import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Sample DataFrame
data = {
    'Company Name': ['AgPlenus', 'AgPlenus', 'AgPlenus', 'AgSeed', 'AgSeed', 'AgSeed'],
    'Team Name': ['Algo', 'Bi', 'Dev', 'Algo', 'Bi', 'Dev'],
    'January Planned': [2.541667, 2.083333, 0.25, 0.083333, 1.0, 0.0],
    'January Time Spent': [0.0, 0.0, 0.0, 0.0, 0.0, 0.0],
    'June Difference': [-8.083333, -6.041667, 0.25, 0.083333, 0.5, 0.0],
    'July Planned': [2.541667, 2.083333, 0.25, 0.083333, 1.0, 0.0],
    'July Time Spent': [20.0, 5.875, 0.0, 0.0, 0.0, 0.0],
    'July Difference': [-17.458333, -3.791667, 0.25, 0.083333, 1.0, 0.0],
}

df = pd.DataFrame(data)

# Create a new Excel writer object
writer = pd.ExcelWriter('output.xlsx', engine='openpyxl')
df.to_excel(writer, index=False, sheet_name='Sheet1')

# Access the Excel workbook and worksheet
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Iterate through the 'Company Name' column and merge cells
previous_company = None
merge_start_row = 2  # Start from row 2
for row, company in enumerate(df['Company Name'], start=2):
    if company != previous_company:
        if previous_company is not None:
            worksheet.merge_cells(f'A{merge_start_row}:A{row-1}')
        merge_start_row = row
    previous_company = company

# Merge the last group
if previous_company is not None:
    worksheet.merge_cells(f'A{merge_start_row}:A{row}')

# Save the Excel file
writer.close()
