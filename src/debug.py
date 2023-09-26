import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill

# Your DataFrame
data = {
    'Issue id': [49404, 50600, 48700, 48407, 49705, 46303],
    'Issue_key': ['LBI-3', 'INTBIOM-7', 'CPB-986', 'CPB-979', 'MIC-905', 'MIC-822'],
    'Company Name': ['CPB', 'Biomica', 'CPB', 'CPB', 'CPB', 'MicroBoost'],
    'Team Name': ['Dev', 'Bi', 'Bi', 'Bi', 'Bi', 'Bi'],
    'Assignee': ['nerias', 'iliab', 'iliab', 'iliab', 'iliab', 'iliab'],
    'Time Spent (Days)': [0.5, 0.5, 3.0, 3.0, 0.5, 2.0],
    'Sprint': ['2023-06-01', '2023-06-01', '2023-06-01', '2023-06-01', '2023-06-01', '2023-06-01'],
    'Custom field (Budget)': [
        'P271 - CPB Upkeep Computational (2023)',
        'P255 - IBD (Biomica 2023)',
        'P279 - CPB projects Computational (CPB 2023)',
        'P279 - CPB projects Computational (CPB 2023)',
        'P271 - CPB Upkeep Computational (2023)',
        'P273 - Product- Upkeep MB  (2023)'
    ],
}

df = pd.DataFrame(data)

# Step 1: Sort the DataFrame by 'Custom field (Budget)' and 'Team Name'
df = df.sort_values(by=['Custom field (Budget)', 'Team Name'])

print(df)