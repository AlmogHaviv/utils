import os

import pandas as pd
  
# gives the path of demo.py
path = os.path.realpath(__file__)
  
# gives the directory where demo.py 
# exists
dir = os.path.dirname(path)
  
# replaces folder name of Sibling_1 to 
# Sibling_2 in directory
dir = dir.replace('src', 'data')
  
# changes the current directory to 
# Sibling_2 folder
os.chdir(dir)

df = pd.read_excel("jira-missions-yearly2023.xlsx")

print(df)