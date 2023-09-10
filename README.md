# utils
this repo will hold veriaty of scripts that will hold useful scripts and more that will provide help for my academic work and job
This repo works with python venv, needs to be activate.

In order to enter the venv:
1. in order to activete the venv entr the terminal and activate it using: myenv\Scripts\activate (in windows)
2. in order to have all the packeges that are needed to run this script write in the terminal : pip install -r requirements.txt
3. in order to leave the venv, enter the terminal and deactivate it using: deactivate

Current scripts:
1. work_ratio.py - reads the excel file extracted from jira and returnts pie charts in html of the work ratio divided by companies, teams and products
2. mission_planning.py - reads the excel file extracted from jira and returnts an excel file that contains all the data about the expected work donr by each company and how much work was dine in reality.
3. scraping_txt.py - reads a pdf and with regex it returnts the book names, its code and its description in threee excel columns.
