# utils

This repository contains a variety of scripts designed to provide assistance for academic work and job-related tasks. These scripts are intended to be used with a Python virtual environment (venv).

## Getting Started

To use the scripts in this repository, follow these steps:

1. Activate the virtual environment in your terminal using the command (for Windows):
   myenv\Scripts\activate
2. Install the required packages by running the following command:
   pip install -r requirements.txt
3. To deactivate the virtual environment, use the following command:
   deactivate


## Project Overview

This project includes multiple Python scripts, each serving a specific purpose. Below, you'll find an overview of these scripts and their functionalities.

### 1. Script 1: work_ratio.py

- **Description**: Generates pie charts illustrating work ratios based on various parsers. Outputs an Excel file containing the data used for the charts.

- **Dependencies**:
- `worker-names.csv` (static CSV file, updated only when team member names change).
- `v1-yearly-jira.csv` (dynamic CSV file, updated for each run with yearly work ratio data from Jira)
- To generate v1-yearly-jira.csv, you need to export yearly mission data from Jira. It's crucial to note that all the data should be in English. While exporting, please ensure that the following essential columns are included: 'Assignee', 'Work Ratio', 'Custom field (Product)', 'Issue key', 'Sprint', 'Issue id'. Please be aware that sometimes the 'Sprint' column may display dates in Hebrew.

### 2. Script 2: mission_planning.py

- **Description**: Generates an Excel file displaying monthly planned, actual, and usage differences for BI, Algo, and Dev teams. Also includes yearly totals.

- **Dependencies**: 
- `worker-names.csv` (static CSV file, updated only when team member names change).
- `budget-naming.xlsx` (static Excel file, updated when budget names change).
- `yearly-company-budget.xlsx` (static Excel file, updated when yearly budgets change).
- `jira-missions-yearly2023.xlsx` (dynamic Excel file, updated for each run with yearly data from Jira).
- To generate jira-missions-yearly2023.xlsx, you should export yearly mission data from Jira. It's important to ensure that all data is in English during the export process. Please include the following essential columns: Issue id, Parent id, Status, Due Date, Created, Summary, Original Estimate, Time Spent, Assignee, Custom field (Budget), Custom field (Product), Sprint, Creator, Custom field (Teams), Custom field (Teams), Custom field (Epic Link). Note that occasionally, the 'Sprint' column may display dates in Hebrew.

## Usage

To use Script 1 (`mission_planning.py`), you can specify the type of report you want to generate using command-line arguments. Here are the available options:

- `-m` or `--monthly`: If entered, it will return the last month report.
- `-y` or `--yearly`: If entered, it will return the last year report.
- `-p` or `--product`: If entered, it will return the product report.
- `-s` or `--sprint`: If entered, it will return the last month sprint report.
- `-t` or `--team`: If entered, it will return the teams report.
- '-a' or '--all': If entered, it will return all the monthly\yearly reports.

Example command to generate a monthly report for product team:
```bash
python work_ratio.py -m -p
```
As you can see it takes first the flag of the year or month and then the flag of the team, sprint or product

To use Script 2 (`mission_planning.py`), you can specify the type of report you want to generate and, if needed, a specific company name using command-line arguments. Here are the available options:

- `-f` or `--full_report`: Generates the full yearly report.
- `-y` or `--yearly_report`: Generates the yearly summary report.
- `-m` or `--monthly_report`: Generates the up-to-this-month summary report.
- `-c` or `--company [COMPANY_NAME]`: Generates a report for a specific company. Replace `[COMPANY_NAME]` with the desired company name.

Example command to generate a full yearly report:
```bash
python mission_planning.py -f
```

Example command to generate a report for a specific company (e.g., AgPlenus):

```bash
python mission_planning.py -c AgPlenus
```


## Installation

To set up the environment for these scripts, follow these steps:

1. Create and activate a virtual environment (e.g., myenv).
2. Install the required libraries using `pip` or another package manager. Here's a list of required libraries:
- python>=3.10.7
- pandas
- calendar
- openpyxl
- openpyxl.styles
- matplotlib.pyplot
- plotly.express
- numpy
- argparse
- datetime

## Configuration

Before running the scripts, ensure the following:

- There are no Hebrew words in any of the Excel or CSV files, as pandas uses UTF-8 encoding.


