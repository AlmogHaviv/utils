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
- `v1-yearly-jira.csv` (dynamic CSV file, updated for each run with yearly work ratio data from Jira).

### 2. Script 2: another_script.py

- **Description**: Generates an Excel file displaying monthly planned, actual, and usage differences for BI, Algo, and Dev teams. Also includes yearly totals.

- **Dependencies**: 
- `worker-names.csv` (static CSV file, updated only when team member names change).
- `budget-naming.xlsx` (static Excel file, updated when budget names change).
- `yearly-company-budget.xlsx` (static Excel file, updated when yearly budgets change).
- `jira-missions-yearly2023.csv` (dynamic CSV file, updated for each run with yearly data from Jira).

## Usage

Provide detailed instructions on how to run each script, including command-line arguments and configuration settings. Include example commands where applicable.

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


