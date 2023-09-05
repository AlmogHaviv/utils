import pandas as pd


def main():
    filename = 'jira-missions-yearly2023.xlsx'


def read_excel_file(filename):
    try:
        # Attempt to read the Excel file into a DataFrame
        df = pd.read_excel(filename)

        # Convert 'Time Spent' from seconds to days
        df['Time Spent (Days)'] = df['Time Spent'] / 3600 / 8  # Convert seconds to days (8-hour workdays)

        # Extract the desired columns by their names
        # Assuming that the Excel file contains columns named 'Time Spent,' 'Sprint,' and 'Assignee'
        time_spent = df['Time Spent (Days)']    
        sprint = df['Sprint']         
        assignee = df['Assignee']
        budget  = df['Custom field (Budget)']
       
        # Return the extracted data
        return time_spent, sprint, assignee, budget

    except Exception as e:
        # Return None if there's an error while reading the file
        return None


### excecution###
if __name__ == "__main__":
    main()