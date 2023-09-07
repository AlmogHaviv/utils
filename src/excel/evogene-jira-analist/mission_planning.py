import pandas as pd


def main():
    filename = 'jira-missions-yearly2023.xlsx'
    df = read_excel_file(filename)
    grouped_data_by_months = divide_by_months(df)
    january_data = dividing_by_team(grouped_data_by_months, 1)
    for group_name, data in january_data.items():
        print(f"January Data for Group '{group_name}':")
        print(data)


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
        company_name = creating_company_column(budget)
       
        # Return the extracted data
        new_df = pd.DataFrame({
            'Time Spent (Days)': time_spent,
            'Sprint': sprint,
            'Assignee': assignee,
            'Custom field (Budget)': budget,
            'company name' : company_name,
        })
        return new_df

    except Exception as e:
        # Return None if there's an error while reading the file
        return None
    

def divide_by_months(df):
    # Convert the 'Date' column to datetime
    df['Date'] = pd.to_datetime(df['Sprint'])
    # Create a new column 'Month' to store the month from the 'Date' column
    df['Month'] = df['Date'].dt.month
    # Assign a custom month value to records from December 2022
    df.loc[((df['Month'] == 12) & (df['Date'].dt.year == 2022)) | ((df['Month'] == 11) & (df['Date'].dt.year == 2022)), 'Month'] = 1  # Change November & December 2022 to January 2023
    # Group the data by 'Month'
    grouped_data = dict(tuple(df.groupby('Month')))    # Perform aggregation or analysis on the grouped data, for example, calculate the mean
    return grouped_data


def dividing_by_team(grouped_data, month_number):
    if month_number in grouped_data:
        # Filter data for the specified month
        month_data = grouped_data[month_number]
        
        # Create a dictionary to store DataFrames for each group
        group_data = {'Bi': None, 'Algo': None, 'Dev': None}
        
        # Define custom grouping based on 'Assignee'
        custom_groups = {
            'markb': 'Bi',
            'iliab': 'Bi',
            'michala': 'Bi',
            'liorr': 'Bi',
            'robertoo': 'Algo',
            'itair': 'Algo',
            'renanam': 'Algo',
            'anatm': 'Algo',
            'duduz': 'Dev',
            'nerias': 'Dev',
            'noama': 'Dev',
            'hodayam': 'Dev',
        }
        
        # Iterate through unique 'Assignee' values and group accordingly
        for assignee, group_name in custom_groups.items():
            assignee_data = month_data[month_data['Assignee'] == assignee]
            if group_data[group_name] is None:
                group_data[group_name] = assignee_data
            else:
                group_data[group_name] = pd.concat([group_data[group_name], assignee_data])
        
        return group_data
    else:
        return None


def creating_company_column(budget_column):
    # creating company column
    df = pd.read_excel("budget-naming.xlsx")

    # Extract the desired columns by their names
    budget_names = df["budget"]
    company_names = df["company name"]

    # Desired new list for dataframe
    res = []

    # conditioning
    for val in budget_column:
        code = val[:4]
        if code[-1] == " ":
            code = code[:3]
        print(code)
        counter = 0
        while budget_names[counter] != code and len(budget_names) > counter:
            counter += 1
        print(counter)
    return res



### excecution###
if __name__ == "__main__":
    read_excel_file('jira-missions-yearly2023.xlsx')
    