import os

import pandas as pd

import mission_planning

def main():
    # Input and output file names
    jira_data_filename = 'jira-missions-yearly2023.xlsx'
    planning_data_filename = 'yearly-company-budget.xlsx'
    output_filename = 'budget-and-time-spent-report.xlsx'
    monthly_dict = process_jira_data(jira_data_filename)
    planned_df = mission_planning.transform_yearly_to_monthly_and_divide_by_12(planning_data_filename)
    for month, df in monthly_dict.items():
        new_df = merge_company_and_team_inplace(df)
        monthly_dict[month] = new_df
        print(f'this the data frame of: {month}')
        print(new_df)
    print(planned_df)


# Function to navigate to the right directory
def getting_to_the_right_dir(dir_name):
    # Get the path of the current script
    path = os.path.realpath(__file__)
    # Extract the directory where the data exists
    dir = os.path.dirname(path)
    # Replace folder names to change the current directory
    dir = dir.replace('src', dir_name)  # Adjust 'src' to the actual folder structure
    # Change the current directory to the specified folder
    os.chdir(dir)


# Function to create and process a DataFrame
def process_jira_data(data_file):
    getting_to_the_right_dir('data')

    # Read the data from the given Excel file
    df = mission_planning.read_excel_file(data_file)

    # Modify 'Custom field (Budget)' column to remove parentheses and trailing spaces
    df['Custom field (Budget)'] = df['Custom field (Budget)'].str.replace(r'\([^)]*\)', '', regex=True).str.strip()

    # Group the data by months
    month_dataframes = mission_planning.divide_by_months(df)[0]

    # Initialize an empty dictionary to store the result
    team_monthly_data = {}

    # Iterate through each month's DataFrame
    for month, df in month_dataframes.items():
        # Group the DataFrame by the "Team Name" column
        grouped = df.groupby('Team Name')

        # Create a dictionary to store DataFrames for each team within the month
        team_data = {}

        # Iterate through the groups (teams) within the month
        for team, team_df in grouped:
            # Store each team's DataFrame with the team name as the key
            team_data[team] = team_df

        # Store the team-specific DataFrames for this month in the result dictionary
        team_monthly_data[month] = team_data

    # Create an empty dictionary to store the summed DataFrames
    summed_data = {}

    # Iterate through the team_monthly_data dictionary
    for month, team_data in team_monthly_data.items():
        for team, team_df in team_data.items():
            # Group the mini DataFrame by "Custom field (Budget)" and sum "Time Spent (Days)"
            budget_sums = team_df.groupby(['Custom field (Budget)', 'Company Name'])['Time Spent (Days)'].sum()

            # Convert the result to a DataFrame with appropriate column names
            budget_sums_df = budget_sums.reset_index(name='Total Time Spent')

            # Add a 'Team Name' column with the current team name
            budget_sums_df['Team Name'] = team

            # Sort the DataFrame by 'Company Name'
            budget_sums_df.sort_values(by='Company Name', inplace=True)

            # Store the DataFrame in the 'summed_data' dictionary
            if month not in summed_data:
                summed_data[month] = budget_sums_df
            else:
                # Merge the current DataFrame with the existing one for the same month
                summed_data[month] = pd.concat([summed_data[month], budget_sums_df])

    # Return the final 'summed_data' dictionary
    return summed_data

def merge_company_and_team_inplace(df):
    """
    Merge 'Company Name' and 'Team Name' columns into a single column and sort the DataFrame in place.

    Args:
    df (pd.DataFrame): The input DataFrame with 'Company Name' and 'Team Name' columns.

    Returns:
    resulted_df: the data frame with the new totals rows.
    """
    # Merge 'Company Name' and 'Team Name' columns with a hyphen in between
    df['Company Name'] = df['Company Name'] + ' - ' + df['Team Name']

    # Drop the 'Team Name' column if it's no longer needed
    df.drop(columns=['Team Name'], inplace=True)

    # Sort the DataFrame based on the merged 'Company Name' column
    df.sort_values(by=['Company Name'], inplace=True)
    
    # Group the DataFrame by 'Company Name' and sum the 'Total Time Spent' column
    sum_df = df.groupby('Company Name')['Total Time Spent'].sum().reset_index()

    # Add a new column 'Custom field (Budget)' with 'totals'
    sum_df['Custom field (Budget)'] = 'A - Totals'

    # Concatenate the original DataFrame and the summed DataFrame
    result_df = pd.concat([df, sum_df], ignore_index=True)


    # Sort the DataFrame by 'Company Name'
    result_df.sort_values(by=['Company Name', 'Custom field (Budget)'], inplace=True)

    return result_df





### excecution###
if __name__ == "__main__":
    main()