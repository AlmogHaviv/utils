import pandas as pd

import mission_planning

df = mission_planning.read_excel_file('jira-missions-yearly2023.xlsx')

 # Group the data by months
month_dataframes = mission_planning.divide_by_months(df)[0]

# Assuming you have a dictionary of DataFrames called month_dataframes
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

# Now, 'summed_data' is a dictionary where each entry corresponds to a month,
# and each value is a DataFrame with the merged, summed, and sorted data for that month

# Now, 'summed_data' is a dictionary where each entry corresponds to a month and team,
# and each value is a DataFrame with the summed data for that combination



for month, df in summed_data.items():
    print(month)
    print(df)

