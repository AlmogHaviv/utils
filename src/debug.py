import calendar

import pandas as pd

import budget_report


def reformat_string(input_string):
    parts = input_string.split('-')
    if len(parts) >= 2:
        return f'P{parts[-2].strip()} - {"".join(parts[-1:]).strip()}'
    return input_string

def reformating_names(dict_of_dfs, planned):
    new_dict = {}
    for key, df in dict_of_dfs.items():
        for index, row in df.iterrows():
            budget = row['Custom field (Budget)']
            
            # Check if the budget column starts with 'P997'
            if budget.startswith('P997'):
                # Reformat the budget column using the reformat_string function
                df.at[index, 'Custom field (Budget)'] = reformat_string(budget)
        
        # Extract the 'Budget' column from 'Custom field (Budget)'
        df['Budget'] = df['Custom field (Budget)'].str.extract(r'(\w+)')[0]

        # Merge the dataframes based on 'Budget' and 'Team Name'
        merged_df = pd.merge(planned, df, on=['Budget', 'Team Name'], how='left')
        
        # Fill NaN values in 'Total Time Spent' with 0
        merged_df['Total Time Spent'].fillna(0, inplace=True)
        
        # Drop rows where 'Company Name' is NaN
        merged_df = merged_df.dropna(subset=['Company Name'])
        
        # Drop the 'Custom field (Budget)' column
        merged_df = merged_df.drop(['Custom field (Budget)'], axis=1)
        
        # Sort the dataframe
        merged_df.sort_values(by=['Company Name', 'Team Name'], inplace=True)

        # Group by 'Budget' and 'Team Name', then sum the columns
        summed_df = merged_df.groupby(['Budget', 'Team Name'], as_index=False).agg({
            'Company Name': 'first',
            'Time planned': 'sum',
            'Total Time Spent': 'sum'
        })

        # Rearrange the columns
        summed_df = summed_df[['Budget', 'Company Name', 'Team Name', 'Time planned', 'Total Time Spent']]

        # Calculate 'Delta'
        summed_df['Delta'] = summed_df['Time planned'] - summed_df['Total Time Spent']


        new_dict[key] = summed_df 

    return new_dict



jira_data_filename = 'jira-missions-yearly.xlsx'
planning_data_filename = 'yearly-company-budget.xlsx'
planning_budget_filename = 'CPBC.xlsx'
output_filename = 'budget-and-time-spent-report.xlsx'
monthly_dict = budget_report.process_jira_data(jira_data_filename)
planned = budget_report.create_planned_column(planning_budget_filename)
print(planned)
new_dict = reformating_names(monthly_dict, planned)

for key, df in new_dict.items():
    print(df.round(2))



### to take to main ###

def reformat_string(input_string):
    parts = input_string.split('-')
    if len(parts) >= 2:
        return f'P{parts[-2].strip()} - {"".join(parts[-1:]).strip()}'
    return input_string

def reformating_names(dict_of_dfs):
    new_dict = {}
    for key, df in dict_of_dfs.items():
        for index, row in df.iterrows():
            budget = row['Custom field (Budget)']
            
            # Check if the budget column starts with 'P997'
            if budget.startswith('P997'):
                # Reformat the budget column using the reformat_string function
                df.at[index, 'Custom field (Budget)'] = reformat_string(budget)
        df['Budget'] = df['Custom field (Budget)'].str.extract(r'(\w+)')[0]

        # Merge the dataframes based on 'Budget' and 'Team Name'
        merged_df = pd.merge(planned, df, on=['Budget', 'Team Name'], how='left')
        merged_df['Total Time Spent'].fillna(0, inplace=True)
        merged_df = merged_df.dropna(subset=['Company Name'])
        merged_df = merged_df.drop(['Custom field (Budget)'], axis=1)
        merged_df = merged_df[['Budget', 'Company Name', 'Team Name',  'Time planned', 'Total Time Spent']]
        merged_df['Delta'] = merged_df['Time planned'] - merged_df['Total Time Spent']

        new_dict[key] = merged_df 

    return new_dict 
