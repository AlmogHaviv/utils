import pandas as pd

import matplotlib.pyplot as plt

import plotly.express as px

import numpy as np

import argparse

import datetime

import calendar

## global variabels ###

file_name = "v1-yearly-jira.csv"
worker_file = "worker-names.csv"

### parsers ###
def setup_arg_parser():
    """
    Set up the argument parser.
    """
    parser = argparse.ArgumentParser(description='parse for type of report')
    parser.add_argument('-m', '--monthly', action='store_true', help='If entered, return the last month report')
    parser.add_argument('-y', '--yearly', action='store_true', help='If entered, return the last year report')
    parser.add_argument('-p', '--product', action='store_true', help='If entered, return the product report')
    parser.add_argument('-s', '--sprint', action='store_true', help='If entered, return the last month sprint report')
    parser.add_argument('-t', '--team', action='store_true', help='If entered, return the teams report')
    return parser


def main():
    """
    Main logic of the script using parsed arguments.
    """

    parser = setup_arg_parser()

    # Parse the arguments
    args = parser.parse_args()

    # Check which report type was selected and perform corresponding action
    if args.monthly:
        print("Generating the last month report...")
        parsing_report(args, True)
        
    elif args.yearly:
        print("Generating the last year report...")
        parsing_report(args, False)
    else:
        print("please specify the type of report: monthly or yearly")

    
def parsing_report(args, bol):
    if args.product:
        print("Generating the product report...")
        pie_chart_by_product(file_name, worker_file, bol)

    elif args.sprint:
        print("Generating the last month sprint report...")
        pie_chart_by_sprint(file_name, worker_file, bol)

    elif args.team:
        print("Generating the teams report...")
        pie_chart_by_team(file_name, worker_file, bol)
    else:
        print("please specify which report you want: product, team or sprint")


### data reciving from csv ###
# reciving the names of the workers by demand.
def worker_teams(worker_path, column_name):
    df = pd.read_csv(worker_path)
    return df[column_name].tolist()


def creating_dataframe(csv_file_path, worker_path, monthly):
    # List of columns you want to select from the CSV file
    columns_to_select = ['Assignee', 'Work Ratio', 'Custom field (Product)', 'Issue key', 'Sprint', 'Issue id']
    # Read the CSV file and select specific columns
    df = pd.read_csv(csv_file_path, usecols=columns_to_select)
    # List of names to remove
    names_to_remove = worker_teams(worker_path, "Devops") + worker_teams(worker_path, "Product")
    # Filter the DataFrame to exclude rows with names in the 'names_to_remove' list
    df_filtered = df[~df['Assignee'].isin(names_to_remove)]
    # Convert 'Work Ratio' from percentage strings to float values
    df_filtered['Work Ratio'] = df_filtered['Work Ratio'].str.rstrip('%').astype(float) / 100
    # Create a new column 'Number for Ratio' based on 'Work Ratio' values
    conditions = [
        (df_filtered['Work Ratio'].isna()),  # Condition for empty cells
        (df_filtered['Work Ratio'] < 0.75),
        ((df_filtered['Work Ratio'] >= 0.75) & (df_filtered['Work Ratio'] <= 1.20)),
        (df_filtered['Work Ratio'] > 1.20)
    ]
    choices = [0, 1, 2, 3]

    df_filtered['Number for Ratio'] = np.select(conditions, choices, default=-1)

    if monthly:
        current_year = datetime.datetime.now().year
        last_year = current_year - 1       
        df_filtered['Date'] = pd.to_datetime(df_filtered['Sprint'])
        # Create a new column 'Month' to store the month from the 'Date' column
        df_filtered['Month'] = df_filtered['Date'].dt.month
        # Assign a custom month value to records from December 2022
        df_filtered.loc[((df_filtered['Month'] == 12) & (df_filtered['Date'].dt.year == last_year)) | ((df_filtered['Month'] == 11) & (df_filtered['Date'].dt.year == last_year)), 'Month'] = 1  # Change November & December 2022 to January 2023
        # Group the data by 'Month'
        grouped_data = dict(tuple(df_filtered.groupby('Month')))    # Perform aggregation or analysis on the grouped data, for example, calculate the mean
        # Find the number of the last month
        last_month = df_filtered['Month'].max()
        df_filtered = grouped_data[last_month]

    return df_filtered


### cleaning data and prepring it for pie charts ###
def data_for_sprint_visualition(df, monthly):
    labels = ['no time has been logged' ,'uder 75%', '75 % - 120 %','over 120%']
    grouped = df.groupby('Number for Ratio').size().reset_index(name='Count')
    grouped = grouped.sort_values(by='Number for Ratio')
    sizes = grouped['Count'].tolist()
    number_for_ratio = grouped['Number for Ratio'].tolist()
    final_lables = []
    for num in range(0,4):
         if num in number_for_ratio:
            final_lables.append(labels[num])

    if monthly:
        name_at_row = df["Month"].tolist()
        sprint_month =  calendar.month_name[name_at_row[0]]
        return (f'{sprint_month} Sprint Ratio', final_lables, sizes)
    
    return ('yearly Sprint Ratio', final_lables, sizes)


# organizing the data for visualizion by team or product
def data_for_team_visualizion(df, team):
    labels = ['no time has been logged' ,'uder 75%', '75 % - 120 %','over 120%']
    grouped = df.groupby('Number for Ratio').size().reset_index(name='Count')
    grouped = grouped.sort_values(by='Number for Ratio')
    sizes = grouped['Count'].tolist()
    number_for_ratio = grouped['Number for Ratio'].tolist()
    final_lables = []
    for num in range(0,4):
         if num in number_for_ratio:
            final_lables.append(labels[num])
    return (f'{team} Sprint Ratio', final_lables, sizes)


### pie charts cretors ###
def pie_chart_data(title, labels, sizes):
    """
    Generate and display a pie chart using Plotly Express.

    Parameters:
    - title (str): The title of the pie chart.
    - labels (list): List of labels for each segment of the pie chart.
    - sizes (list): List of sizes (values) corresponding to each segment.

    Returns:
    None
    """
    # Create a DataFrame from the labels and sizes
    data = {"Name": labels, "Value": sizes}
    df = pd.DataFrame(data)

    # Create the pie chart using Plotly Express
    figure = px.pie(df, values='Value', names='Name', title=title)
    figure.update_traces(textposition='inside', textinfo='percent+label+value', 
                         textfont_size=20,
                         marker=dict(line=dict(color='#000000', width=2)))

    # Display the pie chart
    figure.write_html( f"{title}.html")


#  creates pie chart for  full sprint
def pie_chart_by_sprint(file_name, worker_path, monthly):
    df = creating_dataframe(file_name, worker_path, monthly)
    full_data_for_sprint = data_for_sprint_visualition(df, monthly)
    df.to_excel(f'{full_data_for_sprint[0]}.xlsx', index=False)
    print(f'excel file: {full_data_for_sprint[0]}.xlsx was created')
    pie_chart_data(full_data_for_sprint[0], full_data_for_sprint[1], full_data_for_sprint[2])


#creates the pie chart for team
def pie_chart_by_team(file_name, worker_path, monthly):
    df = creating_dataframe(file_name, worker_path, monthly)
    teams = [("Bi", worker_teams(worker_path, "Bi")), ("Algo", worker_teams(worker_path, "Algo")), ("Dev", worker_teams(worker_path, "Dev"))]
    for team_name, team_list in teams:
        filtered_df = df[df['Assignee'].isin(team_list)]
        data_list_for_chart = data_for_team_visualizion(filtered_df, team_name)
        pie_chart_data(data_list_for_chart[0], data_list_for_chart[1], data_list_for_chart[2])
        df.to_excel(f'{data_list_for_chart[0]}.xlsx', index=False)
        print(f'excel file: {data_list_for_chart[0]}.xlsx was created')



#creates the pie chart for product
def pie_chart_by_product(file_name, worker_path, monthly):
    df = creating_dataframe(file_name, worker_path, monthly)
    grouped = df.groupby("Custom field (Product)")
    for product, group_df in grouped:
        data_for_chart = data_for_team_visualizion(group_df, product)
        pie_chart_data(data_for_chart[0], data_for_chart[1], data_for_chart[2])
        df.to_excel(f'{data_for_chart[0]}.xlsx', index=False)
        print(f'excel file: {data_for_chart[0]}.xlsx was created')

             


### excecution###
if __name__ == "__main__":
   main()
