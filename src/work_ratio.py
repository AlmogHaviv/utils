import pandas as pd

import matplotlib.pyplot as plt

import plotly.express as px

import numpy as np

import argparse

import datetime

import calendar

import os

import traceback

import zipfile

## global variabels ###

file_name = "v1-yearly-jira2.csv"
worker_file = "worker-names.csv"

### parsers ###
def setup_arg_parser():
    """
    Set up the argument parser.
    """
    parser = argparse.ArgumentParser(description='parse for type of report')
    parser.add_argument('-m', '--monthly', type=int, help='Specify a particular month number (1-12)')
    parser.add_argument('-y', '--yearly', action='store_true', help='If entered, return the last year report')
    parser.add_argument('-p', '--product', action='store_true', help='If entered, return the product report')
    parser.add_argument('-s', '--sprint', action='store_true', help='If entered, return the last month sprint report')
    parser.add_argument('-t', '--team', action='store_true', help='If entered, return the teams report')
    parser.add_argument('-a', '--all', action='store_true', help='If entered, return all the report')
    return parser


def main():
    try:
        """
        Main logic of the script using parsed arguments.
        """

        parser = setup_arg_parser()

        # Parse the arguments
        args = parser.parse_args()

        # Check which report type was selected and perform corresponding action
        if args.monthly:
            month_abbr = calendar.month_abbr[int(args.monthly)]
            print(f"Generating the monthly {month_abbr} report...")
            parsing_report(args, True, int(args.monthly))
            zip_filename = 'monthly-report-work-ratio.zip'
            folder_to_zip = '\monthly-report-work-ratio'
            
            
        elif args.yearly:
            print("Generating the last year report...")
            parsing_report(args, False)
            zip_filename = 'yearly-report-work-ratio.zip'
            folder_to_zip = '\yearly-report-work-ratio'

        else:
            print("Please specify the type of report: monthly or yearly")
            getting_to_the_right_dir("reports")
            filename = "exception_report.txt"
            with open(filename, "a") as file:
                file.write("Exception occurred in work_ratio.py:\n")
                file.write("Please specify which report you want: monthly (-m) or yearly (-y)\n")
                file.write("Thank You\n")

            zip_filename = 'exception_report.zip'
            folder_to_zip = ""
        
        zipping_files(zip_filename, folder_to_zip)

    
    except UnicodeDecodeError as e:
        print("An exception occurred:", str(e))
        getting_to_the_right_dir("reports")
        filename = "exception_report.txt"
        with open(filename, "a") as file:
            file.write("Exception occurred in work_ratio.py:\n")
            file.write("You have letters that are not in english in the excel file you have inputed in one of the following columns:\n")
            file.write("Assignee, Work Ratio, Custom field (Product), Issue key, Sprint, Issue id\n")
            file.write("It is probably on Sprint column, please check that the dates are formated like this: 01/01/2023\n")
            file.write("For the developer this is the exception:\n")
            file.write("Thank You\n")

        zip_filename = 'exception_report.zip'
        folder_to_zip = ""
        zipping_files(zip_filename, folder_to_zip)

        print(f"An exception occurred. Details saved in {filename}")
        print(str(e))


    except Exception as e:
        # Handle other exceptions here
        filename = "exception_report.txt"
        with open(filename, "a") as file:
            file.write("Exception occurred in main() function:\n")
            file.write(str(e) + "\n")
            file.write("Stack Trace:\n")
            file.write(traceback.format_exc() + "\n")
        
        zip_filename = 'exception_report.zip'
        folder_to_zip = ""
        zipping_files(zip_filename, folder_to_zip)
        
        print(f"An exception occurred. Details saved in {filename}")
        print(str(e))


def parsing_report(args, bol, month_number):
    if args.product:
        print("Generating the product report...")
        pie_chart_by_product(file_name, worker_file, bol, month_number)

    elif args.sprint:
        print("Generating the last month sprint report...")
        pie_chart_by_sprint(file_name, worker_file, bol, month_number)

    elif args.team:
        print("Generating the teams report...")
        pie_chart_by_team(file_name, worker_file, bol, month_number)

    elif args.all:
        print("Generating the product report...")
        pie_chart_by_product(file_name, worker_file, bol, month_number)
        print("Generating the last month sprint report...")
        pie_chart_by_sprint(file_name, worker_file, bol, month_number)
        print("Generating the teams report...")
        pie_chart_by_team(file_name, worker_file, bol, month_number)

    else:
        print("Please specify which report you want: product, team, or sprint")
        getting_to_the_right_dir("reports")
        filename = "exception_report.txt"
        with open(filename, "a") as file:
            file.write("Exception occurred in work_ratio.py:\n")
            file.write("Please specify which report you want: product (-p), team (-t), or sprint (-s)\n")
            file.write("Thank You\n")


def zipping_files(zip_filename, folder_to_zip):
    # Create a zip file and add all the files and subfolders from the current directory
    getting_to_the_right_dir("reports")

    # Get the current directory
    current_directory = os.getcwd()

    # Define the name of the zip file you want to create
    zip_filename = f'{zip_filename}.zip'

    # which folder to zip
    dit_to_zip = current_directory + f'{folder_to_zip}'


    # Create a zip file and add all the files and subfolders from the current directory
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for foldername, subfoldername, filenames in os.walk(dit_to_zip):
            for filename in filenames:
                # Calculate the file's full path
                file_path = os.path.join(foldername, filename)
                print(file_path)
                # Calculate the path to store the file inside the zip file
                zip_path = os.path.relpath(file_path, dit_to_zip)
                # Add the file to the zip file
                zipf.write(file_path, zip_path)

    print(f'Successfully created {zip_filename}')


### data reciving from csv ###
def getting_to_the_right_dir(dir_name):
    # gives the path of demo.py
    path = os.path.realpath(__file__)
    # gives the directory where the data exists
    dir = os.path.dirname(path)
  
     #replaces folder names 
    dir = dir.replace('src', dir_name)
  
    # changes the current directory to data 
    os.chdir(dir)



# reciving the names of the workers by demand.
def worker_teams(worker_path, column_name):
    getting_to_the_right_dir('config')

    df = pd.read_csv(worker_path)
    return df[column_name].tolist()


def creating_dataframe(csv_file_path, worker_path, monthly, month_number):
    getting_to_the_right_dir('data')
    # List of columns you want to select from the CSV file
    columns_to_select = ['Assignee', 'Work Ratio', 'Custom field (Product)', 'Issue key', 'Sprint', 'Issue id']
    # Read the CSV file and select specific columns
    df = pd.read_csv(csv_file_path, usecols=columns_to_select)
    # List of names to remove
    names_to_remove = worker_teams(worker_path, "Devops") + worker_teams(worker_path, "Product")
    # Filter the DataFrame to exclude rows with names in the 'names_to_remove' list
    df_filtered = df[~df['Assignee'].isin(names_to_remove)]
    # Convert 'Work Ratio' from percentage strings to float values
    df_filtered['Work Ratio'] = pd.to_numeric(df_filtered['Work Ratio'].str.rstrip('%'), errors='coerce') / 100
    # Create a new column 'Number for Ratio' based on 'Work Ratio' values
    conditions = [
        (df_filtered['Work Ratio'].isna()),  # Condition for empty cells
        (df_filtered['Work Ratio'] == 0),  # Condition for Work Ratio equal to 0
        ((df_filtered['Work Ratio'] > 0) & (df_filtered['Work Ratio'] < 0.75)),
        ((df_filtered['Work Ratio'] >= 0.75) & (df_filtered['Work Ratio'] <= 1.20)),
        (df_filtered['Work Ratio'] > 1.20)
    ]
    choices = [0, 0, 1, 2, 3]

    df_filtered['Number for Ratio'] = np.select(conditions, choices, default=-1)

    if monthly:
        current_year = datetime.datetime.now().year
        last_year = current_year - 1       
        df_filtered['Date'] = pd.to_datetime(df_filtered['Sprint'])
        # Create a new column 'Month' to store the month from the 'Date' column
        df_filtered['Month'] = df_filtered['Date'].dt.month
        # Assign a custom month value to records from December 2022
        df_filtered.loc[((df_filtered['Month'] == 12) & (df_filtered['Date'].dt.year == last_year)) | ((df_filtered['Month'] == 11) & (df_filtered['Date'].dt.year == last_year)), 'Month'] = 1  # Change November & December to January
        # Group the data by 'Month'
        grouped_data = dict(tuple(df_filtered.groupby('Month')))    # Perform aggregation or analysis on the grouped data, for example, calculate the mean
        # extracting the neccesery month
        df_filtered = grouped_data[month_number]

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
def pie_chart_data(title, labels, sizes, monthly):
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
    getting_to_the_right_dir('reports')

    
    # Get the current working directory as the main folder
    main_folder = os.getcwd()

    if monthly:
        # Define the name of the subfolder you want to enter
        subfolder_name = "monthly-report-work-ratio"
    else:
        subfolder_name = "yearly-report-work-ratio"

    # Create the path to the subfolder
    subfolder_path = os.path.join(main_folder, subfolder_name)

    # Check if the subfolder exists; if not, create it
    if not os.path.exists(subfolder_path):
        os.makedirs(subfolder_path)

    # Now you can save files or outputs to the subfolder
    output_file_path = os.path.join(subfolder_path, f"{title}.html")

    figure.write_html(output_file_path)
    return subfolder_path


#  creates pie chart for  full sprint
def pie_chart_by_sprint(file_name, worker_path, monthly, month_number):
    df = creating_dataframe(file_name, worker_path, monthly, month_number)
    full_data_for_sprint = data_for_sprint_visualition(df, monthly)
    file_path = pie_chart_data(full_data_for_sprint[0], full_data_for_sprint[1], full_data_for_sprint[2], monthly)
    output_file_path = os.path.join(file_path, f'{full_data_for_sprint[0]}.xlsx')
    filter_and_export_work_ratio(df, output_file_path)
    print(f'excel file: {full_data_for_sprint[0]}.xlsx was created')


#creates the pie chart for team
def pie_chart_by_team(file_name, worker_path, monthly, month_number):
    df = creating_dataframe(file_name, worker_path, monthly, month_number)
    teams = [("Bi", worker_teams(worker_path, "Bi")), ("Algo", worker_teams(worker_path, "Algo")), ("Dev", worker_teams(worker_path, "Dev"))]
    for team_name, team_list in teams:
        filtered_df = df[df['Assignee'].isin(team_list)]
        data_list_for_chart = data_for_team_visualizion(filtered_df, team_name)
        file_path = pie_chart_data(data_list_for_chart[0], data_list_for_chart[1], data_list_for_chart[2], monthly)
        output_file_path = os.path.join(file_path, f'{data_list_for_chart[0]}.xlsx')
        filter_and_export_work_ratio(filtered_df, output_file_path)
        print(f'excel file: {data_list_for_chart[0]}.xlsx was created')



#creates the pie chart for product
def pie_chart_by_product(file_name, worker_path, monthly, month_number):
    df = creating_dataframe(file_name, worker_path, monthly, month_number)
    grouped = df.groupby("Custom field (Product)")
    for product, group_df in grouped:
        data_for_chart = data_for_team_visualizion(group_df, product)
        file_path = pie_chart_data(data_for_chart[0], data_for_chart[1], data_for_chart[2], monthly)
        output_file_path = os.path.join(file_path, f'{data_for_chart[0]}.xlsx')
        filter_and_export_work_ratio(group_df, output_file_path)
        print(f'excel file: {data_for_chart[0]}.xlsx was created')


## excel creator ##
def filter_and_export_work_ratio(dataframe, name):
    # Create a new Excel writer object
    excel_writer = pd.ExcelWriter(f'{name}', engine='openpyxl')

    # Filter the DataFrame based on Work Ratio and export to separate sheets
    dataframe[(dataframe['Work Ratio'] < 0.75) & (dataframe['Work Ratio'] != 0)].to_excel(excel_writer, sheet_name='under 75%', index=False)
    dataframe[(dataframe['Work Ratio'] >= 0.75) & (dataframe['Work Ratio'] <= 1.20)].to_excel(excel_writer, sheet_name='between 75% to 120%', index=False)
    dataframe[dataframe['Work Ratio'] > 1.20].to_excel(excel_writer, sheet_name='above 120%', index=False)
    # Filter and export rows with empty Work Ratio cells
    dataframe[(dataframe['Work Ratio'].isna()) | (dataframe['Work Ratio'] == 0)].to_excel(excel_writer, sheet_name='no time was logged', index=False)

    # Save the Excel file
    excel_writer.close()

### excecution###
if __name__ == "__main__":
   main()
