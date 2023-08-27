import pandas as pd

import matplotlib.pyplot as plt

import plotly.express as px

import numpy as np

from datetime import datetime, timedelta

import argparse


### global variabels ###
file_name = "v1-yearly.csv"
column1_name = 'Assignee'
column2_name = 'Work Ratio'
column3_name = 'Custom field (Product)'
save_path1 = '/graphs'
sprint_name = "July"

### parsers ###
def setup_arg_parser():
    """
    Set up the argument parser.
    """
    parser = argparse.ArgumentParser(description='parse for type of report')
    parser.add_argument('-h', '--help', action='help', default=argparse.SUPPRESS, help='Show this help message and exit.')
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
        pie_chart_by_product(column1_name, column2_name, column3_name, file_name, bol)

    elif args.sprint:
            print("Generating the last month sprint report...")
            pie_chart_by_sprint(sprint_name, column1_name, column2_name, file_name, bol)

    elif args.team:
        print("Generating the teams report...")
        pie_chart_by_team(column1_name, column2_name, file_name, bol)
    else:
        print("please specify which report you want: product, team or sprint")


### data reciving from csv ### 
# reading csv file and returning the data of the two columns by thier name
def getting_2_columns_from_csv_file(file_path, column1_name, column2_name, monthly=True):
    # teams that are in the data and we do not want to pay attention to
    dev_team = ['TALN', 'liavs', 'maorw', 'morank', 'vladz']
    product_team = ['nira', 'hamutal.e', 'boazm', 'larisar', "elanitec", 'anatol', 'drorf']
    # the sprint we want to check on from jira
    df = pd.read_csv(file_path)
    
    column1_data = df[column1_name].tolist()
    column2_data = df[column2_name].tolist()

    # works only if the monthly report is asked
    if monthly:
        current_date_time = datetime.now()
        first_day_of_current_month = current_date_time.replace(day=1)
        last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
        last_month_name = last_day_of_previous_month.strftime("%b")
        if last_day_of_previous_month.year < first_day_of_current_month.year:
            last_year = last_day_of_previous_month.year
        else:
            last_year = first_day_of_current_month.year
        last_year = str(last_day_of_previous_month.year)[2:]
        df["Sprint"] = pd.to_datetime(df["Sprint"], format="%d %b %y")
        months_list = df["Sprint"].dt.strftime("%b").tolist()
        years_list = df["Sprint"].dt.strftime("%y").tolist()
        full_lst =[]
        for i in range(len(column1_data)):
            obj = [[months_list[i],years_list[i]], column1_data[i], column2_data[i]]
            full_lst.append(obj)
        column1_data = []
        column2_data = []
        for lst in full_lst:
            if lst[0][0] == last_month_name and lst[0][1] == last_year:
                column1_data.append(lst[1])
                column2_data.append(lst[2])
        

    integrated_dict = {} # create dictionery to calculate the values
    for i, name  in enumerate(column1_data):
        if (name in dev_team) or (name in product_team): # loose redundent names
            pass
        else:
            if name in integrated_dict:
                integrated_dict[name].append(column2_data[i])
            else:
                integrated_dict[name] = [column2_data[i]]
    
    return integrated_dict



# reciving three columns for data integration
def getting_3_columns_from_csv_file(file_path, column1_name, column2_name, column3_name, monthly=True):
    # teams that are in the data and we do not want to pay attention to
    dev_team = ['TALN', 'liavs', 'maorw', 'morank', 'vladz', 'gala']
    product_team = ['nira', 'hamutal.e', 'boazm', 'larisar', "elanitec", 'anatol', 'drorf']
    # the sprint we want to check on from jira
    df = pd.read_csv(file_path)
    
    column1_data = df[column1_name].tolist()
    column2_data = df[column2_name].tolist()
    column3_data = df[column3_name].tolist()

    
    # works only if the monthly report is asked
    if monthly:
        current_date_time = datetime.now()
        first_day_of_current_month = current_date_time.replace(day=1)
        last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
        last_month_name = last_day_of_previous_month.strftime("%b")
        if last_day_of_previous_month.year < first_day_of_current_month.year:
            last_year = last_day_of_previous_month.year
        else:
            last_year = first_day_of_current_month.year
        last_year = str(last_day_of_previous_month.year)[2:]
        df["Sprint"] = pd.to_datetime(df["Sprint"], format="%d %b %y")
        months_list = df["Sprint"].dt.strftime("%b").tolist()
        years_list = df["Sprint"].dt.strftime("%y").tolist()
        full_lst =[]
        for i in range(len(column1_data)):
            obj = [[months_list[i],years_list[i]], column1_data[i], column2_data[i], column3_data[i]]
            full_lst.append(obj)
        column1_data = []
        column2_data = []
        column3_data = []
        for lst in full_lst:
            if lst[0][0] == last_month_name and lst[0][1] == last_year:
                column1_data.append(lst[1])
                column2_data.append(lst[2])
                column3_data.append(lst[3])

    integrated_dict = {} # create dictionery to calculate the values
    for i, name  in enumerate(column1_data):
        if (name in dev_team) or (name in product_team): # loose redundent names - usage of column 1
            pass
        else:
            if column3_data[i] in integrated_dict: # creating the dict that has the  column 3 as a key
                integrated_dict[column3_data[i]].append(column2_data[i])
            else:
                integrated_dict[column3_data[i]] = [column2_data[i]]
    
    return integrated_dict


### cleaning data and prepring it for pie charts ###
# taking a value from the dictionery and fix it for use in pie chart - suited for the asignee column
def data_for_worker_visualition(dict, name):
    labels = ['over 120%', "120 % - 75 %", "uder 75%", "no time has been logged"]
    sizes = [0, 0, 0, 0]
    for num in dict[name]:
        if type(num) is float: #dealing with NaN
            sizes[3] += 1
        else:
            if int(num[:-1]) > 120 :
                sizes[0] += 1
            elif (int(num[:-1]) <= 120) and (int(num[:-1]) >= 76):
                sizes[1] += 1
            else:
                sizes[2] +=1
    final_sizes = []
    final_lables = []
    for i,num in enumerate(sizes):
        if num != 0:
            final_sizes.append(num)
            final_lables.append(labels[i])        
    return (name, final_lables, final_sizes)


# organizing the data for a full sprint pie chart
def data_for_sprint_visualition(dict, sprint_month):
    labels = ['over 120%', "120 % - 75 %", "uder 75%", "no time has been logged"]
    sizes = [0, 0, 0, 0]
    for key in dict: # taking the given information on every worker and adding it up 
        worker_data = data_for_worker_visualition(dict, key)
        for i, name in enumerate(worker_data[1]):
            if name == 'over 120%':
                sizes[0] += worker_data[2][i]
            elif name == '120 % - 75 %':
                sizes[1] += worker_data[2][i]
            elif name == 'uder 75%':
                sizes[2] += worker_data[2][i]
            else:
                sizes[3] += worker_data[2][i]
    return (f'{sprint_month} Sprint Ratio', labels, sizes)


# organizing the data for visualizion by team or product
def data_for_team_visualizion(dict, team):
    labels = ['over 120%', "120 % - 75 %", "uder 75%", "no time has been logged"]
    sizes = [0, 0, 0, 0]
    for key in dict: # taking the given information on every worker and adding it up
        if key in team: 
            worker_data = data_for_worker_visualition(dict, key)
            for i, name in enumerate(worker_data[1]):
                if name == 'over 120%':
                    sizes[0] += worker_data[2][i]
                elif name == '120 % - 75 %':
                    sizes[1] += worker_data[2][i]
                elif name == 'uder 75%':
                    sizes[2] += worker_data[2][i]
                else:
                    sizes[3] += worker_data[2][i]
    final_sizes = []
    final_lables = []
    for i,num in enumerate(sizes):
        if num != 0:
            final_sizes.append(num)
            final_lables.append(labels[i])
    return ('team', final_lables, final_sizes)


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
    figure.show()


# creates the pie chart by worker name 
def worker_by_name_pie(file_name, column1_name, column2_name, monthly=True):
    a = getting_2_columns_from_csv_file(file_name, column1_name, column2_name, monthly)
    for key in a:
        worker = data_for_worker_visualition(a, key)
        print(worker)
        pie_chart_data(worker[0], worker[1], worker[2])


#  creates pie chart for  full sprint
def pie_chart_by_sprint(sprint_name, column1_name, column2_name, file_name, monthly=True):
        dict_for_ratio = getting_2_columns_from_csv_file(file_name, column1_name, column2_name, monthly)
        full_data_for_sprint = data_for_sprint_visualition(dict_for_ratio, sprint_name)
        pie_chart_data(full_data_for_sprint[0], full_data_for_sprint[1], full_data_for_sprint[2])


#creates the pie chart for team
def pie_chart_by_team(column1_name, column2_name, file_name, monthly=True):
        bi = ['markb', 'iliab', 'michala', 'liorr']
        algo = [ 'robertoo', 'itair', 'renanam', 'anatm']
        dev = ['duduz', 'nerias', 'noama', 'hodayam']
        teams = [bi, algo, dev]
        teams1 = ['bi', 'algo', 'dev']
        dict_for_ratio = getting_2_columns_from_csv_file(file_name, column1_name, column2_name, monthly)
        for i, team in enumerate(teams):
            full_data_for_team = data_for_team_visualizion(dict_for_ratio, team)
            print(f'{teams1[i]} sprint ratio', full_data_for_team[1], full_data_for_team[2])
            pie_chart_data(f'{teams1[i]} sprint ratio', full_data_for_team[1], full_data_for_team[2])


#creates the pie chart for product
def pie_chart_by_product(column1_name, column2_name, column3_name, file_name, monthly=True):
        dict_for_ratio = getting_3_columns_from_csv_file(file_name, column1_name, column2_name, column3_name, monthly)
        for key in dict_for_ratio:
            full_data_for_team = data_for_team_visualizion(dict_for_ratio, [key])
            pie_chart_data(f'{key} sprint ratio', full_data_for_team[1], full_data_for_team[2])


### excecution###
if __name__ == "__main__":
    main()