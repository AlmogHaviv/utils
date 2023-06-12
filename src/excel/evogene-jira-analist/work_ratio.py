import pandas as pd

import matplotlib.pyplot as plt

import numpy as np

### global variabels ###
file_name = "jira-may-5.csv"
column1_name = 'Assignee'
column2_name = 'Work Ratio'
column3_name = 'Custom field (Product)'

### data reciving from csv ### 
# reading csv file and returning the data of the two columns by thier name
def getting_2_columns_from_csv_file(file_path, column1_name, column2_name):
    # teams that are in the data and we do not want to pay attention to
    dev_team = ['TALN', 'liavs', 'maorw', 'morank', 'vladz']
    product_team = ['nira', 'hamutal.e', 'boazm', 'larisar', "elanitec", 'anatol', 'drorf']
    # the sprint we want to check on from jira
    df = pd.read_csv(file_path)
    
    column1_data = df[column1_name].tolist()
    column2_data = df[column2_name].tolist()

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
def getting_3_columns_from_csv_file(file_path, column1_name, column2_name, column3_name):
    # teams that are in the data and we do not want to pay attention to
    dev_team = ['TALN', 'liavs', 'maorw', 'morank', 'vladz']
    product_team = ['nira', 'hamutal.e', 'boazm', 'larisar', "elanitec", 'anatol', 'drorf']
    # the sprint we want to check on from jira
    df = pd.read_csv(file_path)
    
    column1_data = df[column1_name].tolist()
    column2_data = df[column2_name].tolist()
    column3_data = df[column3_name].tolist()

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
    labels = ['over 110%', "110 % - 90 %", "uder 90%", "no time has been logged"]
    sizes = [0, 0, 0, 0]
    for num in dict[name]:
        if type(num) is float: #dealing with NaN
            sizes[3] += 1
        else:
            if int(num[:-1]) > 110 :
                sizes[0] += 1
            elif (int(num[:-1]) <= 110) and (int(num[:-1]) >= 90):
                sizes[1] += 1
            else:
                sizes[2] +=1

    return (name, labels, sizes)


# organizing the data for a full sprint pie chart
def data_for_sprint_visualition(dict, sprint_month):
    labels = ['over 110%', "110 % - 90 %", "uder 90%", "no time has been logged"]
    sizes = [0, 0, 0, 0]
    for key in dict: # taking the given information on every worker and adding it up 
        worker_data = data_for_worker_visualition(dict, key)
        for i in range(4):
            sizes[i] += worker_data[2][i]
    return (f'{sprint_month} sprint ratio', labels, sizes)


# organizing the data for visualizion by team or product
def data_for_team_visualizion(dict, team):
    labels = ['over 110%', "110 % - 90 %", "uder 90%", "no time has been logged"]
    sizes = [0, 0, 0, 0]
    for key in dict: # taking the given information on every worker and adding it up
        if key in team: 
            worker_data = data_for_worker_visualition(dict, key)
            for i in range(4):
                sizes[i] += worker_data[2][i]
    return ('team', labels, sizes)


### pie charts cretors ###
# creating the charts
def pie_chart_data(title, labels, sizes):
    # Create the pie chart
    plt.pie(sizes, labels=labels, autopct='%1.1f%%')
    # Add a title
    plt.title(title)
    # Display the chart
    plt.show()


def pie_chart_data(title, labels, sizes):
    # Create the pie chart
    _, _, text = plt.pie(sizes, labels=labels, autopct=lambda pct: func(pct, sizes), labeldistance=1.05)
    
    # Add a title
    plt.title(title)
    
    # Add labels with percentages and numbers
    plt.setp(text, size=12, weight="bold")

    # Display the chart
    plt.show()


# Custom function to display both percentage and number
def func(pct, allvals):
    absolute = int(pct/100.*np.sum(allvals))
    return "{:.1f}%\n({:d})".format(pct, absolute)


# creates the pie chart by worker name 
def worker_by_name_pie(file_name, column1_name, column2_name):
    a = getting_2_columns_from_csv_file(file_name, column1_name, column2_name)
    for key in a:
        print(key, a[key])
    for key in a:
        worker = data_for_worker_visualition(a, key)
        pie_chart_data(worker[0], worker[1], worker[2])


#  creates pie chart for  full sprint
def pie_chart_by_sprint(sprint_name, column1_name, column2_name, file_name):
        dict_for_ratio = getting_2_columns_from_csv_file(file_name, column1_name, column2_name)
        full_data_for_sprint = data_for_sprint_visualition(dict_for_ratio, sprint_name)
        pie_chart_data(full_data_for_sprint[0], full_data_for_sprint[1], full_data_for_sprint[2])


#creates the pie chart for team
def pie_chart_by_team(column1_name, column2_name, file_name):
        bi = ['gala', 'markb', 'iliab', 'michala', 'liorr']
        algo = [ 'robertoo', 'itair', 'anatm', 'renanam']
        dev = ['duduz', 'nerias', 'noama', 'hodaya']
        teams = [bi, algo, dev]
        teams1 = ['bi', 'algo', 'dev']
        dict_for_ratio = getting_2_columns_from_csv_file(file_name, column1_name, column2_name)
        for i, team in enumerate(teams):
            full_data_for_team = data_for_team_visualizion(dict_for_ratio, team)
            pie_chart_data(f'{teams1[i]} sprint ratio', full_data_for_team[1], full_data_for_team[2])


#creates the pie chart for product
def pie_chart_by_product(column1_name, column2_name, column3_name, file_name):
        dict_for_ratio = getting_3_columns_from_csv_file(file_name, column1_name, column2_name, column3_name)
        for key in dict_for_ratio:
            full_data_for_team = data_for_team_visualizion(dict_for_ratio, [key])
            pie_chart_data(f'{key} sprint ratio', full_data_for_team[1], full_data_for_team[2])



### excecution###

a = getting_3_columns_from_csv_file(file_name, column1_name, column2_name, column3_name)

pie_chart_by_product(column1_name, column2_name, column3_name, file_name)