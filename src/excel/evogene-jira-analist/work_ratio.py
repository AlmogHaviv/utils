import pandas as pd

import matplotlib.pyplot as plt

import numpy as np


# reading csv file and returning the data of the two columns by thier name
def getting_columns_from_csv_file(file_path, column1_name, column2_name):
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


# taking a value from the dictionery and fix it for use in pie chart
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


# creating the charts
def pie_chart_data(title, labels, sizes):
    # Create the pie chart
    plt.pie(sizes, labels=labels, autopct='%1.1f%%')
    # Add a title
    plt.title(title)
    # Display the chart
    plt.show()


def worker_by_name_pie():
    # Example usage - for workers by name
    file_name = "jira-may2023-2.csv"
    column1_name = 'Assignee'
    column2_name = 'Work Ratio'
    a = getting_columns_from_csv_file(file_name, column1_name, column2_name)
    for key in a:
        print(key, a[key])
    for key in a:
        worker = fix_data_for_visualition(a, key)
        pie_chart_data(worker[0], worker[1], worker[2])


# organizing the data for a full sprint pie chart
def data_for_sprint_visualition(dict, sprint_month):
    labels = ['over 110%', "110 % - 90 %", "uder 90%", "no time has been logged"]
    sizes = [0, 0, 0, 0]
    for key in dict: # taking the given information on every worker and adding it up 
        worker_data = data_for_worker_visualition(dict, key)
        for i in range(4):
            sizes[i] += worker_data[2][i]
    return (f'{sprint_month} sprint ratio', labels, sizes)



def pie_chart_by_sprint(sprint_name):
        column1_name = 'Assignee'
        column2_name = 'Work Ratio'
        file_name = "jira-may2023-2.csv"
        dict_for_ratio = getting_columns_from_csv_file(file_name, column1_name, column2_name)
        full_data_for_sprint = data_for_sprint_visualition(dict_for_ratio, sprint_name)
        pie_chart_data(full_data_for_sprint[0], full_data_for_sprint[1], full_data_for_sprint[2])
        

pie_chart_by_sprint("may")