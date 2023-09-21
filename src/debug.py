import pandas as pd

import mission_planning


df = mission_planning.read_excel_file('jira-missions-yearly2023.xlsx')
print(df)

# Example usage:
# result = process_jira_data('jira-missions-yearly2023.xlsx')
