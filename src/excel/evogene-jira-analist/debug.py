import pandas as pd

def merge_dataframes(df1, df2):
    # Merge the two DataFrames using an outer join to keep all companies
    merged_df = pd.merge(df1, df2, on=['Company Name', 'Team Name'], how='outer', suffixes=('_df1', '_df2'))

    # Fill NaN values with 0 for the 'Time Spent (Days)' columns
    merged_df['Time Spent (Days)_df1'].fillna(0, inplace=True)
    merged_df['Time Spent (Days)_df2'].fillna(0, inplace=True)

    return merged_df

# Sample DataFrames
data1 = {
    'Company Name': ['AgPlenus', 'AgPlenus', 'AgPlenus', 'AgPlenus', 'Biomica'],
    'Team Name': ['Algo', 'Bi', 'Dev', 'Irrelevant', 'Algo'],
    'Time Spent (Days)_df1': [0.0, 6.0625, 0.0, 0.0, 7.75]
}

data2 = {
    'Company Name': ['AgPlenus', 'AgPlenus', 'AgPlenus', 'AgPlenus', 'AgSeed'],
    'Team Name': ['Algo', 'Bi', 'Dev', 'Irrelevant', 'Algo'],
    'Time Spent (Days)_df2': [10.625, 8.125, 0.0, 0.0, 0.5]
}

df1 = pd.DataFrame(data1)
df2 = pd.DataFrame(data2)

# Call the merge_dataframes function
merged_df = merge_dataframes(df1, df2)

# Print the merged DataFrame
print(merged_df)
