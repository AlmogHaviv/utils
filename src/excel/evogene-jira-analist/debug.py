import pandas as pd

import mission_planning

def export_dataframe_to_excel(df, output_filename, company_column_name):
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # Initialize some formats to use later.
    merge_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})

    # Iterate through unique Company Names and merge cells.
    unique_companies = df[company_column_name].unique()
    for company in unique_companies:
        start_row = df[df[company_column_name] == company].index[0]
        end_row = df[df[company_column_name] == company].index[-1]
        if start_row != end_row:
            worksheet.merge_range(f'A{start_row+2}:A{end_row+2}', company, merge_format)

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()

df = mission_planning.main()
print(df)
export_dataframe_to_excel(df, 'output.xlsx', 'Company Name')





