import os
import sys

import pandas as pd
import xlsxwriter

if len(sys.argv) < 2:
    print("Please provide the report file path as an argument.")
    sys.exit(1)
data_file = sys.argv[1]
data_sheet = 'Detail Table'
result_file = os.path.join(os.path.dirname(__file__), 'result.xlsx')
result_sheet = 'Details by Date'
year = 2023


def export(dataframe, export_file, export_sheet):
    workbook = xlsxwriter.Workbook(export_file)
    sheet = workbook.add_worksheet(export_sheet)
    sheet.set_row(0, 25)
    sheet.set_column("A:C", 15)
    sheet.set_column("B:B", 8)
    sheet.set_column("D:E", 45)
    fmt_header = workbook.add_format(
        {'bold': True, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E5E3'})
    fmt_data = workbook.add_format({'border': 1})
    fmt_time = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
    sheet.write(0, 0, "Date", fmt_header)
    sheet.write(0, 1, "Hours", fmt_header)
    sheet.write(0, 2, "Resource Name", fmt_header)
    sheet.write(0, 3, "Project Description", fmt_header)
    sheet.write(0, 4, "Work Description", fmt_header)

    row = 1
    for i, row_data in dataframe.iterrows():
        sheet.write(row, 0, row_data['Date'], fmt_time)
        sheet.write(row, 1, row_data['Hours'], fmt_data)
        sheet.write(row, 2, row_data['Resource Name'], fmt_data)
        sheet.write(row, 3, row_data['Project Description'], fmt_data)
        sheet.write(row, 4, row_data['Work Description'], fmt_data)
        row += 1

    workbook.close()


dfs = pd.read_excel(data_file, sheet_name=[data_sheet], skiprows=2)
df = dfs[data_sheet]
total_hours_idx = df.columns.get_loc('Total (Hours)')
timesheet_remarks_idx = df.columns.get_loc('Timesheet Remarks')
df = df.iloc[:, [0, 5, 6] + list(range(total_hours_idx + 1, timesheet_remarks_idx))]
new_cols = {col: pd.to_datetime(f"{year} {col}", format='%Y %a\n%b %d').strftime('%Y-%m-%d') for col in df.columns[3:]}
df = df.rename(columns=new_cols)
melted_df = pd.melt(df, id_vars=['Resource Name', 'Project Description', 'Work Description'], var_name='Date',
                    value_name='Hours')
melted_df = melted_df[melted_df['Hours'] != 0]
melted_df = melted_df[['Date', 'Hours', 'Resource Name', 'Project Description', 'Work Description']]
melted_df['Date'] = pd.to_datetime(melted_df['Date'])

export(melted_df, result_file, result_sheet)
