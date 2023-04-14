import os
import sys

import pandas as pd

if len(sys.argv) < 2:
    print("Please provide the report file path as an argument.")
    sys.exit(1)
data_file = sys.argv[1]
result_file = os.path.join(os.path.dirname(__file__), 'result.xlsx')
sheet_name = 'Detail Table'
year = 2023

dfs = pd.read_excel(data_file, sheet_name=[sheet_name], skiprows=2)
df = dfs[sheet_name]
total_hours_idx = df.columns.get_loc('Total (Hours)')
timesheet_remarks_idx = df.columns.get_loc('Timesheet Remarks')
df = df.iloc[:, [0, 5, 6] + list(range(total_hours_idx + 1, timesheet_remarks_idx))]
new_cols = {col: pd.to_datetime(f"{year} {col}", format='%Y %a\n%b %d').strftime('%Y-%m-%d') for col in df.columns[3:]}
df = df.rename(columns=new_cols)
melted_df = pd.melt(df, id_vars=['Resource Name', 'Project Description', 'Work Description'], var_name='Date',
                    value_name='Hours')
melted_df = melted_df[melted_df['Hours'] != 0]
melted_df = melted_df[['Date', 'Hours', 'Resource Name', 'Project Description', 'Work Description']]
with pd.ExcelWriter(result_file) as writer:
    melted_df.to_excel(writer, sheet_name=sheet_name, index=False)
