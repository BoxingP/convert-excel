import datetime
import os

import pandas as pd

data_file = os.path.join(os.path.dirname(__file__), 'data.xlsx')
result_file = os.path.join(os.path.dirname(__file__), 'result.xlsx')
sheet_name = 'Detail Table'
year = 2023

dfs = pd.read_excel(data_file, sheet_name=[sheet_name])
df = dfs[sheet_name]
date_columns = df.columns[3:]
df.rename(columns={col: datetime.datetime.strptime(f"{year} {col}", '%Y %a\n%b %d').strftime('%Y-%m-%d') for col in
                   date_columns}, inplace=True)
melted_df = pd.melt(df, id_vars=['Resource Name', 'Project Description', 'Work Description'], var_name='Date',
                    value_name='Hours')
melted_df = melted_df[melted_df['Hours'] != 0]
melted_df = melted_df[['Date', 'Hours', 'Resource Name', 'Project Description', 'Work Description']]

with pd.ExcelWriter(result_file) as writer:
    melted_df.to_excel(writer, sheet_name=sheet_name, index=False)
