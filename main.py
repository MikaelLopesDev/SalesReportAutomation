import pandas as pd
import camelot


table_sales_list = camelot.read_pdf("Data/Input/Sales List.pdf")

df_sales_list = table_sales_list[0].df

new_columns =  df_sales_list.iloc[0]
df_sales_list = df_sales_list[1:].copy()

df_sales_list.columns = new_columns



df_sales_list.to_excel("Data/Output/Sales List.xlsx", index=False)

for row in df_sales_list.itertuples():
    print(f"Essa é a linha {row.INVOICE}")



