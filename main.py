import pandas as pd
import camelot
import numpy as np
from datetime import datetime
import requests

# Define the names of the additional columns
columns_to_add = [
    'Unit Cost (BRL)',
    'Total Price (BRL)',
    'Exchange Conversion Date/Time',
    'Address',
    'Neighborhood',
    'Location/State',
    'ERRORS_FOUND'
]


def parse_date(date_str):
    know_formats_dates = [
        "%d/%m/%Y",  # dd/MM/yyyy
        "%d-%b-%y",  # dd-MMM-yy (e.g., "15-Jan-23")
        "%m/%d/%Y",  # MM/dd/yyyy
    ]

    for fmt in know_formats_dates:
        try:
            parsed_date = datetime.strptime(str(date_str), fmt)
            return parsed_date.strftime("%d/%m/%Y")  # Correct: strftime (not strftim)
        except ValueError:
            continue
    return 0

df_vendor_list = pd.read_excel("Data/Input/Vendor List.xlsx")

table_sales_list = camelot.read_pdf("Data/Input/Sales List.pdf")
df_sales_list = table_sales_list[0].df
new_columns =  df_sales_list.iloc[0]
df_sales_list = df_sales_list[1:].copy()
df_sales_list.columns = new_columns
# Add each new column to the DataFrame, initialized with NaN
for col in columns_to_add:
    df_sales_list[col] = np.nan


for index, row in df_sales_list.iterrows():
    if not str(row['INVOICE']).strip().isdigit():
        df_sales_list.at[index, 'ERRORS_FOUND'] = "Erro: Número da Fatura contém letras e deve ser apenas números"

    if parse_date(str(row['DATE']).strip()) != 0:
        df_sales_list.at[index, 'DATE'] = parse_date(str(row['DATE']).strip())
    else:
        df_sales_list.at[index, 'ERRORS_FOUND'] = str(df_sales_list.at[index, 'ERRORS_FOUND']) + "; " + "Erro: Número da Fatura contém data inválida"

    if df_vendor_list.loc[df_vendor_list["Vendor ID"] == str(row['VENDOR ID']).strip(), "Status"].eq("Pending Registration").any():
        df_sales_list.at[index, 'ERRORS_FOUND'] = str(df_sales_list.at[index, 'ERRORS_FOUND']) + "; " + "Erro: Vendor ID Possui Pending Registration"



valid_errors = df_sales_list[
    df_sales_list['ERRORS_FOUND'].notna() &
    (df_sales_list['ERRORS_FOUND'].str.strip() != '')
]

print(valid_errors)


