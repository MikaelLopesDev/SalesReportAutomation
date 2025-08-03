import pandas as pd
import camelot
import numpy as np
from datetime import datetime
import requests
from decimal import Decimal, ROUND_HALF_UP
import re
from openpyxl import load_workbook
import os
import shutil


def verify_formate_vendor_id(vendor_id):
    default  = r'^[A-Za-z]{2}\d{6}$'  # 2 letras + 6 números
    return bool(re.fullmatch(default, vendor_id))

def currency_convertion(id_currency):
    url = f"http://economia.awesomeapi.com.br/json/last/{id_currency}-BRL"
    try:
        response = requests.get(url, timeout=10)  # Added timeout
        response.raise_for_status()  # Raises exception for 4XX/5XX errors
        data = response.json()

        # Fixed dictionary access syntax
        currency_key = f"{id_currency}BRL"
        if currency_key in data:
            return Decimal(data[currency_key]['bid'])
        return 0.0
    except (requests.exceptions.RequestException, ValueError, KeyError) as e:
        print(f"Error in currency conversion: {e}")
        return 0.0

def find_address_by_postal_code(postal_code):
    url = f"https://viacep.com.br/ws/{postal_code}/json/"

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json()

        return data
    except (requests.exceptions.RequestException, ValueError, KeyError) as e:
        print(f"Error in finding adress by postal code: {e}")
        return 0


def parse_date(date_str):
    know_formats_dates = [
        "%d/%m/%Y",  # dd/MM/yyyy
        "%d-%b-%y",  # dd-MMM-yy (e.g., "15-Jan-23")
        "%m/%d/%Y",  # MM/dd/yyyy
    ]

    for fmt in know_formats_dates:
        try:
            parsed_date = datetime.strptime(str(date_str), fmt)
            if parsed_date > datetime.strptime(str("01/01/2018"), fmt) and parsed_date < datetime.strptime(
                    str("01/01/2022"), fmt):
                return parsed_date.strftime("%d/%m/%Y")  # Correct: strftime (not strftim)
            else:
                return 1

        except ValueError:
            continue
    return 0

dict_tax_by_state = {
    "São Paulo": "0.05",
    "Rio de Janeiro": "0.02",
    "Minas Gerais": "0.01"


}

dict_position_cell = {
    "idVendorPositionTemplate":"B7",
    "dateTodayPostionTemplate":"C7",
    "vendorNamePositionTemplate":"B9",
    "streetPositionTemplate":"B10",
    "districtCityStatePositionTemplate":"B11",
    "phoneNumberPostionTemplate":"B12",
    "emailPostionTemplate":"B13",
    "discountPostionTemplate":"G28",
    "unitCostBrlFirstPostionTemplate":"D18",
    "qtyPostionTemplate" : "E18",
    "termPostionTemplate":"B35",
}

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


wb = load_workbook("Data/Input/Sales Report_Template.xlsx")
ws = wb.active  # Pega a planilha ativa

path_report_template = "Data/Input/Sales Report_Template.xlsx"
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

    if parse_date(str(row['DATE']).strip()) == 0:
        df_sales_list.at[index, 'ERRORS_FOUND'] = str(
            df_sales_list.at[index, 'ERRORS_FOUND']) + "; " + "Erro: Número da Fatura contém data inválida"
    elif parse_date(str(row['DATE']).strip()) == 1:
        df_sales_list.at[index, 'ERRORS_FOUND'] = str(df_sales_list.at[index, 'ERRORS_FOUND']) + "; " + "Erro: Data fora do intervalo de 2018 a 2021"
    else:
        df_sales_list.at[index, 'DATE'] = parse_date(str(row['DATE']).strip())

    if not verify_formate_vendor_id(str(row['VENDOR ID']).strip()):
        df_sales_list.at[index, 'ERRORS_FOUND'] = str(
            df_sales_list.at[index, 'ERRORS_FOUND']) + "; " + "Erro: VerdorID fora do formato AA000000"


    if df_vendor_list.loc[df_vendor_list["Vendor ID"] == str(row['VENDOR ID']).strip(), "Status"].eq("Pending Registration").any():
        df_sales_list.at[index, 'ERRORS_FOUND'] = str(df_sales_list.at[index, 'ERRORS_FOUND']) + "; " + "Erro: Vendor ID Possui Pending Registration"
    identify_currency_unit_cost  = str(row['UNIT COST']).split()[1]
    identify_currency_total_price = str(row['TOTAL PRICE']).split()[1]
    convert_unit_cost = currency_convertion(identify_currency_unit_cost)
    convert_total_price = currency_convertion(identify_currency_total_price)

    if convert_unit_cost == 0:
        df_sales_list.at[index, 'ERRORS_FOUND'] = str(df_sales_list.at[index, 'ERRORS_FOUND']) + "; " + "Erro: Não foi possivel converter a moeda para BRL"
    else:
        df_sales_list.at[index, 'Unit Cost (BRL)'] = Decimal(str(row['UNIT COST']).split()[0]).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP) * convert_unit_cost

    if convert_total_price == 0:
        df_sales_list.at[index, 'ERRORS_FOUND'] = str(df_sales_list.at[index, 'ERRORS_FOUND']) + "; " + "Erro: Não foi possivel converter a moeda para BRL"
    else:
        df_sales_list.at[index, 'Total Price (BRL)'] = Decimal(str(row['TOTAL PRICE']).split()[0]).quantize(Decimal('0.00'), rounding=ROUND_HALF_UP) * convert_unit_cost

    df_sales_list.at[index, 'Exchange Conversion Date/Time'] = datetime.now().strftime("%d/%m/%Y")

    adresses = find_address_by_postal_code(str(row['POSTAL CODE']).strip())
    if adresses == 0:
        df_sales_list.at[index, 'ERRORS_FOUND'] = str(
            df_sales_list.at[index, 'ERRORS_FOUND']) + "; " + "Erro: Não foi possivel Encontrar endereço válido para o cep"
    else:
        df_sales_list.at[index, 'Address'] = f"{str(adresses['logradouro'])} {str(adresses['complemento'])}"
        df_sales_list.at[index, 'Neighborhood'] = f"{str(adresses['bairro'])}"
        df_sales_list.at[index, 'Location/State'] = f"{str(adresses['estado'])}"


df_sales_list_clean = df_sales_list[
    df_sales_list['ERRORS_FOUND'].isna()]

list_vendor_id = df_sales_list_clean['VENDOR ID'].unique()



for vendor_id in list_vendor_id:
    shutil.copy(path_report_template, f"Data/Output/{vendor_id}.xlsx")
    df_report_by_vendor_id = df_sales_list_clean[
        df_sales_list_clean['VENDOR ID'] == vendor_id]
    total_sum_sales = df_report_by_vendor_id['Total Price (BRL)'].sum()
    discount = 0.10 if total_sum_sales > 2000000 else 0.0
    taxa = dict_tax_by_state[str(df_report_by_vendor_id['Location/State'].iloc[0])] #should added handling possible input erro



df_sales_list.to_excel("Data/Output/Sales List.xlsx", index=False)

valid_errors = df_sales_list[
    df_sales_list['ERRORS_FOUND'].notna() &
    (df_sales_list['ERRORS_FOUND'].str.strip() != '')
]

print(valid_errors)


