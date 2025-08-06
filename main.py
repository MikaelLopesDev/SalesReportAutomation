import pandas as pd
import camelot
import numpy as np
from datetime import datetime, timedelta
import requests
from decimal import Decimal, ROUND_HALF_UP
import re
from openpyxl import load_workbook
import shutil
import win32com.client
from dotenv import load_dotenv
import os
from PyPDF2 import PdfReader, PdfWriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def set_pdf_password(input_pdf, output_pdf, password):
    # read original pdf
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    # ADD ALL PAGES TO WRITER
    for page in reader.pages:
        writer.add_page(page)

    # Set PDF with password
    writer.encrypt(user_password=password, owner_password=None,
                   use_128bit=True)

    # SAVE THE NEW PDF PROTECT
    with open(output_pdf, "wb") as f:
        writer.write(f)


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



def enviar_email_com_anexos(
        remetente: str,
        senha: str,
        destinatarios: list,
        assunto: str,
        corpo: str,
        anexos: list,
        servidor_smtp: str = "smtp.gmail.com",
        porta: int = 587
):
    """
    Envia e-mail com múltiplos anexos usando SMTP.

    Args:
        remetente (str): E-mail do remetente
        senha (str): Senha do e-mail ou senha de app
        destinatarios (list): Lista de e-mails dos destinatários
        assunto (str): Assunto do e-mail
        corpo (str): Corpo do e-mail (pode ser HTML)
        anexos (list): Lista de caminhos dos arquivos para anexar
        servidor_smtp (str): Servidor SMTP (padrão: Gmail)
        porta (int): Porta SMTP (padrão: 587 para TLS)
    """

    # Criar mensagem
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = ", ".join(destinatarios)
    msg['Subject'] = assunto

    # Adicionar corpo do e-mail
    msg.attach(MIMEText(corpo, 'plain'))  # Use 'html' para formato HTML

    # Adicionar anexos
    for caminho_anexo in anexos:
        try:
            with open(caminho_anexo, 'rb') as arquivo:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(arquivo.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {caminho_anexo.split("/")[-1]}'
                )
                msg.attach(part)
        except FileNotFoundError:
            print(f"Erro: Arquivo não encontrado - {caminho_anexo}")
            continue

    # Enviar e-mail
    try:
        with smtplib.SMTP(servidor_smtp, porta) as server:
            server.starttls()  # Ativar TLS
            server.login(remetente, senha)
            server.sendmail(remetente, destinatarios, msg.as_string())
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Falha ao enviar e-mail: {str(e)}")


vba_code = """
Sub AdjustSheetToExportAsPDF()
    With ActiveSheet

        .Cells.EntireColumn.AutoFit

        .Columns("B").ColumnWidth = 48
        .Columns("D").ColumnWidth = 18
        .Columns("E").ColumnWidth = 6
        .Columns("G").ColumnWidth = 20

        With .PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
    End With
End Sub
"""

dict_tax_by_state = {
    "São Paulo": "0,05",
    "Rio de Janeiro": "0,02",
    "Minas Gerais": "0,01"


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
    "qtyFirstPostionTemplate" : "E18",
    "termPostionTemplate":"B35",
    "taxRatePosition": "G29"
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

load_dotenv()
password_email = os.getenv("EMAIL_PASSWORD")
list_files_report = []
path_report_template = "Data/Input/Sales Report_Template.xlsx"
df_vendor_list = pd.read_excel("Data/Input/Vendor List.xlsx")

table_sales_list = camelot.read_pdf("Data/Input/Sales List.pdf")
df_sales_list = table_sales_list[0].df
new_columns =  df_sales_list.iloc[0]
df_sales_list = df_sales_list[1:].copy()
df_sales_list.columns = new_columns
date_invoice_30days = datetime.now() + timedelta(days=30)
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
        df_sales_list.at[index, 'Unit Cost (BRL)'] = Decimal(str(row['UNIT COST']).split()[0])

    if convert_total_price == 0:
        df_sales_list.at[index, 'ERRORS_FOUND'] = str(df_sales_list.at[index, 'ERRORS_FOUND']) + "; " + "Erro: Não foi possivel converter a moeda para BRL"
    else:
        df_sales_list.at[index, 'Total Price (BRL)'] = Decimal(str(row['TOTAL PRICE']).split()[0])

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
    output_vendor_id_report = f"Data/Output/{vendor_id}.xlsx"
    shutil.copy(path_report_template, output_vendor_id_report)
    sales_report_vendor_id = load_workbook(output_vendor_id_report)

    sales_report_vendor_id_sheet = sales_report_vendor_id.active  # Pega a planilha ativa
    df_report_by_vendor_id = df_sales_list_clean[
        df_sales_list_clean['VENDOR ID'] == vendor_id].reset_index(drop=True)
    total_sum_sales = df_report_by_vendor_id['Total Price (BRL)'].sum()
    discount = float(total_sum_sales) * 0.10 if total_sum_sales > 2000000 else  0.0
    tax = dict_tax_by_state[str(df_report_by_vendor_id['Location/State'].iloc[0])] #should added handling possible input erro


    # Start fill cell in invoice sheet
    sales_report_vendor_id_sheet[dict_position_cell["idVendorPositionTemplate"]] = str(df_report_by_vendor_id['VENDOR ID'].iloc[0])
    sales_report_vendor_id_sheet[dict_position_cell["dateTodayPostionTemplate"]] = datetime.now().strftime("%d/%m/%Y")
    sales_report_vendor_id_sheet[dict_position_cell["vendorNamePositionTemplate"]] = "Mikael Lopes"
    sales_report_vendor_id_sheet[dict_position_cell["streetPositionTemplate"]] = str(df_report_by_vendor_id['Address'].iloc[0])
    sales_report_vendor_id_sheet[dict_position_cell["districtCityStatePositionTemplate"]] = str(df_report_by_vendor_id['Location/State'].iloc[0])
    sales_report_vendor_id_sheet[dict_position_cell["phoneNumberPostionTemplate"]] = "86994xxxx92"
    sales_report_vendor_id_sheet[dict_position_cell["emailPostionTemplate"]] = "mikaelslopesit@gmail.com"
    sales_report_vendor_id_sheet[dict_position_cell["discountPostionTemplate"]] =  discount
    sales_report_vendor_id_sheet[dict_position_cell["termPostionTemplate"]] = f"Please, Generate Invoices by {date_invoice_30days.strftime('%d/%m/%Y')}"
    sales_report_vendor_id_sheet[dict_position_cell["taxRatePosition"]] = tax


    start_index_row_costbrl_and_quant  = int(dict_position_cell["unitCostBrlFirstPostionTemplate"][1:])
    start_colum_unit_costbrl = str(dict_position_cell["unitCostBrlFirstPostionTemplate"])[0]
    start_colum_quant = str(dict_position_cell["qtyFirstPostionTemplate"])[0]
    for index, row in df_report_by_vendor_id.iterrows():
        sales_report_vendor_id_sheet[f"{start_colum_unit_costbrl}{start_index_row_costbrl_and_quant+index}"] = round(float(row['Unit Cost (BRL)']),2)
        sales_report_vendor_id_sheet[f"{start_colum_quant}{start_index_row_costbrl_and_quant+index}"] = round(float(row['QTY']),2)

    sales_report_vendor_id.save(output_vendor_id_report)

    path_pdf_without_password = output_vendor_id_report.replace(".xlsx","temp.pdf")
    path_pdf_with_password = output_vendor_id_report.replace(".xlsx", ".pdf")
    #Create PDF File
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False  # Execute in second plan
    workbook = excel.Workbooks.Open(os.path.abspath(output_vendor_id_report))
    vba_module = workbook.VBProject.VBComponents.Add(1)  # 1 = Módulo padrão
    vba_module.CodeModule.AddFromString(vba_code)
    excel.Application.Run("AdjustSheetToExportAsPDF")
    workbook.ExportAsFixedFormat(0, os.path.abspath(path_pdf_without_password))  # 0 = PDF
    workbook.Close(SaveChanges=True)
    excel.Quit()
    set_pdf_password(path_pdf_without_password,path_pdf_with_password,re.sub(r'\D', '', vendor_id))
    list_files_report.append(path_pdf_with_password)


df_sales_list.to_excel("Data/Output/Sales List.xlsx", index=False)

df_valid_errors = df_sales_list[
    df_sales_list['ERRORS_FOUND'].notna() &
    (df_sales_list['ERRORS_FOUND'].str.strip() != '')
]


df_valid_errors.to_excel("Data/Output/Errors Found.xlsx", index=False)
list_files_report.append("Data/Output/Errors Found.xlsx")


enviar_email_com_anexos(
        remetente="mikaellopes777@gmail.com",
        senha=password_email,  # Use senha de app se tiver 2FA ativado
        destinatarios=["mikaelslopesit@gmail.com"],
        assunto="Relatório Mensal",
        corpo="Follow the invoices and report in the attachments.",
        anexos= list_files_report,
        servidor_smtp="smtp.gmail.com",
        porta=587
    )

print(list_files_report)


