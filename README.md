# **Sales Report Automation (RPA)**

This project automates the generation and distribution of sales reports using Robotic Process Automation (RPA). The goal is to extract data from a PDF file, validate the information based on business rules, enrich the data with currency conversion and address lookups, and finally, generate password-protected PDF sales reports for each vendor. The process culminates in sending an email with all generated reports and a detailed error log.

-----

## **Features**

  * **PDF Data Extraction:** The process reads and extracts tabular data from the `Sales List.pdf` file using the **Camelot** library.
  * **Data Validation:**
      * Verifies that the `VENDOR ID` matches the `AA000000` format.
      * Checks if the invoice date is between 2018 and 2021.
      * Ensures that the `Invoice Number` contains only numbers.
      * Filters out vendors with a "Pending Registration" status.
  * **Data Enrichment:**
      * **Currency Conversion:** Converts `UNIT COST` and `TOTAL PRICE` from USD to BRL using a public API (`http://economia.awesomeapi.com.br`).
      * **Address Lookup:** Finds the full address, neighborhood, and state based on the provided postal code, using the `https://viacep.com.br` API.
  * **Dynamic Report Generation:**
      * Creates an individual Excel report for each valid vendor, using the `Sales Report_Template.xlsx` template.
      * Applies a **10% discount** if the total sales value exceeds 2 million BRL.
      * Calculates and applies taxes based on the vendor's state (São Paulo: 5%, Rio de Janeiro: 2%, Minas Gerais: 1%).
      * **Template Adjustment:** Uses embedded VBA code (`win32com.client`) to adjust the spreadsheet's formatting before exporting to PDF, ensuring the layout fits perfectly.
  * **Secure Output:** Converts the Excel reports into password-protected PDF files using **PyPDF2**. The password is generated from the last six digits of the Vendor ID.
  * **Email Communication:** Sends an email containing all generated sales reports and a detailed error log, using the **smtplib** library for sending emails.

-----

## **Technologies Used**

  * **Python:** The core language for the automation script.
  * **Pandas:** For data manipulation and processing.
  * **Camelot:** Table extraction from PDF files.
  * **Openpyxl:** Reading and writing data in Excel templates.
  * **win32com.client:** For Excel automation, enabling VBA macro execution and PDF export.
  * **PyPDF2:** For manipulating and password-protecting PDF files.
  * **Requests:** For interacting with external currency conversion and postal code APIs.
  * **smtplib:** For secure email sending.
  * **dotenv:** For securely managing environment variables (e.g., email credentials).

-----

## **How to Run the Project**

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/MikaelLopesDev/SalesReportAutomation.git
    cd SalesReportAutomation
    ```
2.  **Configure environment variables:** Create a `.env` file in the root folder and add your email password:
    ```
    EMAIL_PASSWORD="your_email_password"
    ```
3.  **Adjust the configuration:**
      * Edit the `Data/config.json` file to update file paths, email settings, and other business rules.
      * Ensure that the `Sales List.pdf` and `Vendor List.xlsx` files are in the correct folders, as configured.
4.  **Execute the main script:**
    ```bash
    python main.py
    ```

-----

## **Project Structure**

```
.
├── .idea/
├── Data/
│   ├── config.json               # Configuration file for variables
│   ├── input/
│   │   ├── Sales Report_Template.xlsx  # Sales report template
│   │   ├── Vendor List.xlsx            # Vendor list for validation
│   │   └── Sales List.pdf              # Input file with the sales list
│   └── output/
│       ├── reports/
│       └── errors_found.xlsx
├── .env                            # Secure environment variables (e.g., email password)
├── .gitignore                      # Files and folders to be ignored by Git
├── PDD RPA Case - Sales Report.docx # PDD (Process Design Document) documentation
├── main.py                         # The main automation script
├── sendEmails.py                   # Module for sending emails
└── README.md                       # The project documentation (this file)
```

-----

## **Author**

  * [Mikael Lopes](https://www.google.com/search?q=https://github.com/MikaelLopesDev)
