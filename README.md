Monthly Sales Report Generator

This Python project generates a monthly sales report from Excel files containing site-level production data. The report summarizes quantities, calculates income based on a pricelist, and provides both a summary and daily breakdown for each site.

Features

Reads all production Excel files in the specified folder (*production*site*.xls*).

Reads a pricelist Excel file (pricelist*.xls*) to assign prices.

Generates a consolidated Monthly Sales Report:

Summary per site

Daily breakdown per site

Separate sheets for each site

Automatically formats columns (dates, numeric, income) and adjusts column widths for readability.

GUI interface for easy input of the target period.

Handles missing or empty data gracefully.

Requirements

Python 3.10+

pandas

openpyxl

tkinter (comes with standard Python installation)

Install dependencies via pip if needed:

pip install pandas openpyxl

Usage

Place your production Excel files and pricelist in the same folder.

Run the Python script:

python Monthly_Sales_Report.py


On Windows, if you want to avoid the black console window, rename the script to .pyw and double-click.

The GUI will appear. Enter the period in MMYYYY format (e.g., 022026 for February 2026).

The program will scan the folder, process the files, and generate Monthly_Sales_Report_MMYYYY.xlsx in the same folder.

Notes

If the output file already exists, you will be asked whether to overwrite it.

Quantity and Product columns are required in all source files. The script automatically detects the date column.

Income is calculated as Quantity × Price. Price is taken from the pricelist for the closest date before or equal to the transaction date.

Column widths and formatting are adjusted for readability (currency, dates, quantities).

Example File Naming

Pricelist: pricelist2026.xlsx

Production files: production_site_Szolno.xls, production_site_Mezotur.xlsx, etc.

Future Improvements

Comparison with financial invoiced amounts per site/product/month.

Automatic detection of new columns in pricelist.

Export as .exe for Windows users who prefer a standalone application.



