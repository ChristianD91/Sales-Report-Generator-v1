# Sales Report Generator

Created by ChristianD91 @ GitHub.

Description
-----------
Sales Report Generator is a lightweight, portable desktop application that lets you generate Excel, PDF, or CSV reports from your sales data. No installation is required.

Features
--------

- Supports CSV and Excel (.xls, .xlsx) input files.
- Outputs reports in Excel, PDF, or CSV format.
- Select which reports to generate:
  - Top Products.
  - Top Customers.
  - Sales by Country.
  - Sales by Month.
  - Sales by Quarter.
  - Sales by Range (Low, Middle, High).
  - Total Sales Summary.
- Automatically includes charts in Excel reports.
- Remembers your last used file paths and report selections.
- Optional Windows 98 retro theme with Comic Sans and bold colors.
- Fully portable — no installation or admin rights required.

Requirements
----------

1- pandas
2- openpyxl
3- fpdf

Creating .exe file
----------

1. Install PyInstaller: pip install pyinstaller
2. From terminal, run:
   pyinstaller --onefile --icon=sales_report_generator.ico app.py
3. The .exe will appear in the 'dist' folder.

How to Use
----------

1. Extract the application folder to any location (e.g., C:\MyReports\).
2. Double-click SalesReportGenerator.exe to launch the app.
3. Choose your input file (CSV or Excel).
4. Select an output folder.
5. Choose the desired report format (Excel, PDF, or CSV).
6. Select which reports to include.
7. Click "Generate Report".
8. Your report will be saved in the selected output folder.

Included Files
--------------

- SalesReportGenerator.exe
- README.txt

Note: settings.json will be created automatically after your first report is generated, retaining your previously selected file paths and the report generation options selected previously.

Portability
-----------

- No installation required.
- Self-contained executable.
- Compatible with Windows 10 and 11 (64-bit).

Themes
------

You can switch between:

- Default: Standard system appearance.
- Windows 98: Retro theme with Comic Sans and colorful interface.

Troubleshooting
---------------

- Ensure your input file includes expected columns like SALES, PRODUCTLINE, CUSTOMERNAME, ORDERDATE, etc.
- If your file uses different column names, adjust them before using the app.

Support
-------

For feedback or questions, ChristianD91 @ GitHub.
