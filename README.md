📊 Monthly Partner Report Generator
Description
The Monthly Partner Report Generator is an Excel macro that automates the creation of a "SOCIOS" sheet based on the existing "FLUJO" sheet. The macro allows you to select a specific month and generate a weekly breakdown of financial data while maintaining the original format of the "FLUJO" sheet. The main motivation for this project is to facilitate quick and accurate financial data analysis, reducing manual errors and saving time in report generation.

The primary technology used is VBA (Visual Basic for Applications), as it allows cell manipulation and sheet handling within Excel. A key challenge in this project was coding the dynamic generation of weeks. This was particularly difficult because the number of weeks varies by month, requiring flexible column range adjustments.

Table of Contents
Installation Instructions
Usage
Credits
License
Installation Instructions
Download the Excel file that contains the macro.
Open the file in Microsoft Excel.
Enable editing and macro execution to allow the macro to run properly:
Go to File > Options > Trust Center > Trust Center Settings > Macro Settings.
Select Enable all macros.
Access the VBA Editor by pressing ALT + F11 and paste the code into a module.
Usage
Open the Excel file with the macro.
Press ALT + F8 and select "Reporte_Socios_Mes" to run the macro.
A window will appear prompting you to enter the name of the month (e.g., "ENERO", "FEBRERO").
The macro will generate a new sheet named "SOCIOS" with the financial report corresponding to the selected month.
The generated sheet will include the breakdown of weeks, bolding for key rows, and borders based on the original format.
Review the generated "SOCIOS" sheet in Excel to ensure the data is accurate.
Credits
Carlos - Development and testing of the macro.
VBA Documentation - Use of references and functions for project implementation.
Microsoft Excel - Platform used for macro development and execution.
License
This project is licensed under the MIT License. This means you can use, modify, and distribute the code, provided the original license and credits are included. No warranties are associated with the use of this project.










