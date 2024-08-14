# invoice-generator
Use Now: Download the folder and run the script to try it out! (Run Script, Choose excel data (Invoice_Data.xlsx), and watch the magic happen! 
!Important!: make sure the word document is in the same folder as your python script!

Intro: This project provides a Python-based invoice generator that creates invoices using data from an Excel spreadsheet and a Word template. It utilizes the openpyxl library to read Excel files and docxtpl to fill out the Word template. I created this for payroll based off of data on an excel sheet. Made for a payroll job I had a while back. Removed all confidential information as my actual invoice generator was for a specific use case. This project provides a Python-based invoice generator that creates invoices using data from an Excel spreadsheet and a Word template. It utilizes the openpyxl library to read Excel files and docxtpl to fill out the Word template.

Project Components:
Excel Data: The Excel file should contain the invoice data. The top row is considered a header and is skipped. The data in subsequent rows will be used to generate invoices.
Template Invoice Document: A Word document template with placeholders that will be replaced by data from the Excel file.
Python Code: The script that reads the Excel file, processes the data, and generates the invoices using the Word template.

Features:
Excel Data Handling: Skips the header row and processes all subsequent rows until it finds no more data.
Dynamic Naming: Generated invoices are named after the invoice number.
Customizable Data Handling: Easy to adjust data types and modify the template.

Getting Started:
Setup:
Ensure you have Python installed.
Install the necessary libraries using pip:
pip install openpyxl docxtpl

Project Files:
invoice-generator.py: The main Python script that runs the invoice generation.
template_invoice.docx: The Word document template with placeholders. (This file must be in the same directory as the Python script.)
Invoice_Data.xlsx: Sample Excel data file. You can choose any Excel file; it does not need to be in the same directory as the Python script.

How to Use:
Place your Excel file anywhere on your system.
Ensure that the Word template file (template_invoice.docx) is in the same directory as the Python script.
Update template_invoice.docx with your desired format and placeholders.
Run the Python script:
The script will prompt you to select your Excel file. Once selected, it will generate invoices named after the invoice numbers found in the Excel file.

Testing:
For a quick test, you can place the .py, Word template, and Excel file in the same folder and run the script. The script will use these files directly to generate invoices.

Excel Customization:
Excel File
Column Order: Ensure your columns match the order expected by the script (adjust the order if need be (please see the example excel sheet for an idea):
Invoice Date
Invoice Number
Employee Name
Service Provided
Hours Completed
Hourly Rate
Total Amount
Address (if applicable)
Formatting: The script skips the first row, assuming it contains headers. Ensure the data starts from the second row.

Word Template Customization:
Placeholders: Update the placeholders in the template to match those used in the doc.render method (change placeholder types if need be but change on excel sheet as well:
{{INVOICE_DATE}}
{{INVOICENUMBER}}
{{EMPLOYEE}}
{{SERVICE}}
{{HOURS}}
{{RATE}}
{{TOTAL}}
{{ADDRESS}}
Layout: Customize the layout and design as needed for your invoices.

Important!
Make sure you double check that your excel sheet has the correct data in the correct columns. Additionally,
make sure your word document has the correct placeholders as the data is put in specific places for a reason.

Data Handling: Modify the read_excel_data function if you change the column order or data format in your Excel file.
Error Handling: Add additional error handling as needed for your specific use case.

Example:
Here’s an example of how the generated invoices will be named:

Invoice_12345.docx (where 12345 is the invoice number from the Excel file).












Contact
For questions or contributions, please contact @johnny0182 at johnny.leyva182@gmail.com .
