import openpyxl
from tkinter import *
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from docxtpl import DocxTemplate
from datetime import datetime

# Initialize word document
document_path = Path(__file__).parent / "template_invoice.docx"
doc = DocxTemplate(document_path)

# Function to format date
def format_date(date_value):
    if isinstance(date_value, datetime):
        return date_value.strftime("%m/%d/%Y")
    elif isinstance(date_value, str):
        try:
            return datetime.strptime(date_value, "%m/%d/%Y").strftime("%m/%d/%Y")
        except ValueError:
            try:
                return datetime.strptime(date_value, "%Y-%m-%d %H:%M:%S").strftime("%m/%d/%Y")
            except ValueError:
                return date_value  # Return original value if parsing fails
    return date_value  # Return original value for other types

# Function to format number with commas
def format_number(number):
    try:
        return "{:,}".format(float(number))
    except ValueError:
        return number  # Return original value if formatting fails

# Function to read data from Excel
def read_excel_data(file_path):
    workbook = openpyxl.load_workbook(file_path, data_only=True)  # Open with data_only=True to get calculated values not formulas
    sheet = workbook.active
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Skip the header row
        # Extract values based on new column order
        invoice_date = format_date(row[0]) if len(row) > 0 and row[0] else ""
        invoice_number = row[1] if len(row) > 1 and row[1] else ""
        employee_name = row[2] if len(row) > 2 and row[2] else ""
        service_provided = row[3] if len(row) > 3 and row[3] else ""
        hours_completed = row[4] if len(row) > 4 and row[4] else ""
        hourly_rate = row[5] if len(row) > 5 and row[5] else ""
        total_amount = format_number(row[6]) if len(row) > 6 and row[6] else ""  # Format total amount with commas
        address = row[7] if len(row) > 7 and row[7] else ""

        # Append the collected data to the list
        data.append({
            'invoice_date': invoice_date,
            'invoice_number': invoice_number,
            'employee_name': employee_name,
            'service_provided': service_provided,
            'hours_completed': hours_completed,
            'hourly_rate': hourly_rate,
            'total_amount': total_amount,
            'address': address
        })
    return data

# Function to create invoices
def create_invoices(data):
    for item in data:
        try:
            # Render the document
            doc.render({
                "INVOICE_DATE": item['invoice_date'],
                "INVOICENUMBER": item['invoice_number'],
                "EMPLOYEE": item['employee_name'],
                "SERVICE": item['service_provided'], # For both instances marked as {{SERVICE}}
                "HOURS": item['hours_completed'],
                "RATE": item['hourly_rate'],
                "TOTAL": item['total_amount'],  # For both instances marked as {{TOTAL}}
                "ADDRESS": item['address']  # BILL TO: address
            })

            # Generate a unique filename for each invoice
            save_path = f"Invoice_{item['invoice_number']}.docx"
            
            # Save the document
            doc.save(save_path)
            print(f"File has been saved: {save_path}")
        except Exception as e:
            print(f"An error occurred: {e}")

# Create the main window
root = Tk()
root.title("Excel to Invoice Generator")

# Set background color, text color, and border color
root.configure(bg='black')

# Function to handle file selection and processing
def process_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        data = read_excel_data(file_path)
        create_invoices(data)
        messagebox.showinfo("Success", f"Generated {len(data)} invoices.")

# Create and configure the button
style = ttk.Style()
style.configure('TButton',
    foreground='#6495ED',
    background='#6495ED',  # Cornflower blue
    bordercolor='#6495ED',  # Cornflower blue
    font=('Arial', 14, 'bold')
)

ttk.Button(root, text="Select Excel File and Generate Invoices", command=process_excel, style='TButton').pack(pady=20)

# Start the Tkinter event loop
root.mainloop()
