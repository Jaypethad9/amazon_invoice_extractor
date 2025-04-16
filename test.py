import fitz  # PyMuPDF
import pandas as pd
import re
import calendar
import os
import shutil  # Importing shutil to move files
from datetime import datetime



def extract_data_from_pdf(pdf_path):
    # Initialize the PDF document
    pdf_document = fitz.open(pdf_path)
    
    # Prepare an empty list to store extracted data
    data = []

    # Define regex patterns
    invoice_pattern = r'Invoice Number\s*:\s*([^\n]+?)\s*(?=Order Date\s*:|$)'
    gst_pattern = r'\b\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z][A-Z]\d{1}\b'
    order_date_pattern = r'Order Date\s*:\s*(\d{2}\.\d{2}\.\d{4})'
    billing_address_pattern = r'Billing Address\s*:\s*(.*?)(?=\n[A-Z])'
    description_pattern = r'Amount\s*\d+\s*([\s\S]*?)\s*\|\s*[A-Z0-9]+\s*\(.*?\)\s*HSN:\d+'
    tax_rate_pattern = r'\d+%?\s(?:CGST|SGST|IGST)'
    amount_pattern = r'TOTAL:\s*₹([0-9,]+\.\d{2})'
    sale_bill_pattern = r'TOTAL:\s*₹[0-9,]+\.\d{2}\s*₹([0-9,]+\.\d{2})'  # Pattern to extract sale bill amounts
    net_amount_pattern = r'HSN:\d+\s*₹([0-9,]+\.\d{2})'

    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        #+print(text) #for debugging
        

        # Extract data using regex
        invoices = re.findall(invoice_pattern, text)
        order_dates = re.findall(order_date_pattern, text)
        gst_numbers = re.findall(gst_pattern, text)
        billing_addresses = re.findall(billing_address_pattern, text, re.DOTALL)
        descriptions = re.findall(description_pattern, text, re.DOTALL)
        tax_rates = re.findall(tax_rate_pattern, text)
        amounts = re.findall(amount_pattern, text)
        sale_bills = re.findall(sale_bill_pattern, text)
        net_amounts = re.findall(net_amount_pattern, text)
        
        # Extract CGST and IGST
        igst_amounts = [match for match in tax_rates if "IGST" in match]
        cgst_amounts = [match for match in tax_rates if "CGST" in match]
        sgst_amounts = [match for match in tax_rates if "SGST" in match]

        # Process each entry
        for i in range(len(invoices)):  # Iterate based on the maximum number of matches
            order_date = order_dates[i] if i < len(order_dates) else ''
            order_date_split = order_date.split('.')
            month_index = int(order_date_split[1]) if order_date_split[1].isdigit() else 0
            month = calendar.month_name[month_index] if month_index > 0 else ''

            data.append({
                'SL no': i + 1,
                'Invoice No': invoices[i] if i < len(invoices) else '',
                'Month': month,
                'Order Date': order_dates[i] if i < len(order_dates) else '',
                'GST 1': gst_numbers[i] if i < len(gst_numbers) else '',
                'GST 2': gst_numbers[i + 1] if i + 1 < len(gst_numbers) else '',  # Adding the second GST number
                'Billing Address': billing_addresses[i] if i < len(billing_addresses) else '',
                'Description': descriptions[i ] if i  < len(descriptions) else '',
                'Tax Rate': tax_rates[i] if i < len(tax_rates) else '',
                'Net Amount': net_amounts[i] if i < len(net_amounts) else '',
                'Tax Amount': amounts[i] if i < len(amounts) else '',
                'IGST Amount': igst_amounts[i] if i < len(igst_amounts) else '',
                'CGST Amount': cgst_amounts[i] if i < len(cgst_amounts) else '',
                'SGST Amount': sgst_amounts[i] if i < len(sgst_amounts) else '',
                'Sale Bill': sale_bills[i] if i < len(sale_bills) else '',
                'Upload Time': current_time,  # Add the current date and time
            })
    
    return data

def save_to_excel(data, output_file):
    # Create a DataFrame and save to Excel
    df = pd.DataFrame(data)
    df.to_excel(output_file, index=False)

def process_all_pdfs_in_folder(folder_path, output_file, scanned_folder_path):
    all_data = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            print(f"Processing file: {filename}")  # Print the name of the PDF being processed
            data = extract_data_from_pdf(pdf_path)
            all_data.extend(data)

            # Move the processed PDF to the scanned folder
            scanned_pdf_path = os.path.join(scanned_folder_path, filename)
            shutil.move(pdf_path, scanned_pdf_path)
            print(f"Moved {filename} to {scanned_folder_path}")

    save_to_excel(all_data, output_file)

# Folder path containing PDFs and the output Excel file
folder_path = 'all_format_pdf'
output_file = 'output/Book1.xlsx'
scanned_folder_path = 'Scaned_pdfs'

# Create the scanned     folder if it doesn't exist
os.makedirs(scanned_folder_path, exist_ok=True)

# Process all PDFs in the folder and save the results to Excel
process_all_pdfs_in_folder(folder_path, output_file, scanned_folder_path)
print("Data extraction, saving, and moving files completed.")
