import os
import pandas as pd
import webbrowser
import random
import datetime
import xlrd
import pdfkit
import subprocess
import time
current_directory = os.getcwd()
#Testing
print(current_directory)
# Load data from Excel file into a DataFrame
xls_path_rel = r'Data\MyBills1.xls'
html_template_file_rel = r'SampleBillHTML\SampleBill1.html'
output_folder_rel = r'Output'
# Path to the directory containing HTML files
html_folder_rel = r'Output'

# Path to the directory where PDF files will be saved
pdf_folder_rel = r'OutputPDF'

xls_path = os.path.join(current_directory, xls_path_rel)
html_template_file = os.path.join(current_directory, html_template_file_rel)
output_folder = os.path.join(current_directory, output_folder_rel)
html_folder = os.path.join(current_directory, html_folder_rel)
pdf_folder = os.path.join(current_directory, pdf_folder_rel)

try:
    df = pd.read_excel(xls_path)
except Exception as e:
    print(f"Error loading Excel file: {e}")
    # Handle the error as needed

# Read HTML template from a file

with open(html_template_file, 'r') as template_file:
    html_template = template_file.read()

save_html_template=html_template
# Convert column names to lowercase
df.columns = df.columns.str.lower()

# Output folder

# Iterate over rows in the DataFrame
for _, row in df.iterrows():
    # Additional logic for bill generation based on new parameters
    from_date = pd.to_datetime(row['fromdate'])
    to_date = pd.to_datetime(row['todate'])
    total_monthly_amt = row['totalmonthlyamt']
    total_bills_per_month = int(row['totalbillspermonth'])
    vtype = row['vtype'].lower()  # Assuming vtype is a column in the DataFrame
    pumplogo = row['pumplogo'].lower()
    userbillid = row['id']
    # Loop over months
    for i in range((to_date - from_date).days // 30 + 2):
        current_month = from_date + pd.DateOffset(months=i)

        # Calculate total bill amount for the month
        total_monthly_amt_remaining = total_monthly_amt

        # Loop for totalbillspermonth
        for j in range(total_bills_per_month):
            days_in_month = pd.Timestamp(current_month.year, current_month.month, 1) + pd.DateOffset(months=1) - pd.Timestamp(current_month.year, current_month.month, 1)
            current_month_new = current_month + pd.DateOffset(days=random.randint(0, days_in_month.days))
       
            html_template = save_html_template

            # Calculate parameters for each bill
            # You can adjust this part based on your specific logic
            rate = row['rate'] * random.uniform(0.75, 1.25)
            receipt_no = random.randint(1000, 50000) + j

            # Distribute total_monthly_amt among bills
            if j == total_bills_per_month - 1:
                #bill_amount = total_monthly_amt_remaining
                bill_amount=round(total_monthly_amt_remaining, 0)
            else:
                bill_amount = round(random.uniform(0, total_monthly_amt_remaining), 0)

            ltr = round(bill_amount / rate, 2)  # Round to two decimal points

            # Generate random time between 10 AM and 9 PM
            random_hour = random.randint(10, 21)
            random_minute = random.randint(0, 59)
            time_str = f"{random_hour:02d}:{random_minute:02d}"

            # Replace placeholders with actual values
            placeholders = {
                "__stationname__": row['stationname'],
                "__stationadd__": row['stationadd'],
                "__dispetrol__": row['dispetrol'],
                "__rate__": row['rate'],
                "__vtype__": vtype,
                "__vno__": row['vno'],
                "__mode__": row['mode'],
                "__name__": row['name'],
                "__receiptno__": receipt_no,
                "__amt__": bill_amount,
                "__ltr__": ltr,
                "__date__": current_month_new.strftime('%Y-%m-%d'),
                "__time__": time_str,
                "__pumplogo__":row['pumplogo'],
            }

            for placeholder, value in placeholders.items():
                html_template = html_template.replace(placeholder, str(value))

            # Update the remaining monthly amount
            total_monthly_amt_remaining -= bill_amount

            # Generate bill name
            bill_name = f"{userbillid}_{pumplogo}_{vtype}_{current_month.strftime('%Y-%m-%d')}_{receipt_no}"

            # Save the updated HTML to a file
            output_html_path = os.path.join(output_folder, f"{bill_name}.html")
            with open(output_html_path, 'w') as updated_html_file:
                updated_html_file.write(html_template)

            print(f"Bill saved as: {output_html_path}")


#pdfkit.from_file(html_path, pdf_path, config={'wkhtmltopdf': 'C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe'})


# Make sure the output PDF folder exists
os.makedirs(pdf_folder, exist_ok=True)

# Iterate over HTML files in the HTML folder
for html_file in os.listdir(html_folder):
    if html_file.endswith('.html'):
        html_path = os.path.join(html_folder, html_file)
        pdf_file = os.path.splitext(html_file)[0] + '.pdf'
        pdf_path = os.path.join(pdf_folder, pdf_file)

        command = [
            r'C:\Program Files\Google\Chrome\Application\chrome.exe',  # Path to your Chrome executable
            '--headless',
            '--disable-gpu',
            '--no-print-to-pdf-header',
            '--disable-software-rasterizer',
            '--print-to-pdf=' + pdf_path,
            'file://' + html_path
        ]

        # Run the command
        subprocess.run(command, check=True)

        # Wait for a few seconds to ensure the print operation is complete
        time.sleep(2)
        os.remove(html_path)
        print(f'PDF saved at: {pdf_path}')

# Open the updated HTML file in the default web browser
#webbrowser.open(os.path.join(output_folder, 'updated_template.html'))
