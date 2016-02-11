# This is a simple tool for generating invoices. The tool sends a simple plain text email containing invoice information
# to every recipient in a CSV file.

import csv
import email
from time import localtime, strftime


# Config
invoice_path = './invoice_types.csv'
member_path = './test.csv'
template_path = './invoice_template.txt'
current_year = strftime('%y', localtime())
print(current_year)

# Import invoice template
with open(template_path, 'r') as template_file:
    data = template_file.read().replace('\n', '')

# Define invoices
# Invoice_info has two fields per row: Invoice type, and invoice total
with open(invoice_path, 'r') as inv_file:
    invoice_info = csv.DictReader(inv_file)

# Read recipient information from CSV file

with open(member_path, 'r') as mem_file:
    members = csv.DictReader(mem_file)

