# This is a simple tool for generating invoices. The tool sends a simple plain text email containing invoice information
# to every recipient in a CSV file (up to 99999 recipients).

import csv
import email
from time import localtime, strftime
import math


def calculate_ref_nro(type, year, memberno):
    no_str = str(type)+str(year)+str(memberno).zfill(5) # ref no is always 8+1 digits long
    weights = [3, 7, 1, 3, 7, 1, 3, 7]
    checksum = 0.0
    for i in range(len(no_str)-1, -1, -1):
        checksum += int(no_str[i]) * weights[i]
    check_no = math.ceil(checksum / 10) * 10 - checksum
    if check_no == 10:
        check_no = 0
    return str(no_str) + str(int(check_no))


# Config
invoice_path = './invoice_types.csv'
member_path = './test.csv'
template_path = './invoice_template.txt'
current_year = strftime('%y', localtime())

print(calculate_ref_nro(1, 16, 1))
print(calculate_ref_nro(2,16,1))


# # Import invoice template
# with open(template_path, 'r') as template_file:
#     data = template_file.read().replace('\n', '')
#
#
# # Define invoices
# # Invoice_info has two fields per row: Invoice type, and invoice total
# with open(invoice_path, 'r') as inv_file:
#     invoice_info = csv.DictReader(inv_file)
#
#
# # Read recipient information from CSV file
# with open(member_path, 'r') as mem_file:
#     members = csv.DictReader(mem_file)
#
# # Go through each recipient, calculate
# for row in members:
#     membernumber = row['membernumber']
#     type = row['membertype']
