# This is a simple tool for generating invoices. The tool sends a simple plain
# text email containing invoice information to every recipient in a CSV file
# (up to 99999 recipients).

import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from time import localtime, strftime
import math


def calculate_ref_nro(invoice_type, year, member_number):
    no_str = str(invoice_type)+str(year)+str(member_number).zfill(5)  # ref no is always 8+1 digits long
    weights = [3, 7, 1, 3, 7, 1, 3, 7]
    checksum = 0.0
    for i in range(len(no_str)-1, -1, -1):
        checksum += int(no_str[i]) * weights[i]
    check_no = math.ceil(checksum / 10) * 10 - checksum
    if check_no == 10:
        check_no = 0
    complete_str = str(no_str) + str(int(check_no))
    # Comment below two lines for testing if testing material does not have
    # spaces every five characters
    for j in range(len(complete_str)-5, -1, -5):
        complete_str = complete_str[:j] + " " + complete_str[j:]
    return complete_str


def create_payment_options_str(invoice_info, member_number, current_year, plain_text=False):
    payment_options_str = ""
    newline = "<br>"
    if plain_text:
        newline = "\n"
    for row in invoice_info:
        addline = "Invoice type: " + row['Type'] + \
                  ", Total: " + row['Total'] + ", Reference Number: " + \
                  calculate_ref_nro(row['ID'], current_year, member_number) + newline
        payment_options_str += addline
    return payment_options_str


def main():
    # Config
    invoice_path = './invoice_types.csv'
    member_path = './member_list.csv'
    html_template_path = './templates/html_template.txt'
    text_template_path = './templates/plain_text_template.txt'
    current_year_long = strftime('%Y', localtime())
    current_year_short = strftime('%y', localtime())
    html_template = ''
    text_template = ''
    my_email = 'info@tampereenurheiluampujat.fi'

    # Import invoice HTML template
    with open(html_template_path, 'r') as template_file:
        html_template = template_file.read().replace('\n', '')

    # Import invoice plain text template
    with open(text_template_path, 'r') as template_file:
        text_template = template_file.read().replace('\n', '')

    # Define invoices
    # Invoice_info has two fields per row: Invoice type, and invoice total
    with open(invoice_path, 'r') as inv_file:
        invoice_info = list(csv.DictReader(inv_file, delimiter=';'))

        # Read recipient information from CSV file
        with open(member_path, 'r') as mem_file:
            members = csv.DictReader(mem_file, delimiter=';')

            # Go through each recipient, calculate reference number,
            # create text for email and send it
            for row in members:
                member_number = row['membernumber']
                member_type = row['membertype']
                first_name = row['firstname']
                surname = row['surname']
                address = row['address']
                zip = row['zip']
                city = row['city']
                recipient_email = row['email']
                html_msg = html_template.format(firstname=first_name,
                                        surname=surname,
                                        address=address,
                                        zip=zip,
                                        city=city,
                                        recipientemail=recipient_email,
                                        membernumber=member_number,
                                        membertype=member_type,
                                        paymentoptions=create_payment_options_str(invoice_info,
                                                                                  member_number,
                                                                                  current_year_short),
                                        year=current_year_long)
                text_msg = text_template.format(firstname=first_name,
                                        surname=surname,
                                        address=address,
                                        zip=zip,
                                        city=city,
                                        recipientemail=recipient_email,
                                        membernumber=member_number,
                                        membertype=member_type,
                                        paymentoptions=create_payment_options_str(invoice_info,
                                                                                  member_number,
                                                                                  current_year_short, True),
                                        year=current_year_long)
                print(html_msg)
                print(text_msg)
                with open("./emails/"+member_number+surname+first_name+current_year_long+".html", 'w') as html_file:
                    html_file.write(html_msg)

                with open("./emails/"+member_number+surname+first_name+current_year_long+".txt", 'w') as txt_file:
                    txt_file.write(text_msg)

                msg = MIMEMultipart('alternateive')
                msg['Subject'] = 'Test invoice from TaU'
                msg['From'] = my_email
                msg['To'] = recipient_email
                msg.attach(MIMEText(text_msg, 'plain'))
                msg.attach(MIMEText(html_msg, 'html'))
                




main()