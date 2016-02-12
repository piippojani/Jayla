"""
JAYLA - a simple invoice generator and emailing tool.
Version 2016.2.12 (0.1)

This is a simple tool for generating invoices for yearly fees, for example.
The tool sends a simple html email containing invoice information to
every recipient in a CSV file (up to 99999 recipients).

Copyright (c) 2016, Jani Piippo
jani.vo.piippo@gmail.com

All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:
    * Redistributions of source code must retain the above copyright
      notice, this list of conditions and the following disclaimer.
    * Redistributions in binary form must reproduce the above copyright
      notice, this list of conditions and the following disclaimer in the
      documentation and/or other materials provided with the distribution.
    * Neither the name of the Tampere University of Technology nor the
      names of its contributors may be used to endorse or promote products
      derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
"""

import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from time import localtime, strftime
import math
import traceback
import sys


"""
Calculates the Finnish type reference number
The reference number can be practically anything, but
this system creates a form of abbcccccd, in which
a == invoice type
bb == current year
ccccc == member number, padded with zeros if necessary
d == control number
"""
def calculate_ref_nro(invoice_type, year, member_number):
    # Make the number into a string, so we can index it and it's immutable
    no_str = str(invoice_type)+str(year)+str(member_number).zfill(5)
    # Weights needed for checksum calculation
    weights = [3, 7, 1, 3, 7, 1, 3, 7]
    checksum = 0.0
    # Calculate checksum (Finnish type)
    for i in range(len(no_str)-1, -1, -1):
        checksum += int(no_str[i]) * weights[i]
    # Calculate control number (d)
    check_no = math.ceil(checksum / 10) * 10 - checksum
    if check_no == 10:
        check_no = 0
    # Add control number at the end
    complete_str = str(no_str) + str(int(check_no))
    # Add spacing to every five character for clarity
    # Comment below two lines for testing if testing material does not have
    # spaces every five characters
    for j in range(len(complete_str)-5, -1, -5):
        complete_str = complete_str[:j] + " " + complete_str[j:]
    return complete_str


"""
Creates a atring containing each payment option on its own line.
The line consists of invoice type, invoice total for that type and the
respective reference number created using calculate_ref_nro
"""
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
    # If the paths are provided on the command line, we use them
    if len(sys.argv) == 4:
        invoice_path = sys.argv[1]
        member_path = sys.argv[2]
        html_template_path = sys.argv[3]
    # Current year number for email and reference numbers.
    current_year_long = strftime('%Y', localtime())
    current_year_short = strftime('%y', localtime())

    # Ask the user for necessary information
    print("#################################################")
    print("Welcome to JAYLA, a simple invoice emailing tool.")
    print("#################################################")
    print("\nPlease, enter the following email service information.")
    fromaddr = input("Enter the sender email: ")
    server_smtp = input("Enter SMTP server address: ")
    port_smtp = int(input("Enter SMTP port number: "))
    username = input("Enter the username to SMTP server: ")
    password = input("Enter the password to SMTP server: ")
    tls = input("Use TLS instead of SSL? ('y' / Leave empty if you want to use SSL): ")
    msg_subject = input("Enter the email subject: ")

    print("\nStarting process...")

    # Import invoice HTML template
    with open(html_template_path, 'r') as template_file:
        html_template = template_file.read().replace('\n', '')
        print("Email template processed...")

    # Define invoices
    # Invoice_info has two fields per row: Invoice type, and invoice total
    with open(invoice_path, 'r') as inv_file:
        invoice_info = list(csv.DictReader(inv_file, delimiter=';'))
        print("Invoice information processed...")

        # Read recipient information from CSV file
    with open(member_path, 'r') as mem_file:
        members = list(csv.DictReader(mem_file, delimiter=';'))
        print("Member list processed...\n")

        # Go through each recipient, calculate reference number,
        # create text for email and send it
        email_total = len(members)
        emails_sent = 0
        for row in members:
            member_number = row['member_number']
            member_type = row['member_type']
            first_name = row['first_name']
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
            # Save invoice to html file for archiving
            with open("./emails/"+member_number+surname+first_name+current_year_long+".html", 'w') as html_file:
                html_file.write(html_msg)
                print("HTML file for {} {} saved...".format(first_name, surname))
            # Construct email message
            msg = MIMEMultipart('alternative')
            msg['Subject'] = msg_subject
            msg['From'] = fromaddr
            msg['To'] = recipient_email
            msg.attach(MIMEText(html_msg, 'html'))

            try:
                # Connect to SMTP server, open TLS instead of SSL if necessary
                server = smtplib.SMTP(host=server_smtp, port=port_smtp)
                if len(tls) != 0:
                    server.starttls()
                # Flag for debug mode
                server.set_debuglevel(False)
                # Login to server
                server.esmtp_features['auth'] = 'LOGIN PLAIN'
                server.login(username, password)
                # Send the email
                server.sendmail(fromaddr, recipient_email, str(msg))
                # Disconnect
                server.quit()
                emails_sent += 1
                print("Sent email to {}... [{} / {} done]\n".format(recipient_email, emails_sent, email_total))

            except smtplib.SMTPServerDisconnected:
                print("smtplib.SMTPServerDisconnected")
            except smtplib.SMTPResponseException as e:
                print ("smtplib.SMTPResponseException: {} {}".format(str(e.smtp_code), str(e.smtp_error)))
            except smtplib.SMTPSenderRefused:
                print("smtplib.SMTPSenderRefused")
            except smtplib.SMTPRecipientsRefused:
                print("smtplib.SMTPRecipientsRefused")
            except smtplib.SMTPDataError:
                print("smtplib.SMTPDataError")
            except smtplib.SMTPConnectError:
                print("smtplib.SMTPConnectError")
            except smtplib.SMTPHeloError:
                print("smtplib.SMTPHeloError")
            except smtplib.SMTPAuthenticationError:
                print("smtplib.SMTPAuthenticationError")
            except Exception as e:
                print("Exception", e)
                print(traceback.format_exc())
                print(sys.exc_info()[0])

    print("Success! All {} emails were sent!".format(emails_sent))


main()
