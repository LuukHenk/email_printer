#!/usr/bin/env python3.6
##################################################################################
"""
## Email downloader ##

Download email attachments and body in pdf format using python3
Author: Luuk Perdaems
"""
##################################################################################

import base64
import email
import getpass
import imaplib
import os
import re
import pdfkit

##################################################################################
"""
Input data
"""
##################################################################################

# Set mail host and it's correct port
HOST = "smtp-mail.outlook.com"
PORT = 993

# Optional: Choose a pre-set account
PRE_SET_ACCOUNT = False
USER = "asdf@outlook.com" if PRE_SET_ACCOUNT else input("Email: ")
PASS = "asdf" if PRE_SET_ACCOUNT else getpass.getpass(prompt="Password: ")

# Set mailbox to open
MAILBOX = "Inbox"

# Set mail prefix (Don't leave empty)
FROM = '(FROM "manetje76@outlook.com")'
SUBJECT = '(SUBJECT "CnC / Online Bestellen")'

# Set attachment download prefix
ATTACH_PREFIX = "bon"

# Set the saving path of the pdf files
SAVING_PATH = "./output"

# Use the ordernumber for the output file name
BARCODE_REGEX = r"\d\d\d\d\d\d\d"

##################################################################################
" Functions and main loop "
##################################################################################

def save_file(data, path, print_save=True, byte_object=True):
    " Save data to path "

    data_type = 'wb' if byte_object else 'w'

    file_ = open(path, data_type)
    file_.write(data)
    file_.truncate()
    file_.close()

    if print_save:
        print(f"Saved {path}")

def main():
    " Main loop of the programm "

    # Load mail host
    mail = imaplib.IMAP4_SSL(HOST, PORT)
    # Login user
    mail.login(USER, PASS)
    #Select mailbox
    mail.select(MAILBOX)
    # Selection for specific emails
    _, selected_mails = mail.search(None, FROM, SUBJECT)

    # Loop over all the found emails
    for selected_mail_data in selected_mails[0].split():
        # Fetch and decode message
        _, mess = mail.fetch(selected_mail_data, "RFC822")
        message = email.message_from_string(mess[0][1].decode("utf-8"))

        # Get barcode of current email message (for saves later)
        barcode = re.search(re.compile(BARCODE_REGEX), message["Subject"]).group(0)

        # Split the email in parts
        # All image data in the email will be saved in image_ids
        # All html data in the email will be saved in html_data
        image_ids = {}
        html_data = []
        for part in message.walk():

            # Receive decoded payload of the message part
            payload = part.get_payload(decode=True)

            # Save the image source code with the cid as key
            if "image" in part["Content-Type"]:
                image_ids[part["Content-ID"]] = "data:" + part["Content-Type"].split(";")[0]
                image_ids[part["Content-ID"]] += ";base64,"
                image_ids[part["Content-ID"]] += base64.b64encode(payload).decode("utf-8")

            # Check if the message part is an attachment
            # also select for the word 'bon' is in the filename (Specific for a IKEA delivery)
            # Finally, save the attachment in the saving path
            if part.get_content_maintype != "multipart" and \
               part.get("Content-Disposition") is not None and \
               ATTACH_PREFIX in part.get_filename():
                save_file(payload, os.path.join(SAVING_PATH + barcode + "_bon.pdf"))

            # Check if there is text in the email and add it to the html string
            if part.get_content_type() == "text/html":
                html_data.append(payload.decode("utf-8"))

        # Match the cid in the html code with the cid key from the image dict.
        # Replace the html src with the image source code
        for html_str in html_data:
            for image_id in image_ids:
                image_loc_in_html = re.search(image_id[1:-1], html_str)
                image_html_str = html_str[image_loc_in_html.start()-4:image_loc_in_html.end()]
                html_str = re.sub(image_html_str, image_ids[image_id], html_str)
            ##save the complete html as a html file
            pdfkit.from_string(html_str, os.path.join(SAVING_PATH + barcode + "_mail.pdf"))

main()
