# MIT LICENSE
# Copyright 2021 Jeremy Tennant
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

# ---------------------------------------

# Rename the olm file to a .zip file, and unzip the contents of the OLM using any zip tool
# Then modify this script to meet your needs
# The script reassembles EML files (Standard MIME) from the XML structure provided in the OLM
# It's hacky, but should get the job done....

#UPDATE BY xperrotti: As in March 2023, put this script in the root of the unzipped file (the folder that contains the 'Accounts' and 'Local' folders).
#If needed (e.g., if your 'Local' folder (and subfolders) don't have anything inside them, modify the line 'for path in Path('Local').rglob('*.xml'):'
#to get the appropriate value. For example, './Accounts/<username>@<domain>/com.microsoft.__Messages/INBOX'.
#After that, run the script. You will get several .eml files inside the folder. You can convert them to a single MBOX file 
#using the script located at "
#to get the appropriate value. For example, './Accounts/<username>@<domain>/com.microsoft.__Messages/INBOX'.

import xml.etree.ElementTree as ET
from pathlib import Path

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import os
import re
import datetime

def main():
    xml_filelist =  []

    for path in Path('Local').rglob('*.eml'):
        os.remove(path)

    for path in Path('Local').rglob('*.xml'):
        xml_filelist.append(path)

    for xml_filename in xml_filelist:
        print(f'Looking @ {xml_filename}')
        try:
            tree = ET.parse(xml_filename)
        except ET.ParseError:
            print('Unprocessable - unable to parse the file located at: {xml_filename}')
            continue

        try:
            root = tree.getroot()
        except:
            print('Unprocessable - Unable to get the xml root for the file located at: {xml_filename}')
            continue

        try:
            email_xml = root.find('email')
        except:
            print(f'Unprocessable - Unable to find the email root xml for {xml_filename}')
            continue

        try:
            email_from_list = email_xml.find('OPFMessageCopyFromAddresses')
            email_from = extract_single_email_address(email_from_list)
        except:
            email_from = None
            # print(f'Unable to find FROM for {xml_filename}')

        try:
            email_to_list = email_xml.find('OPFMessageCopyToAddresses')
            email_to = extract_multiple_email_addresses(email_to_list)
        except:
            email_to = None
            # print(f'Unable to find TO for {xml_filename}')

        try:
            email_cc_list = email_xml.find('OPFMessageCopyCCAddresses')
            email_cc = extract_multiple_email_addresses(email_cc_list)
        except:
            email_cc = None
            # print(f'Unable to find CC for {xml_filename}')

        try:
            email_bcc_list = email_xml.find('OPFMessageCopyBCCAddresses')
            email_bcc = extract_multiple_email_addresses(email_bcc_list)
        except:
            email_bcc = None
            # print(f'Unable to find BCC for {xml_filename}')


        email_subject = extract_subject(email_xml)
        email_body_html = extract_html_body(email_xml)
        email_received_time = extract_receive_time(email_xml)

        email_attachment_list = email_xml.find('OPFMessageCopyAttachmentList')
        email_attachment_metas = extract_attachment_details(email_attachment_list)
        
        # try:
        #     email_attachment_list = email_xml.find('OPFMessageCopyAttachmentList')
        #     email_attachment_metas = extract_attachment_details(email_attachment_list)
        # except:
        #     email_attachment_metas = []
        #     print(f'Something happened with the attachment')




        # Create a multipart message and set headers
        message = MIMEMultipart()
        if email_from:
            message["From"] = email_from
        if email_to:
            message["To"] = email_to
        if email_cc:
            message["Cc"] = email_cc
        if email_bcc:
            message["Bcc"] = email_bcc

            
        message["Date"] = datetime.datetime.strptime(email_received_time, '%Y-%m-%dT%H:%M:%S').strftime('%a, %d %b %Y %H:%M:%S +1000')
        message["Subject"] = email_subject
        message.attach(MIMEText(email_body_html, "html"))



        # Add attachments
        for attachment in email_attachment_metas:
            email_part = process_attachment(attachment)
            message.attach(email_part)

        text = message.as_string()

        clean_filename = re.sub('[^0-9a-zA-Z]+', ' ', email_subject)[0:60]
        f = open(f'{xml_filename}_{clean_filename}.eml', "a")
        f.write(text)
        f.close()


def extract_subject(xml, elementName = 'OPFMessageCopySubject'):
    interesting_text = xml.find(elementName)

    if interesting_text == None:
        return "Unknown Subject"
    elif type(interesting_text) == ET.Element and interesting_text.text:
        return interesting_text.text
    else:
        return ""

def extract_html_body(xml, elementName = 'OPFMessageCopyHTMLBody'):
    interesting_text = xml.find(elementName)


    if interesting_text == None:
        return "<p>No Email Body</p>"
    elif type(interesting_text) == ET.Element and interesting_text.text:
        return interesting_text.text
    else:
        return "<p>No Email Body</p>"

def extract_receive_time(xml, elementName = 'OPFMessageCopyReceivedTime'):
    interesting_text = xml.find(elementName)

    if type(interesting_text) == ET.Element:
        return interesting_text.text
    else:
        return "2222-12-31T23:59:59"
1
def extract_single_email_address(xml, element_name = 'emailAddress'):
    result = xml.find(element_name)
    return _return_formatted_email(result)

def extract_attachment_details(xml, element_name = 'messageAttachment'):
    attachments = []
    results = None
    
    if xml:
        results = xml.findall(element_name)

    if results:
        for attachment in results:
            parsed_attachment = {
                'content_type' : attachment.get('OPFAttachmentContentType'),
                'attachment_name' : attachment.get('OPFAttachmentName'),
                'attachment_location' : attachment.get('OPFAttachmentURL')
            }
            attachments.append(parsed_attachment)

    return attachments

def extract_multiple_email_addresses(xml, element_name = 'emailAddress'):
    addresses = []
    results = xml.findall(element_name)
    for address in results:
        formatted_email_address = _return_formatted_email(address)
        if formatted_email_address not in addresses:
            addresses.append(formatted_email_address)

    return str(addresses).replace('[','').replace(']','').replace("'",'').replace('"','')
            
def _return_formatted_email(xml):
    name = xml.get('OPFContactEmailAddressName')
    address = xml.get('OPFContactEmailAddressAddress')

    if name and address:
        return f'{name} <{address}>'
    elif address:
        return address
    else:
        return 'unknown@email.address'

def process_attachment(attach):
    filename = "./" + attach['attachment_location']

    # Open file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={attach['attachment_name']}",
    )

    return part


if __name__ == "__main__":
    main()