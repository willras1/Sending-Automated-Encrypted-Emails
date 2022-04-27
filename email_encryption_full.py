# Virtru app ID for rmrpracticeemail@gmail.com: 506388f6-2156-4b01-baa3-805e661cbb9f
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from virtru_sdk import Client, Policy, EncryptFileParams, LogLevel, Protocol
import csv
import imaplib                              
import email
from email.header import decode_header
import webbrowser
import pandas as pd
from pandas import read_excel
from pandas.io.parsers import read_csv 
import numpy as np
from time import sleep

#function to encrypt and send emails 
def email_encryption(smtp_to_address, file_path_plain, subject, body):
    # Variables 
    # SMTP Variables
    smtp_from_address = '#' # your email address 
    # Jake's email: whywaldodude@gmail.com
    password = '#' # your email passswod 
    # smtp_cc_address = "cc.recipinet@domain.com"
    # Virtru Variables
    virtru_appid = "#" # virtru app ID, found through Virtru control center->settings->developers->turn developer mode on
    virtru_owner = smtp_from_address
    # File Variables
    file_name_tdf = "testcsv.csv.tdf3.html" # file that you want encrypted + tdf3.html at the end
    file_path_tdf= file_path_plain + ".tdf3.html"  # file path of the file to be encrypted with .tdf3.html on the end 
    # file_path_plain is the path of the document to be encrypted 
    

    # Virtru encryption
    # Authentication
    client = Client(owner=virtru_owner, app_id=virtru_appid)
    policy = Policy()
    policy.share_with_users(smtp_to_address)
    param = EncryptFileParams(in_file_path=file_path_plain,
                                out_file_path=file_name_tdf)
    param.set_policy(policy)
    # encryption 
    client.encrypt_file(encrypt_file_params=param)
    client.update_policy_for_file(policy, file_path_tdf)


    # SMTP build and send message 
    msg = MIMEMultipart()
    msg['From'] = smtp_from_address
    msg['To'] = ",".join(smtp_to_address)
    # msg['CC'] = smtp_cc_address
    msg['Subject'] = subject # Enter email subject # make parameter 
    # body = "Test email body. Download attachment to view file." # Enter body of email, this part is not encrypted, only the file is # make parameter
    msg.attach(MIMEText(body, 'plain'))
    attachment = open(file_name_tdf, "rb")
    p = MIMEBase('application', 'octet-stream')
    p.set_payload((attachment).read())
    encoders.encode_base64(p)
    attachment_disposition = "attachment; filename= {}".format(file_name_tdf)
    p.add_header('Content-Disposition', attachment_disposition)
    msg.attach(p)
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(smtp_from_address, password)
    text = msg.as_string()
    s.sendmail(smtp_from_address, smtp_to_address, text)
    s.quit()
    # the recieving email has to press download on attachment, clicking on the attachment won't work 



    # read sent emails directly from gmail 
    # establish conneection with Gmail
    server ="imap.gmail.com"                     
    imap = imaplib.IMAP4_SSL(server)
    
    # intantiate the username and the passwprd
    username = smtp_from_address
    # password ="WilliamRMR1"
    
    # login into the gmail account
    imap.login(username, password)               
    
    # select the e-mails
    res, messages = imap.select('"[Gmail]/Sent Mail"')   
    
    # calculates the total number of sent messages
    messages = int(messages[0])
    
    # determine the number of e-mails to be fetched
    n = 1
    
    
    # iterating over the e-mails
    for i in range(messages, messages - n, -1):
        res, msg = imap.fetch(str(i), "(RFC822)")     
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1]) 
                
                # getting the sender's mail id
                sent_to = msg["to"] 

                
                sent_to_split = sent_to.split(',')
                
                if smtp_to_address == sent_to_split:
                    print('Program succesfully sent emails and verified!')

                else:
                    print('different')
                    with open('email_error.txt', 'w') as file:
                        file.write("the following emails were sent by the program: \n")
                        file.write(str(smtp_to_address) + '\n')
                        file.write("but these are the emails that gmail sent to: \n")
                        file.write(str(sent_to_split) + '\n')
                        file.write("error is in nonmatching emails")
                        file.close
                    
                    msg = MIMEMultipart()
                    body1 = "An error was found in these emails: "
                    file1 = 'email_error.txt' 
                    to_address_email_error = smtp_to_address  # change this to whoever will fix emails if an error happens 
                    attachment = open(file1, "rb")
                    p = MIMEBase('application', 'octet-stream')
                    p.set_payload((attachment).read())
                    encoders.encode_base64(p)
                    attachment_disposition = "attachment; filename= {}".format('email_error.txt')
                    p.add_header('Content-Disposition', attachment_disposition)
                    msg.attach(p)
                    msg['From'] = smtp_from_address
                    msg['To'] = to_address_email_error # ",".join(to_address_email_error)
                    msg['Subject'] = 'Error in automatic emails' 
                    body = body1
                    msg.attach(MIMEText(body, 'plain'))
                    s = smtplib.SMTP('smtp.gmail.com', 587)
                    s.starttls()
                    s.login(smtp_from_address, password)
                    text = msg.as_string()
                    s.sendmail(smtp_from_address, to_address_email_error, text)
                    s.quit()
                
    

email_encryption(['#'], "#", "Practice Email", "Test email body. Download attachment to view file. (You have to press the download button when you hover over the attachment)")
# comment in the different labels for these ^^^



host = 'imap.gmail.com'
username = '#' 
password = '#' 


def get_inbox():

    sleep(15)

    mail = imaplib.IMAP4_SSL(host)
    mail.login(username, password)
    mail.select("inbox")

    _, search_data = mail.search(None, 'UNSEEN')
    my_message = []

    for num in search_data[0].split():
        email_data = {}
        _, data = mail.fetch(num, '(RFC822)')
        _, b= data[0]
        email_message = email.message_from_bytes(b)

        for header in ['subject', 'to', 'from']:
            print("{}: {}".format(header, email_message[header]))
            email_data[header] = email_message[header]

        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode = True)
                email_data['body'] = body.decode()

            elif part.get_content_type() == "text/html":
                html_body = part.get_payload(decode = True)
                email_data['html_body'] = html_body.decode()
                my_message.append(email_data)
    

        #if email_message['to'] == 'rmrpracticeemail@gmail.com':
            #print('email bounced')
        if email_message['from'] == 'Mail Delivery Subsystem <mailer-daemon@googlemail.com>':
            print("email bounced from mail delivery") 

            msg = MIMEMultipart()
            body1 = "Check email, there's an incorrect email address that bounced"
            to_bounced_email = '#'  
            msg['From'] = username  
            msg['To'] = to_bounced_email # ",".join(to_address_email_error) 
            msg['Subject'] = 'An email bounced' 
            body = body1
            msg.attach(MIMEText(body, 'plain'))
            s = smtplib.SMTP('smtp.gmail.com', 587)
            s.starttls()
            s.login('cobra@rmrbenefits.com', password)
            text = msg.as_string()
            s.sendmail('cobra@rmrbenefits.com', to_bounced_email, text)
            s.quit()


    return my_message

if __name__ == "__main__":
    my_inbox = get_inbox()

