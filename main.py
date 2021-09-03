import pandas as pd
import smtplib
from email.message import EmailMessage
import os
from socket import gaierror


contacts_file = os.getcwd()+'/files/people.xlsx'
contacts_frame = pd.read_excel(contacts_file)
# print(contacts)
# Contacts file include Fnamem Lname email address and age

# Build the body of the mail
num_contacts = len(contacts_frame)
def buildMail(fname, lname, email_add, add, age) :
    msg = f"""
    Hello {fname},
    I hope you are doing well. 
    We are sending this test mail to you as part of a bulk mailing scheme. 
    We currently have several contacts and will be updated them as well.
    Please confirm that you are {age} and living at {add}.
    """

    mail_obj = EmailMessage()
    mail_obj['Subject'] = "Group Info Update"
    mail_obj['To'] = f"{fname+lname}<{email_add}>"
    mail_obj['From'] = "sender<yourmail@mmail.com>"
    mail_obj.set_content(msg)
    return mail_obj

try :
    #Always verify that you are using the right port to connect to the host
    # Advisable to use smtplib.SMTP_SSL for secured connections
    server_con = smtplib.SMTP("smtp.mailtrap.io", 2525)
    server_con.login("5530fd2492f52d", "22971d1f0bfe53")
    for i in range(num_contacts) :
        fname = contacts_frame.loc[i, 'Fname']
        lname = contacts_frame.loc[i, 'Lname']
        email_add = contacts_frame.loc[i, 'email']
        loc_add = contacts_frame.loc[i, 'address']
        age = contacts_frame.loc[i, 'age']
        email_body  = buildMail(fname, lname, email_add, loc_add, age)
        server_con.send_message(email_body)
        print("Email sent successfully!!" + str(fname+lname))
except (gaierror, ConnectionRefusedError):
 print('Failed to connect to the server. Bad connection settings?')
except smtplib.SMTPServerDisconnected:
 print('Failed to connect to the server. Wrong user/password?')
except smtplib.SMTPException as e:
 print('SMTP error occurred: ' + str(e))