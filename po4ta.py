import imaplib
import email
import os
import argparse
import datetime
import time

parser = argparse.ArgumentParser(description='Process passw')
parser.add_argument('psw', type=str, help='passw')
args=parser.parse_args()

passw = args.psw
server = "172.28.142.*"
ports = "993"
user = 'user'
password = psw
outputdir ="/home/akasymalieva/hur"


subject = 'HUR' #subject line of the emails you want to download attachments from

# connects to email client through IMAP
def connect(server, ports, user, password):
    m = imaplib.IMAP4_SSL(server, ports)
    m.login(user, password)
    m.select()
    return m

# downloads attachment for an email id, which is a unique identifier for an
# email, which is obtained through the msg object in imaplib, see below 
# subjectQuery function. 'emailid' is a variable in the msg object class.

def downloaAttachmentsInEmail(m, emailid, outputdir):
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1]
    mail = email.message_from_bytes(email_body)
    date=datetime.datetime.now()
    name_add=date.strftime("%Y.%m.%d.%H%M%S")+"."
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        try:
            if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
              open(outputdir + '/'+name_add+"_" + part.get_filename(), 'wb').write(part.get_payload(decode=True))
              m.store( emailid,'+FLAGS', '\\Deleted')
        except:
            print("cant save " + str(emailid))


# download attachments from all emails with a specified subject line
# as touched upon above, a search query is executed with a subject filter,


def subjectQuery(subject):
    m = connect(server, ports, user, password)
    m.select("Inbox")
    typ, msgs = m.search(None, 'SUBJECT ' + subject )
    msgs_ = msgs[0].split()
    i = 0
    for emailid in msgs_:
        print(emailid)
        downloaAttachmentsInEmail(m, emailid, outputdir)        
        i+=1
    m.expunge()
              
    

timing = time.time()
subjectQuery(subject)
while True:
    if time.time() - timing > 600.0:
        timing = time.time()
        subjectQuery(subject)
