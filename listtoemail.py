import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import re
import openpyxl
 
#Specify the excel's files to open
doc = openpyxl.load_workbook ("./FILENAME.xslx")
sheet = doc.get_sheet_by_name("SHEET")

#Initialize name and mail lists 
mails=[]
names=[]
lastnames=[]
fullname=[]
x=0
for row in sheet.iter_rows():
    name = row[0].value
    if name != None:
        names.append(name)
    lastname = row[1].value
    if lastname != None:
        lastnames.append(lastname)
    mail = row[2].value
    if mail != None:
        mails.append(mail)
    if name != None and lastname != None: 
        fullname.append(name+lastname)
    

#Sort all the names in a list
def sorted_alphanumeric(data):
    convert = lambda text: int(text) if text.isdigit() else text.lower()
    alphanum_key = lambda key: [ convert(c) for c in re.split('([0-9]+)', key) ] 
    return sorted(data, key=alphanum_key)

content = os.listdir('DIRECTORY_PATH')
content = sorted_alphanumeric(content) 

#Send emails one by one to all the mails
c = 0
for receiver in mails: 
    # Initialize paremeters for sending
    sender = 'someone@this_is_an_example.com'
    subject = 'example'
    body = """<h2>¡Hello World!</h2>""" #You can write html 
       
    s = str(c+1)
    route_attached = './path of the files '+ fullname[c] + '.pdf' #In case you are attaching a file
    name_attached = 'name of the file'+ fullname[c] + '.pdf' #Does not need to have the fullname

    #Create message object with email handler
    message = MIMEMultipart()
    
    #Stablished message attributes 
    message['From'] = sender
    message['To'] = receiver
    message['Subject'] = subject
        
    #Attach the parts of the message
    message.attach(MIMEText(body, 'html'))
        
    #File attachment
    file_attached = open(route_attached, 'rb')    
    attached_MIME = MIMEBase('application', 'pdf')
    attached_MIME.set_payload((file_attached).read())
    file_attached.close()
    encoders.encode_base64(attached_MIME)
    attached_MIME.add_header('Content-Disposition', "attachment", filename=name_attached)
    message.attach(attached_MIME)
        
    #Stablish connection with the server
    session_smtp = smtplib.SMTP('smtp.gmail.com', 587)
    session_smtp.starttls()
 
    session_smtp.login('someone@this_is_an_example.com','thisisyourpassword')
    #The message have to be a string(remember it have an image attached to it)
    text = message.as_string()
    session_smtp.sendmail(sender, receiver, text)

    #Close the connection
    session_smtp.quit()

    #Count of the messages sent
    print("Email "+s+" sent")
    c = c + 1

