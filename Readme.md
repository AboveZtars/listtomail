# listtomail
_This proyect was created to import emails from an excel file and create and attachment following the data of the excel workbook. Then create a message to send it over to those emails_

The creation of this project was intended for a very specific purpose, it search for a pdf file and attach it to a message to send it. Altough you can use it to send a lot of mails without an attachment with just plain text or plain html.

Modifying this:
```
route_attached = './path of the files '+ fullname[c] + '.pdf'
```
You can select the type of file you are sending
```
route_attached = './path of the files '+ fullname[c] + '.HERE GOES THE EXTENSION'
```
_But be careful_
You have to pick what kind of MIMEBase you are going to use.
```
attached_MIME = MIMEBase('application', 'pdf')
```
Here is a list of MIME types:
(https://developer.mozilla.org/en-US/docs/Web/HTTP/Basics_of_HTTP/MIME_types/Common_types)

## Less secure apps

In order to use this app with your email account you have to allow less secure apps on your google account
(https://myaccount.google.com/lesssecureapps)