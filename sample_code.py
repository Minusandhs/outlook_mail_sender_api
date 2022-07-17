import outlook_email_sender as ols

#____Multiple Files_______
# attach = [r'<file path>',r'<file path>']

#________Folder________
# attach = r'<folder path>'

#_______SingleFile__________
# attach = r'<single file path>'

#_______EmailList____________
# email_to = [<mailid>,'<mailid>','<mailid>']

#_______EmailIteration____
# for person in email_to:
#     olmail = ols.outlook_email_sender(person,'Hi this is test','<h1>This is html body<h1>',mail_attachment_folder=attach,HTML=True)
#     olmail.save_email()
#     olmail = None

#______SingleFileOrFolderAttachment______
# olmail = ols.outlook_email_sender('minusandhs@gmail.com','Hi this dis test','<h1>This is html body<h1>',mail_attachment=attach,HTML=True)

#______FolderAttachment______
# olmail = ols.outlook_email_sender('minusandhs@gmail.com','Hi this dis test','<h1>This is html body<h1>',mail_attachment_folder=attach,HTML=True)

#______PlainTextBody_____
# olmail = ols.outlook_email_sender('minusandhs@gmail.com','Hi this dis test','This is plain text body',mail_attachment_folder=attach,HTML=False)

#______AddMultipleContacts_____
#Seperate mailId ";" to add multiple contacts
#olmail = ols.outlook_email_sender('<contact 1>;<contact 2>;...;...','Hi this dis test','<h1>This is body<h1>',mail_attachment=attach,HTML=True)

# olmail.display_email()
# olmail.save_email()
# olmail.send_email()