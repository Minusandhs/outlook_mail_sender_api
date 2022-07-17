import win32com.client as win32


class outlook_email_sender():
    def __init__(self,mail_to,mail_subject,mail_body,mail_attachment=None,HTML = False):
        self.olApp = win32.Dispatch('Outlook.Application')
        self.mail_item = self.olApp.CreateItem(0)
        self.mail_item.To = mail_to
        self.mail_item.Subject = mail_subject
        self.attachment = mail_attachment
        self.is_html = HTML
        self.mail_body = mail_body
        
        #Check if there is a mail attachment and whether its str or list
        if self.attachment == None:
            pass

        elif self.attachment != None and isinstance(self.attachment,list):

            for attach in self.attachment:
                self.mail_item.Attachments.Add(attach)

        elif (self.attachment != None and isinstance(self.attachment,str)):
            self.mail_item.Attachments.Add(self.attachment)

        else:
            raise TypeError('The attachment path should in string or list')
        
        #Add the mail body by checking whether it is HTML
        if self.is_html:
            self.mail_item.HTMLBody = self.mail_body
        else:
            self.mail_item.Body = self.mail_body
        

    def save_email(self):
        self.mail_item.Save()

    def send_email(self):
        self.mail_item.Send()

    def display_email(self):
        self.mail_item.Display()
