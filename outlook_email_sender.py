import win32com.client as win32


class outlook_email_sender():
    def __init__(self,mail_to,mail_subject,mail_body,mail_attachment=None,HTML = False,mail_attachment_folder=None):
        self.olApp = win32.Dispatch('Outlook.Application')
        self.mail_item = self.olApp.CreateItem(0)
        self.mail_item.To = mail_to
        self.mail_item.Subject = mail_subject
        self.attachment = mail_attachment
        self.is_html = HTML
        self.mail_body = mail_body
        self.mail_attachment_folder = mail_attachment_folder
        
        #Check if there is a mail attachment and whether its str or list
        if self.attachment == None:
            pass

        elif self.attachment != None and isinstance(self.attachment,list):
            
            try:
                for attach in self.attachment:
                    self.mail_item.Attachments.Add(attach)
            except:
                raise ('File not found')
        elif (self.attachment != None and isinstance(self.attachment,str)):

            try:
                self.mail_item.Attachments.Add(self.attachment)
            except:
                raise ('File not found')

        else:
            raise ('The attachment path should in string or list')
        
        #Add the mail body by checking whether it is HTML
        if self.is_html:
            self.mail_item.HTMLBody = self.mail_body
        else:
            self.mail_item.Body = self.mail_body

        #If a folder path given as a parameter
        if mail_attachment_folder == None:
            pass
        else:
            try:
                file_list = file_in_dir(self.mail_attachment_folder)
                for attach in file_list:
                    self.mail_item.Attachments.Add(attach)

            except:
                raise ("Please select a valid folder")

    ##Check whether one object is equal to other object
    def __eq__(self,other):
        if (self.mail_item.To == other.mail_item.To) and \
            (self.mail_item.Subject == other.mail_item.Subject) and\
                (self.mail_body == other.mail_body) and \
                    (self.attachment == other.attachment) and\
                        (self.mail_attachment_folder == other.mail_attachment_folder):
            return True
        else:
            return False

    
    @property
    def mail_to(self):
        return self.mail_item.To
    
    @property
    def mail_subject(self):
        return self.mail_item.Subject


    def save_email(self):
        self.mail_item.Save()

    def send_email(self):
        self.mail_item.Send()

    def display_email(self):
        self.mail_item.Display()


def file_in_dir(path):
    import os
    
    # folder path
    dir_path = path

    # list to store files
    res = []

    # Iterate directory
    for path in os.listdir(dir_path):
        # check if current path is a file
        if os.path.isfile(os.path.join(dir_path, path)):
            res.append(r'{}'.format(dir_path + "\\" + path))

    return res