This module will help to send outlook email by just creating the instance.

Arguments:
    mail_to - mail to string 
    mail_subject - mail subject string
    mail_body - mail body if HTML them HTML should be true otherwise it will take the body as plain text
    mail_attachment - can give a list of path to add multiple files or can give a string to add a single file, default is None
    HTML - if the body text in html format, default is False