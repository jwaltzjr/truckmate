import collections
import os
import smtplib

import env

from email.MIMEMultipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

EmailAttachment = collections.namedtuple('EmailAttachment', ['filename', 'content'])

class KrcEmail():

    def __init__(self, send_to, send_from='reports@krclogistics.com',
            subject='', message='', message_html=None, attachments=[],
            server='smtp.office365.com', port=587, password=env.email_password.value):

        self.send_to = send_to
        self.send_from = send_from

        self.subject = subject
        self.message = message
        self.message_html = message_html
        self.attachments = [EmailAttachment._make(x) for x in attachments]

        self.server = server
        self.port = port
        self.password = password

        self.email = None
        self.build(self.send_to, self.send_from, self.subject, self.message, self.message_html, self.attachments)

    def build(self, send_to, send_from, subject, message, message_html, attachments):
        email_object = MIMEMultipart('alternative')

        email_object['To'] = ', '.join(send_to)
        email_object['From'] = send_from
        email_object['Subject'] = subject

        email_object.attach(MIMEText(message, 'plain'))
        if message_html:
            email_object.attach(MIMEText(email_html, 'html'))

        for attachment in attachments:
            mime_attachment = MIMEApplication(attachment.content)
            mime_attachment['Content-Disposition'] = 'attachment; filename="%s"' % attachment.filename
            email_object.attach(mime_attachment)

        self.email = email_object

    def send(self):
        server = smtplib.SMTP(self.server, self.port)
        server.starttls()
        server.login(self.send_from, self.password)
        server.sendmail(self.send_from, self.send_to, self.email.as_string())
        server.quit()
