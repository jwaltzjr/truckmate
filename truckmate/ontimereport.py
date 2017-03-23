import os
import smtplib
import sys

import openpyxl
import pandas

from email.MIMEMultipart import MIMEMultipart
from email.mime.application import MIMEApplication

import database

REPORT_EMAILS = ['jwaltzjr@krclogistics.com']

def email_spreadsheet(email_addresses, spreadsheet):
    email_username = 'reports@krclogistics.com'
    email_password = 'General1'
    
    # Create email
    email_message = MIMEMultipart('alternative')

    email_message['To'] = ', '.join(email_addresses)
    email_message['From'] = email_username
    email_message['Subject'] = 'Weekly Tonnage'

    # Attach Spreadsheet
    attachment = MIMEApplication(spreadsheet)
    attachment['Content-Disposition'] = 'attachment; filename="%s"' % 'weekly_tonnage.xlsx'
    email_message.attach(attachment)

    # Connect to server and send email
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(email_username, email_password)
    server.sendmail(email_username, email_addresses, email_message.as_string())
    server.quit()

def main():
    sql_file_path = os.path.join(sys.path[0], 'ontimereport.sql')
    with open(sql_file_path, 'r') as sql_file:
        sql_query = sql_file.read()

    with database.truckmate as db:
        dataset = pandas.read_sql(sql_query, db.connection)
    
    dataset['RAD'] = pandas.to_datetime(dataset['RAD'])
    dataset['RPD'] = pandas.to_datetime(dataset['RPD'])

    dataset['ONTIME_APPT'] = dataset['DELIVER_BY'].dt.date < dataset['RAD']

    if (dataset['RAD'] > dataset['RPD']) & (dataset['RPD'] > dataset['CREATED_TIME']):
        dataset['ONTIME_APPT_REALISTIC'] = DATASET['ONTIME_APPT']

    if dataset['ARRIVED']:
        dataset['ONTIME_DELV'] = dataset['ARRIVED'] < dataset['DELIVER_BY_END']

    print dataset

if __name__ == '__main__':
    main()
