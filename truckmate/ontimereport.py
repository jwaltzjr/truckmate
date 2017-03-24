# TODO / PROBLEMS
# CONCAT OF FINAL DATAFRAMES CONTAINS LOTS OF NaN VALUES

import os
import smtplib
import sys

import openpyxl
import pandas

from email.MIMEMultipart import MIMEMultipart
from email.mime.application import MIMEApplication

import database

REPORT_EMAILS = ['jwaltzjr@krclogistics.com']

class CalcColumns(object):

    @staticmethod
    def ontime_appt(delivery_date, rad):
        if rad:
            return delivery_date <= rad
        else:
            return True # Always on time if no due date

    @staticmethod
    def ontime_appt_realistic(rad, rpd, created_date, deliver_by):
        if rad and rpd and (rad > rpd) & (rpd >= created_date):
            return deliver_by <= rad
        elif not rad:
            return True
        else:
            return None

    @staticmethod
    def ontime_delv(arrived, deliver_by_end):
        if arrived:
            return arrived <= deliver_by_end
        else:
            return None

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

sql_file_path = os.path.join(sys.path[0], 'ontimereport.sql')
with open(sql_file_path, 'r') as sql_file:
    sql_query = sql_file.read()

with database.truckmate as db:
    dataset = pandas.read_sql(
        sql_query,
        db.connection
        # parse_dates = {
        #     'RECEIVED': '%Y-%m-%d',
        #     'RAD': '%Y-%m-%d',
        #     'RPD': '%Y-%m-%d',
        #     'DELIVER_BY': '%Y-%m-%d-%H.%M.%S.%f',
        #     'DELIVER_BY_END': '%Y-%m-%d-%H.%M.%S.%f',
        #     'CREATED_TIME': '%Y-%m-%d-%H.%M.%S.%f',
        #     'ARRIVED': '%Y-%m-%d-%H.%M.%S.%f'
        # }
    )

dataset['ONTIME_APPT'] = dataset.apply(
    lambda row: CalcColumns.ontime_appt(
        row['DELIVER_BY'].date(),
        row['RAD']
    ),
    axis = 1
)

dataset['ONTIME_APPT_REALISTIC'] = dataset.apply(
    lambda row: CalcColumns.ontime_appt_realistic(
        row['RAD'],
        row['RPD'],
        row['CREATED_TIME'].date(),
        row['DELIVER_BY'].date()
    ),
    axis = 1
)

dataset['ONTIME_DELV'] = dataset.apply(
    lambda row: CalcColumns.ontime_delv(
        row['ARRIVED'],
        row['DELIVER_BY_END']
    ),
    axis = 1
)

dataset2 = dataset.dropna(subset = ['ONTIME_APPT_REALISTIC']).copy()
dataset2['ONTIME_APPT_REALISTIC'] = dataset2['ONTIME_APPT_REALISTIC'].astype('bool')
 
dataset3 = dataset.dropna(subset = ['ONTIME_DELV']).copy()
dataset3['ONTIME_DELV'] = dataset3['ONTIME_DELV'].astype('bool')
 
groupby = dataset.groupby(['DELIVERY_WEEK','DELIVERY_TERMINAL'])[['ONTIME_APPT']].mean()
groupby2 = dataset2.groupby(['DELIVERY_WEEK','DELIVERY_TERMINAL'])['ONTIME_APPT_REALISTIC'].mean()
groupby3 = dataset3.groupby(['DELIVERY_WEEK','DELIVERY_TERMINAL'])['ONTIME_DELV'].mean()

print groupby
print groupby2
print groupby3

final_result = pandas.concat([groupby, groupby2, groupby3], axis=1)

print final_result