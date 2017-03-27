# TODO / PROBLEMS
# TEST SCRIPT WITH DATETIME VALUES OUT OF RANGE AND NULL FOR EACH FIELD

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
    email_message['Subject'] = 'On Time Report'

    # Attach Spreadsheet
    attachment = MIMEApplication(spreadsheet)
    attachment['Content-Disposition'] = 'attachment; filename="%s"' % 'on_time_report.xlsx'
    email_message.attach(attachment)

    # Connect to server and send email
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    server.login(email_username, email_password)
    server.sendmail(email_username, email_addresses, email_message.as_string())
    server.quit()

class OnTimeReport(object):

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

    def __init__(self, file_name, datab):
        self.sql_query = self.load_query_from_file(os.path.join(sys.path[0], file_name))
        self.dataset = self.fetch_data_from_db(self.sql_query, datab)
        self.apply_calculated_columns()

    def load_query_from_file(self, file_path):
        with open(file_path, 'r') as sql_file:
            return sql_file.read()

    def fetch_data_from_db(self, sql_query, datab):
        with datab as db:
            dataset = pandas.read_sql(
                sql_query,
                db.connection
            )
        return dataset

    def apply_calculated_columns(self):
        self.dataset['ONTIME_APPT'] = self.dataset.apply(
            lambda row: self.__class__.CalcColumns.ontime_appt(
                row['DELIVER_BY'].date(),
                row['RAD']
            ),
            axis = 1
        )
        self.dataset['ONTIME_APPT_REALISTIC'] = self.dataset.apply(
            lambda row: self.__class__.CalcColumns.ontime_appt_realistic(
                row['RAD'],
                row['RPD'],
                row['CREATED_TIME'].date(),
                row['DELIVER_BY'].date()
            ),
            axis = 1
        )
        self.dataset['ONTIME_DELV'] = self.dataset.apply(
            lambda row: self.__class__.CalcColumns.ontime_delv(
                row['ARRIVED'],
                row['DELIVER_BY_END']
            ),
            axis = 1
        )

    def get_dataset_of_averages(self):
        numeric_dataset = self.dataset[
            [
                'DELIVERY_WEEK',
                'DELIVERY_TERMINAL',
                'ONTIME_APPT',
                'ONTIME_APPT_REALISTIC',
                'ONTIME_DELV'
            ]
        ].apply(pandas.to_numeric, errors='ignore')

        averages = numeric_dataset.groupby(
            ['DELIVERY_WEEK', 'DELIVERY_TERMINAL']
        )[['ONTIME_APPT', 'ONTIME_APPT_REALISTIC', 'ONTIME_DELV']].mean()

        return averages

def create_report(data):
    wb = openpyxl.Workbook()
    ws = wb.active

    insert_titles_into_spreadsheet(ws)

    current_column = 1
    current_date = None
    for row in data.itertuples():
        if row.Index[0] != current_date:
            current_column += 1
        current_date = row.Index[0]
        insert_row_into_spreadsheet(ws, row, current_column)

    virtual_wb = openpyxl.writer.excel.save_virtual_workbook(wb)
    return virtual_wb

def insert_titles_into_spreadsheet(worksheet):
    worksheet['A1'] = 'DELIVERY WEEK'

    worksheet['A3'] = 'STAL-TERM'
    worksheet['A4'] = 'COMM-TERM'
    worksheet['A5'] = 'KELL-TERM'
    worksheet['A6'] = 'LRFD-TERM'
    worksheet['A7'] = 'FKSP-TERM'
    worksheet['A8'] = 'UPPN-TERM'
    worksheet['A9'] = 'UPPN2-TERM'
    worksheet['A10'] = 'HESS-TERM'
    worksheet['A11'] = 'HCIB-TERM'
    worksheet['A12'] = 'BUBM-TERM'
    worksheet['A13'] = 'RINF-TERM'

def insert_row_into_spreadsheet(worksheet, ontime_week, column):
    report_column = {
        'Delivery Week': worksheet.cell(row=1, column=column),
        'STAL-TERM': worksheet.cell(row=3, column=column),
        'COMM-TERM': worksheet.cell(row=4, column=column),
        'KELL-TERM': worksheet.cell(row=5, column=column),
        'LRFD-TERM': worksheet.cell(row=6, column=column),
        'FKSP-TERM': worksheet.cell(row=7, column=column),
        'UPPN-TERM': worksheet.cell(row=8, column=column),
        'UPPN2-TERM': worksheet.cell(row=9, column=column),
        'HESS-TERM': worksheet.cell(row=10, column=column),
        'HCIB-TERM': worksheet.cell(row=11, column=column),
        'BUBM-TERM': worksheet.cell(row=12, column=column),
        'RINF-TERM': worksheet.cell(row=13, column=column)
    }

    for key, cell in report_column.iteritems():
        if key != 'Delivery Week':
            cell.number_format = '0.00%'

    current_terminal = ontime_week.Index[1].strip()

    report_column['Delivery Week'].value = ontime_week.Index[0]
    report_column[current_terminal].value = ontime_week.ONTIME_APPT

ontime_report = OnTimeReport('ontimereport.sql', database.truckmate)
ontime_avg_dataset = ontime_report.get_dataset_of_averages()

report_file = create_report(ontime_avg_dataset)
email_spreadsheet(REPORT_EMAILS, report_file)
