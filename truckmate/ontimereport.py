# TODO / PROBLEMS
# TEST SCRIPT WITH DATETIME VALUES OUT OF RANGE AND NULL FOR EACH FIELD

import os
import smtplib
import StringIO
import sys

import openpyxl
import pandas

from email.MIMEMultipart import MIMEMultipart
from email.mime.application import MIMEApplication

import database

REPORT_EMAILS = ['jwaltzjr@krclogistics.com']

def email_spreadsheet(email_addresses, attachments):
    email_username = 'reports@krclogistics.com'
    email_password = 'General1'
    
    # Create email
    email_message = MIMEMultipart('alternative')

    email_message['To'] = ', '.join(email_addresses)
    email_message['From'] = email_username
    email_message['Subject'] = 'On Time Report'

    # Attachments
    for attachment in attachments:
        mime_attachment = MIMEApplication(attachment[1])
        mime_attachment['Content-Disposition'] = 'attachment; filename="%s"' % attachment[0]
        email_message.attach(mime_attachment)

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

    @property
    def data_as_csv(self):
        virtual_csv = StringIO.StringIO()
        self.dataset.to_csv(virtual_csv)
        virtual_csv.seek(0)
        return virtual_csv.getvalue()

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
        insert_data_into_spreadsheet(ws, row, current_column, 'ONTIME_APPT')
        insert_data_into_spreadsheet(ws, row, current_column, 'ONTIME_APPT_REALISTIC', row_offset=13)
        insert_data_into_spreadsheet(ws, row, current_column, 'ONTIME_DELV', row_offset=26)

    style_spreadsheet(ws)

    virtual_wb = openpyxl.writer.excel.save_virtual_workbook(wb)
    return virtual_wb

def insert_titles_into_spreadsheet(worksheet):
    worksheet['A1'] = 'DELIVERY WEEK'

    worksheet['A3'] = 'Appt On Time to RAD'
    worksheet['A4'] = 'STAL-TERM'
    worksheet['A5'] = 'COMM-TERM'
    worksheet['A6'] = 'KELL-TERM'
    worksheet['A7'] = 'LRFD-TERM'
    worksheet['A8'] = 'FKSP-TERM'
    worksheet['A9'] = 'UPPN-TERM'
    worksheet['A10'] = 'UPPN2-TERM'
    worksheet['A11'] = 'HESS-TERM'
    worksheet['A12'] = 'HCIB-TERM'
    worksheet['A13'] = 'BUBM-TERM'
    worksheet['A14'] = 'RINF-TERM'

    worksheet['A16'] = 'Appt On Time to RAD (Realistic)'
    worksheet['A17'] = 'STAL-TERM'
    worksheet['A18'] = 'COMM-TERM'
    worksheet['A19'] = 'KELL-TERM'
    worksheet['A20'] = 'LRFD-TERM'
    worksheet['A21'] = 'FKSP-TERM'
    worksheet['A22'] = 'UPPN-TERM'
    worksheet['A23'] = 'UPPN2-TERM'
    worksheet['A24'] = 'HESS-TERM'
    worksheet['A25'] = 'HCIB-TERM'
    worksheet['A26'] = 'BUBM-TERM'
    worksheet['A27'] = 'RINF-TERM'

    worksheet['A29'] = 'Delv On Time to Appt'
    worksheet['A30'] = 'STAL-TERM'
    worksheet['A31'] = 'COMM-TERM'
    worksheet['A32'] = 'KELL-TERM'
    worksheet['A33'] = 'LRFD-TERM'
    worksheet['A34'] = 'FKSP-TERM'
    worksheet['A35'] = 'UPPN-TERM'
    worksheet['A36'] = 'UPPN2-TERM'
    worksheet['A37'] = 'HESS-TERM'
    worksheet['A38'] = 'HCIB-TERM'
    worksheet['A39'] = 'BUBM-TERM'
    worksheet['A40'] = 'RINF-TERM'

def insert_data_into_spreadsheet(worksheet, ontime_week, column, ontime_field, row_offset=0):
    report_column = {
        'Delivery Week': worksheet.cell(row=1, column=column),
        'STAL-TERM': worksheet.cell(row=4+row_offset, column=column),
        'COMM-TERM': worksheet.cell(row=5+row_offset, column=column),
        'KELL-TERM': worksheet.cell(row=6+row_offset, column=column),
        'LRFD-TERM': worksheet.cell(row=7+row_offset, column=column),
        'FKSP-TERM': worksheet.cell(row=8+row_offset, column=column),
        'UPPN-TERM': worksheet.cell(row=9+row_offset, column=column),
        'UPPN2-TERM': worksheet.cell(row=10+row_offset, column=column),
        'HESS-TERM': worksheet.cell(row=11+row_offset, column=column),
        'HCIB-TERM': worksheet.cell(row=12+row_offset, column=column),
        'BUBM-TERM': worksheet.cell(row=13+row_offset, column=column),
        'RINF-TERM': worksheet.cell(row=14+row_offset, column=column),
    }

    for key, cell in report_column.iteritems():
        if key != 'Delivery Week':
            cell.number_format = '0.00%'

    current_terminal = ontime_week.Index[1].strip()

    report_column['Delivery Week'].value = ontime_week.Index[0]
    report_column[current_terminal].value = getattr(ontime_week, ontime_field)

def style_spreadsheet(worksheet):
    for spreadsheet_section in ['A', 3, 16, 29]:
        for cell in worksheet[spreadsheet_section]:
            cell.font = cell.font.copy(bold=True)
    for spreadsheet_cell in ['A3', 'A16', 'A29']:
        worksheet[spreadsheet_cell].font = worksheet[spreadsheet_cell].font.copy(underline='single')

ontime_report = OnTimeReport('ontimereport.sql', database.truckmate)
ontime_avg_dataset = ontime_report.get_dataset_of_averages()

report_file = create_report(ontime_avg_dataset)
email_spreadsheet(
    REPORT_EMAILS,
    [
        ('on_time_report.xlsx', report_file),
        ('on_time_report_data.csv', ontime_report.data_as_csv)
    ]
)
