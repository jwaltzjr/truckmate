import os
import sys

import openpyxl

import database
from truckmate_email import TruckmateEmail

REPORT_EMAILS = [
    'jwaltzjr@krclogistics.com'
]

class RateReport(object):

    def __init__(self, file_name, datab):
        sql_file_path = os.path.join(sys.path[0], file_name)
        self.sql_query = self.load_query_from_file(sql_file_path)
        self.dataset = self.fetch_data_from_db(self.sql_query, datab)

    def load_query_from_file(self, file_path):
        with open(file_path, 'r') as sql_file:
            return sql_file.read()

    def fetch_data_from_db(self, query, db):
        with db as datab:
            with datab.connection.cursor() as cursor:
                cursor.execute(query)
                return cursor.fetchall()

    def export_as_xlsx(self):
        wb = openpyxl.Workbook()
        ws = wb.active

        split_data = self.split_dataset()

        self._excel_insert_titles(ws)

        current_column = 2
        for rate_cell in self.dataset:
            self._excel_insert_data(ws, rate_cell, current_column)
            current_column += 1

        self._excel_apply_styling(ws)

        virtual_wb = openpyxl.writer.excel.save_virtual_workbook(wb)
        return virtual_wb

    def split_dataset(self):
        split_data = {}

        # for rate in self.dataset:

    def get_zone(self, rate):
        if rate.DESTINATION.isdigit():
            if 600 <= rate.DESINATION[:3] <= 606:
                return 'CHICOMM'
            else:
                return rate.DESINATION[:3]
        else:
            return rate.DESTINATION

    def _excel_insert_titles(self, worksheet):
        titles = {}

        for cell, title in titles.items():
            worksheet[cell] = title

    def _excel_insert_data(self, worksheet, rate_cell, column):
        # INSERT DATA HERE
        return

    def _excel_apply_styling(self, worksheet):
        # APPLY STYLING HERE
        return

def main():
    rate_report = RateReport('ratereport.sql', database.truckmate)
    email_message = TruckmateEmail(
        REPORT_EMAILS,
        subject='Rate Report',
        attachments=[
            ('rate_report.xlsx', tonnage_report.export_as_xlsx())
        ]
    )
    email_message.send()

if __name__ == '__main__':
    main()
