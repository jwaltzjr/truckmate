import collections
import os
import sys

import openpyxl

import database
from truckmate_email import TruckmateEmail

REPORT_EMAILS = [
    'jwaltzjr@krclogistics.com'
]

class Rate(object):

    def __init__(self, tariff, customers, origin, destination, rate_break, rate):
        self.tariff = tariff
        self.customers = customers
        self.origin = origin
        self.destination = destination
        self.rate_break = rate_break
        self.rate = rate

    @property
    def three_digit_zip(self):
        if self.destination.isdigit():
            if 600 <= self.destination[:3] <= 606:
                return 'CHICOMM'
            else:
                return self.destination[:3]
        else:
            return self.destination

class RateReport(object):

    def __init__(self, file_name, datab):
        sql_file_path = os.path.join(sys.path[0], file_name)
        self.sql_query = self.load_query_from_file(sql_file_path)
        self.dataset = self.fetch_data_from_db(self.sql_query, datab)
        self.split_data = self.split_dataset(self.dataset)

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

        self._excel_insert_titles(ws)

        current_column = 2
        for rate_cell in self.dataset:
            self._excel_insert_data(ws, rate_cell, current_column)
            current_column += 1

        self._excel_apply_styling(ws)

        virtual_wb = openpyxl.writer.excel.save_virtual_workbook(wb)
        return virtual_wb

    def split_dataset(self, dataset):
        split_data = collections.defaultdict(list)

        for rate in dataset:
            for origin in self.get_origins(rate):
                rate_obj = Rate(rate.TARIFF, rate.CUSTOMERS, origin, rate.DESTINATION, rate.BREAK, rate.RATE)
                split_data[str(rate_obj.three_digit_zip)].append(rate_obj)

        return split_data

    def get_origins(self, rate):
        origins = []

        if rate.ORIGIN_MS:
            for origin in rate.ORIGIN_MS.split(', '):
                origins.append(origin)

        if rate.ORIGIN:
            origins.append(rate.ORIGIN)

        return origins

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

def test():
    rate_report = RateReport('ratereport.sql', database.truckmate)
    print rate_report.split_data['432']

if __name__ == '__main__':
    test()
