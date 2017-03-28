import os
import sys

import openpyxl

import database
import krcemail

REPORT_EMAILS = [
    'jwaltzjr@krclogistics.com',
    'jwaltz@krclogistics.com',
    'dhendriksen@krclogistics.com',
    'djdevries@krclogistics.com',
    'tkatsahnias@krclogistics.com',
    'ekuhowski@krclogistics.com',
    'dpeach@krclogistics.com'
]

class TonnageReport(object):

    def __init__(self):
        sql_file_path = os.path.join(sys.path[0], 'tonnage.sql')
        self.sql_query = self.load_query_from_file(sql_file_path)

        self.dataset = self.fetch_data_from_db(database.truckmate, self.sql_query)

    def load_query_from_file(self, file_path):
        with open(file_path, 'r') as sql_file:
            return sql_file.read()

    def fetch_data_from_db(self, db, query):
        with db as datab:
            with datab.connection.cursor() as cursor:
                cursor.execute(query)
                return cursor.fetchall()

    def export_as_xlsx(self):
        wb = openpyxl.Workbook()
        ws = wb.active

        self._excel_insert_titles(ws)

        current_column = 2
        for tonnage_week in self.dataset:
            self._excel_insert_data(ws, tonnage_week, current_column)
            current_column += 1

        self._excel_apply_styling(ws)

        virtual_wb = openpyxl.writer.excel.save_virtual_workbook(wb)
        return virtual_wb

    def _excel_insert_titles(self, worksheet):
        worksheet['A1'] = 'DELIVERY WEEK'

        worksheet['A3'] = 'WEIGHT 10'
        worksheet['A4'] = 'WEIGHT 11'
        worksheet['A5'] = 'WEIGHT 12'
        worksheet['A6'] = 'WEIGHT 13'
        worksheet['A7'] = 'WEIGHT 14'
        worksheet['A8'] = 'WEIGHT 15'
        worksheet['A9'] = 'WEIGHT TOTAL'

        worksheet['A12'] = '# ORDERS 10'
        worksheet['A13'] = '# ORDERS 11'
        worksheet['A14'] = '# ORDERS 12'
        worksheet['A15'] = '# ORDERS 13'
        worksheet['A16'] = '# ORDERS 14'
        worksheet['A17'] = '# ORDERS 15'
        worksheet['A18'] = '# ORDERS TOTAL'

        worksheet['A21'] = 'AVG WEIGHT 10'
        worksheet['A22'] = 'AVG WEIGHT 11'
        worksheet['A23'] = 'AVG WEIGHT 12'
        worksheet['A24'] = 'AVG WEIGHT 13'
        worksheet['A25'] = 'AVG WEIGHT 14'
        worksheet['A26'] = 'AVG WEIGHT 15'
        worksheet['A27'] = 'AVG WEIGHT TOTAL'

        worksheet['A30'] = 'WEIGHT UNDEF'
        worksheet['A31'] = '# ORDERS UNDEF'
        worksheet['A32'] = 'AVG WEIGHT UNDEF'

    def _excel_insert_data(self, worksheet, tonnage_week, column):
        tonnage_week_column = {
            'title': worksheet.cell(row=1, column=column),
            'weight_10': worksheet.cell(row=3, column=column),
            'weight_11': worksheet.cell(row=4, column=column),
            'weight_12': worksheet.cell(row=5, column=column),
            'weight_13': worksheet.cell(row=6, column=column),
            'weight_14': worksheet.cell(row=7, column=column),
            'weight_15': worksheet.cell(row=8, column=column),
            'weight_total': worksheet.cell(row=9, column=column),
            'num_orders_10': worksheet.cell(row=12, column=column),
            'num_orders_11': worksheet.cell(row=13, column=column),
            'num_orders_12': worksheet.cell(row=14, column=column),
            'num_orders_13': worksheet.cell(row=15, column=column),
            'num_orders_14': worksheet.cell(row=16, column=column),
            'num_orders_15': worksheet.cell(row=17, column=column),
            'num_orders_total': worksheet.cell(row=18, column=column),
            'avg_weight_10': worksheet.cell(row=21, column=column),
            'avg_weight_11': worksheet.cell(row=22, column=column),
            'avg_weight_12': worksheet.cell(row=23, column=column),
            'avg_weight_13': worksheet.cell(row=24, column=column),
            'avg_weight_14': worksheet.cell(row=25, column=column),
            'avg_weight_15': worksheet.cell(row=26, column=column),
            'avg_weight_total': worksheet.cell(row=27, column=column),
            'weight_undef': worksheet.cell(row=30, column=column),
            'num_orders_undef': worksheet.cell(row=31, column=column),
            'avg_weight_undef': worksheet.cell(row=32, column=column)
        }

        tonnage_week_column['title'].value = tonnage_week.DELIVERY_WEEK

        tonnage_week_column['weight_10'].value = tonnage_week.WEIGHT_10 or 0
        tonnage_week_column['weight_11'].value = tonnage_week.WEIGHT_11 or 0
        tonnage_week_column['weight_12'].value = tonnage_week.WEIGHT_12 or 0
        tonnage_week_column['weight_13'].value = tonnage_week.WEIGHT_13 or 0
        tonnage_week_column['weight_14'].value = tonnage_week.WEIGHT_14 or 0
        tonnage_week_column['weight_15'].value = tonnage_week.WEIGHT_15 or 0
        tonnage_week_column['weight_total'].value = tonnage_week.WEIGHT or 0

        tonnage_week_column['num_orders_10'].value = tonnage_week.NUM_ORDERS_10 or 0
        tonnage_week_column['num_orders_11'].value = tonnage_week.NUM_ORDERS_11 or 0
        tonnage_week_column['num_orders_12'].value = tonnage_week.NUM_ORDERS_12 or 0
        tonnage_week_column['num_orders_13'].value = tonnage_week.NUM_ORDERS_13 or 0
        tonnage_week_column['num_orders_14'].value = tonnage_week.NUM_ORDERS_14 or 0
        tonnage_week_column['num_orders_15'].value = tonnage_week.NUM_ORDERS_15 or 0
        tonnage_week_column['num_orders_total'].value = tonnage_week.NUM_ORDERS or 0

        tonnage_week_column['avg_weight_10'].value = tonnage_week.AVG_WEIGHT_10 or 0
        tonnage_week_column['avg_weight_11'].value = tonnage_week.AVG_WEIGHT_11 or 0
        tonnage_week_column['avg_weight_12'].value = tonnage_week.AVG_WEIGHT_12 or 0
        tonnage_week_column['avg_weight_13'].value = tonnage_week.AVG_WEIGHT_13 or 0
        tonnage_week_column['avg_weight_14'].value = tonnage_week.AVG_WEIGHT_14 or 0
        tonnage_week_column['avg_weight_15'].value = tonnage_week.AVG_WEIGHT_15 or 0
        tonnage_week_column['avg_weight_total'].value = tonnage_week.AVG_WEIGHT or 0

        tonnage_week_column['weight_undef'].value = tonnage_week.WEIGHT_UNDEF or 0
        tonnage_week_column['num_orders_undef'].value = tonnage_week.NUM_ORDERS_UNDEF or 0
        tonnage_week_column['avg_weight_undef'].value = tonnage_week.AVG_WEIGHT_UNDEF or 0

        for key, cell in tonnage_week_column.iteritems():
            if key != 'title':
                cell.number_format = '#,##0'

    def _excel_apply_styling(self, worksheet):
        for spreadsheet_section in ['A', 9, 18, 27]:
            for cell in worksheet[spreadsheet_section]:
                cell.font = cell.font.copy(bold=True)
        for cell in ['A1', 'A8', 'A17', 'A26']:
            worksheet[cell].font = worksheet[cell].font.copy(underline='single')

def main():
    tonnage_report = TonnageReport()
    email_message = krcemail.KrcEmail(
        REPORT_EMAILS,
        subject='Weekly Tonnage',
        attachments=[
            ('weekly_tonnage.xlsx', tonnage_report.export_as_xlsx())
        ]
    )
    email_message.send()

if __name__ == '__main__':
    main()
