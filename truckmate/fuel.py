import collections
import datetime
import urllib

# Third-party
import pyodbc
import xlrd

import database
import krcemail

ERROR_EMAIL_ADDRESSES = ['jwaltzjr@krclogistics.com', 'csenti@krclogistics.com']
DB2_DATABASE = 'STALEY'

FuelPrice = collections.namedtuple(
    'FuelPrice',
    ('date', 'price')
)

class FuelSheet(object):

    FUEL_SHEET_URL = 'https://www.eia.gov/petroleum/gasdiesel/xls/psw18vwall.xls'

    def __init__(self):
        self.spreadsheet = self.download_file(FuelSheet.FUEL_SHEET_URL).read()

    def __repr__(self):
        return '{class_name}(current_fuel={current_fuel})'.format(
            class_name = self.__class__.__name__,
            current_fuel = str(self.current_fuel)
        )

    def download_file(self, url):
        return urllib.urlopen(url)

    def unpack_excel_date(self, date, workbook):
        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(date, workbook.datemode)
        return datetime.date(year, month, day)

    @property
    def current_fuel(self):
        wb = xlrd.open_workbook(file_contents=self.spreadsheet)
        ws = wb.sheet_by_name('Data 1')
        newest_row = ws.row_values(ws.nrows-1)

        date = self.unpack_excel_date(newest_row[0], wb)
        price = newest_row[1]
        return FuelPrice(date, price)

    @property
    def is_recent(self, days=6):
        today = datetime.date.today()
        start_date = today - datetime.timedelta(days)
        return start_date <= self.current_fuel.date <= today

class FuelAverage(object):

    def __init__(self, name, days_after_doe):
        self.name = name
        self.days_after_doe = days_after_doe

    def __repr__(self):
        return '{class_name}(name={name}, days_after_doe={days_after_doe})'.format(
            class_name = self.__class__.__name__,
            name = self.name,
            days_after_doe = self.days_after_doe
        )

    def insert_into_db(self, cursor, fuel_average):
        insert_statement = """
            INSERT INTO TMWIN.FP_WEEKLY_PRICE
                (AVERAGE,START_DATE,PRICE)
            VALUES
                (?,?,?)
        """
        fuel_date = fuel_average.date + datetime.timedelta(days=self.days_after_doe)
        cursor.execute(insert_statement, self.name, fuel_date, fuel_average.price)

class CalculatedFuelAverage(FuelAverage):

    def insert_into_db(self, cursor, fuel_average):
        insert_statement = """
            INSERT INTO TMWIN.ACHARGE_DETAIL
                (ACD_ID,ACODE_ID,ACD_FACTOR,ACD_PERCENT,ACD_MINIMUM,ACD_MAXIMUM,ACD_RANGE_FROM,ACD_RANGE_TO,ACD_FIELD_ID,ACD_START_DATE,ACD_END_DATE,VENDOR_ID,FP_AVERAGE,CLIENT_ID,START_ZONE,END_ZONE,ALLOW_BETWEEN,RATE_SHEET_ID,CALC_SEQ,SERVICE_LEVEL,CUST_MIN,CUST_MAX,DOE_FAC_A,DOE_FAC_B,COMMODITY,STAIRS,ELEVATOR,DOCK,SITE_SURVEY,VEHICLE_RESTRICT,USER_COND,STATUS_CODE,IM_SIZE,IM_TYPE,IM_ISO,IM_MOVEMENT,INSTRUCTION_ID,MOVEMENT_TYPE,THRESHOLD_AMOUNT,VEN_ES_EXCLUDE_SAME_ZONE,IM_POOL,USER_COND_TABLE,ROW_TIMESTAMP)
            VALUES
                (NEXT VALUE FOR TMWIN.GEN_ACD_ID,?,0,?,0,0,0,5,8,?,?,'','DOE-US','','','','True',0,0,NULL,0,0,NULL,NULL,NULL,'False','False','False','False','False','','',NULL,NULL,NULL,NULL,0,'*',0,'False',NULL,'TLORDER',CURRENT TIMESTAMP)
        """
        surcharge = self.calculate_fuel_surcharge(fuel_average)
        fuel_start_date = fuel_average.date + datetime.timedelta(days=self.days_after_doe)
        fuel_end_date = datetime.datetime.combine(
            fuel_start_date + datetime.timedelta(days=6),
            datetime.time(hour=23, minute=59, second=59, microsecond=999000)
        )
        cursor.execute(insert_statement, self.name, surcharge, fuel_start_date, fuel_end_date)

    def calculate_fuel_surcharge(self, fuel_price):
        raise NotImplementedError('Fuel surcharge calculation for {} was not implemented.'.format(self.name))

class MarsFuelAverage(CalculatedFuelAverage):

    def calculate_fuel_surcharge(self, fuel_price):
        return (round((fuel_price.price-1.08)/5,2)/2)*100

def email_error_message(error_message, email_addresses):
    email_body = 'There was an error with the automatic fuel insert.\n\nError Message:\n{}\n\nITDEPREQ'    
    email_message = krcemail.KrcEmail(
        email_addresses,
        subject='Automatic Fuel Failure',
        message=email_body.format(error_message)
    )
    email_message.send()

def insert_fuel_into_database(database, fuel_averages, fuel_spreadsheet):
    with database as db:
        with db.connection.cursor() as cursor:
            for fuel_average in fuel_averages:
                fuel_average.insert_into_db(cursor, fuel_spreadsheet.current_fuel)

def main(fuel_averages):
    db = database.truckmate
    db.database = DB2_DATABASE

    fuel_spreadsheet = FuelSheet()
    if fuel_spreadsheet.is_recent:
        try:
            insert_fuel_into_database(db, fuel_averages, fuel_spreadsheet)

        except pyodbc.IntegrityError as error:
            error_message = 'One or more fuel surcharges have already been entered. {}'.format(error)
            email_error_message(error_message, ERROR_EMAIL_ADDRESSES)

        except BaseException as error:
            error_message = 'An unidentified error occured. {}'.format(error)
            email_error_message(error_message, ERROR_EMAIL_ADDRESSES)
    else:
        error_message = 'No fuel record found for this week.'
        email_error_message(error_message, ERROR_EMAIL_ADDRESSES)

fuel_averages = (
    FuelAverage('DOE-US', 1),
    FuelAverage('DOE-LWUS', 8),
    FuelAverage('DOE-2WUS', 15),
    FuelAverage('DOE-US-MON', 0),
    FuelAverage('DOE-US-SAT', 5),
    MarsFuelAverage('FSC-MARS', 1)
)

if __name__ == '__main__':
    main(fuel_averages)
