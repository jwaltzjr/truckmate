import pyodbc
import env
from krc.database import DatabaseConnection

truckmate = DatabaseConnection(
    '{IBM DB2 ODBC DRIVER - DB2COPY1}',
    'TM_Reporting_SVC_00001',
    'STALEY',
    env.db_user.value,
    env.db_password.value,
    hostname='STAY-DB201',
    port='50000'
)

if __name__ == '__main__':
    with truckmate as db:
        print('Connected successfully')
        with db.connection.cursor() as cursor:
            print('Cursor opened')
            cursor.execute('SELECT * FROM TLORDER FETCH FIRST ROW ONLY')
            print('Cursor result:\n%s' % cursor.fetchall())
        print('Cursor commited or rolled back')
    print('Connection closed successfully')
            
