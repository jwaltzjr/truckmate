import pyodbc
import env
from krc.database import DatabaseConnection

truckmate = DatabaseConnection(
    '{IBM DB2 ODBC DRIVER - DB2COPY1}',
    'TM_Reporting_00001',
    'STALEY',
    env.db_user.value,
    env.db_password.value
)
