import pyodbc

class DatabaseConnection(object):

    def __init__(self, odbc_driver, system, database, user, password):
        self.odbc_driver = odbc_driver
        self.system = system
        self.database = database
        self.user = user
        self.password = password

    def __repr__(self):
        return '{class_name}(system={system}, database={database}, user={user})'.format(
            class_name = self.__class__.__name__,
            system = self.system,
            database = self.database,
            user = self.user
        )

    def __enter__(self):
        self.connection = pyodbc.connect(
            'Driver={};System={};DATABASE={};Uid={};Pwd={}'.format(
                self.odbc_driver, self.system, self.database, self.user, self.password
            )
        )
        return self

    def __exit__(self, type, value, traceback):
        self.connection.close()

truckmate = DatabaseConnection(
    '{IBM DB2 ODBC DRIVER - DB2COPY1}',
    'TM_Reporting_00001',
    'STALEY',
    'TMW_SCRIPTING',
    'pY2$t%r\L49f`\QLk'
)