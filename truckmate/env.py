from krc.env import EnvVar
 
def setup():
    email_password.value = raw_input('Email Password: ')
    db_user.value = raw_input('Database User: ')
    db_password.value = raw_input('Database Password: ')

email_password = EnvVar('EmailPassword')
db_user = EnvVar('DBUser')
db_password = EnvVar('DBPassword')

if __name__ == '__main__':
    setup()
