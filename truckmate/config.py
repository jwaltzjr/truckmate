import os
from krc.env import EnvVar

WTF_CSRF_ENABLED = True

SESSION_TYPE = 'filesystem'

DATABASE_NAME = 'STALEY'
DATABASE_HOST = '10.10.81.19'
DATABASE_PORT = '50000'
DATABASE_USER = EnvVar('DBUser').value
DATABASE_PASSWORD = EnvVar('DBPassword').value

SQLALCHEMY_DATABASE_URI = 'db2+ibm_db://{}:{}@{}:{}/{}'.format(
    DATABASE_USER,
    DATABASE_PASSWORD,
    DATABASE_HOST,
    DATABASE_PORT,
    DATABASE_NAME
)

SQLALCHEMY_TRACK_MODIFICATIONS = False

EMAIL_PASSWORD = EnvVar('EmailPassword').value
