from flask import Flask
from flask_session import Session
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.config.from_object('config')

Session(app)

db = SQLAlchemy(app)
db.Model.metadata.reflect(
    db.engine,
    schema='TMWIN',
    only=[
        'tlorder',
        'tlorder_term_plan',
        'trip',
        'client',
        'trace',
        'custom_data'
    ]
)

from app import views, models, forms, filters