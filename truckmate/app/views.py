from app import app, db
from . import models, forms

@app.route('/', methods=['GET','POST'])
def index():
	return 'test index'