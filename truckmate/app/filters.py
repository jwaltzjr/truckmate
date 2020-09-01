from datetime import datetime

from app import app

@app.template_filter()
def get_date(cdf):
    if cdf:
        return cdf.DATE.strftime('%m/%d/%y %H:%M:%S')
    else:
        return ''