import datetime
import smtplib
from datetime import date, timedelta

date_start = datetime.datetime.today().replace(hour=0, minute=1, second=0, microsecond=0)
date_yesterday = date_start - timedelta(days=1)
date_end = datetime.datetime.today().replace(hour=23, minute=59, second=0, microsecond=0)

filepath = "D:\\certs\\time.xlsx"

server = smtplib.SMTP('mail.1mf.ru', 25)

addr_from = "stork@1mf.ru"
addr_to = ['o.burmistrov@1mf.ru']
addr_cc = ['t.sharuk@1mf.ru', 't.kazanskaya@1mf.ru']