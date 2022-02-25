import openpyxl as xl
from os import listdir
from os.path import isfile, join
from common_functions import CommonFunctions
from woocommerce import API
import base64, requests, json
from datetime import datetime
import re
import os
import yagmail
from datetime import datetime, date

excel_path = r"/Users/muralidharpettela/Downloads/Test.xlsx"
from datetime import date
from dateutil.relativedelta import relativedelta

one_months = date.today() + relativedelta(months=+1)
two_months = date.today() + relativedelta(months=+2)
three_months = date.today() + relativedelta(months=+3)

one_month_expiry_list = list()
two_month_expiry_list = list()
three_month_expiry_list = list()
products_already_expired_list = list()

new_products_workbook = xl.load_workbook(excel_path)
new_products_ws = new_products_workbook.worksheets[0]
# column of sort id source file
col_name = new_products_ws['A'][1:]
# category data
col_category = new_products_ws['B'][1:]


def send_email(subject, content):
    user = 'lotusgroceryingolstadt2@gmail.com'
    app_password = 'ofanobixbzpvcpqz'  # a token for gmail
    to = 'info@lotus-grocery.eu'  # To send a group of recipients, simply change ‘to’ to a list.

    subject = subject
    content = content

    with yagmail.SMTP(user, app_password) as yag:
        yag.send(to, subject, content)
        print('Sent email successfully')


for col_nam, expiry_date in zip(col_name, col_category):
    # one month
    if (expiry_date.value.date() <= one_months) and (expiry_date.value.date() >= datetime.now().date()):
        one_month_expiry_list.append(col_nam.value)
    # two months
    if (expiry_date.value.date() <= two_months) and (expiry_date.value.date() >= datetime.now().date()):
        two_month_expiry_list.append(col_nam.value)
    # three months
    if (expiry_date.value.date() <= three_months) and (expiry_date.value.date() >= datetime.now().date()):
        three_month_expiry_list.append(col_nam.value)
    # products already expired
    if expiry_date.value.date() <= datetime.now().date():
        products_already_expired_list.append(col_nam.value)
         # print("products already expired {}".format(col_nam.value))

message = "This is an automated mail, notification of the expiry products within 3 months. \n " \
          "Products expiring in 1 month:\n{}\n{}\n{}\n" \
          "Products expiring in 2 months:\n{}\n{}\n{}\n" \
          "Products expiring in 3 month:\n{}\n{}\n{}\n" \
          "Products already expired:\n{}\n{}\n{}\n".format("==================================", "\n".join(one_month_expiry_list), "==================================",
                                                         "==================================", "\n".join(two_month_expiry_list), "==================================",
                                                         "==================================", "\n".join(three_month_expiry_list), "==================================",
                                                         "==================================", "\n".join(products_already_expired_list), "==================================")
subject = '[Notification] lotus-grocery.eu - Expiry Products in next 3 months ' + datetime.now().strftime("%d/%m/%Y %H:%M:%S")
content = [message]
send_email(subject, content)
print("Products expiring in 1 month:\n{}\n{}\n{}" .format("==================================", "\n".join(one_month_expiry_list), "=================================="))
print("Products expiring in 2 months:\n{}\n{}\n{}" .format("==================================", "\n".join(two_month_expiry_list), "=================================="))
print("Products expiring in 3 month:\n{}\n{}\n{}" .format("==================================", "\n".join(three_month_expiry_list), "=================================="))
print("Products already expired:\n{}\n{}\n{}" .format("==================================", "\n".join(products_already_expired_list), "=================================="))
