import win32com.client as win32
import pandas as pd
from datetime import datetime

def file_path():
    file_path = str(input('File Path: '))
    file_path = file_path.replace('\\', '/')
    file_path = file_path.replace('\'', '')
    file_path = file_path.replace('\"', '')
    
    return file_path


def get_receiver(index, emails_list):
    to = emails_list['Email'].loc[index]
    salutation = str(emails_list['Salutation'].loc[index])
    
    month = int(emails_list['Month'].loc[index])
    day = int(emails_list['Day'].loc[index])
    hour = int(emails_list['Hour'].loc[index]) +2
    minutes = int(emails_list['Minutes'].loc[index])
    delay_date = datetime(2024, month, day, hour, minutes)
    receiver = {'to': to, 'salutation': salutation, 'date': delay_date}
    return receiver


def read_message():
     with open("C:/Users/Baran/Desktop/email/body.html", 'r') as message:
        message_body = message.read()
        return message_body


def send_email(subject, to, body, date):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = subject
    mail.To = to
    mail.HTMLBody = body
    mail.DeferredDeliveryTime = date
    mail.Send()

