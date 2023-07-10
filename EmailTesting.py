import os
import zipfile
import glob
import smtplib
from email.message import EmailMessage
import xml.etree.ElementTree as ET
import logging
import datetime
import shutil


def send_email():
    email_address = 'kelvin.chu1122@gmail.com'
    email_app_password = 'brwgearofnryfinh'
    msg = EmailMessage()
    msg['Subject'] = 'FSD PDF file'
    msg['From'] = email_address
    msg['To'] = 'kelvin.fsd.testing@gmail.com'

    msg.set_content('Testing.')

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email_address, email_app_password)
        smtp.send_message(msg)
        logging.info(f'{datetime.datetime.now()}: successful')


def main():
    send_email()


if __name__ == "__main__":
    main()
