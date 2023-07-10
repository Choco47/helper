import os
import zipfile
import glob
import smtplib
from email.message import EmailMessage
import xml.etree.ElementTree as ET
import logging
import datetime
import shutil
import re

logging.basicConfig(filename='status_log.txt', level=logging.INFO)

fsd_list = [
    # Hong Kong Fire Commend
    # Central Fire Station
    ["cfs", "cjp_itmu_5@hkfsd.gov.hk"],
    # Kotewall Fire Station
    ["kfs", ""],
    # Kong Wan Fire Station
    ["kwfs", ""],
    # Sheung Wan Fire Station
    ["swfs", ""],
    # Sheung Wan Fire Station
    ["vpfs", ""],
    # Wan Chai Fire Station
    ["wcfs", ""],
    # Braemar Hill Fire Station
    ["bhfs", ""],
    # Chai Wan Fire Station
    ["cwfs", ""],
    # North Point Fire Station
    ["npfs", ""],
    # Sai Wan Ho Fire Station
    ["swhfs", ""],
    # Shau Kei Wan Fire Station
    ["skwfs", ""],
    # Tung Lo Wan Fire Station
    ["tlwfs", ""],
    # Aberdeen Fire Station
    ["afs", ""],
    # Ap Lei Chau Fire Station
    ["alcfs", ""],
    # Chung Hom Kok Fire Station
    ["chkfs", ""],
    # Kennedy Town Fire Station
    ["ktfs", ""],
    # Pok Fu Lam Fire Station
    ["pflfs", ""],
    # Sandy Bay Fire Station
    ["sbfs", ""],
    # Aberdeen Fireboat Station
    ["afbs", ""],
    # Central Fireboat Station
    ["", ""],
    # Cheung Chau Fire Station
    ["", ""],
    # Cheung Chau Fireboat Station
    ["", ""],
    # Cheung Sha Fire Station
    ["", ""],
    # Discovery Bay Fire Station
    ["", ""],
    # Lamma Fire Station
    ["", ""],
    # Mui Wo Fire Station
    ["", ""],
    # North Point Fireboat Station
    ["", ""],
    # Peng Chau Fire Station
    ["", ""],
    # Tai O (Old) Fire Station
    ["", ""],
    # Tai O Fire Station
    ["", ""],
    # Tsing Yi Fireboat Station
    ["", ""],
    # Tuen Mun Fireboat Station
    ["", ""],
    # <----------------------------->
    # Kowloon Command
    # Kai Tak Fire Station
    ["", ""],
    # Ma Tau Chung Fire Station
    ["", ""],
    # Ngau Chi Wan Fire Station
    ["", ""],
    # Sai Kung Fire Station
    ["", ""],
    # Shun Lee Fire Station
    ["", ""],
    # Wong Tai Sin Fire Station
    ["", ""],
    # Kowloon Bay Fire Station
    ["", ""],
    # Kwun Tong Fire Station
    ["", ""],
    # Lam Tin Fire Station
    ["", ""],
    # Po Lam Fire Station
    ["", ""],
    # Tai Chik Sha Fire Station
    ["", ""],
    # Yau Tong Fire Station
    ["", ""],
    # Hung Hom Fire Station
    ["", ""],
    # Tsim Sha Tsui Fire Station
    ["", ""],
    # Tsim Tung Fire Station
    ["", ""],
    # Yau Ma Tei Fire Station
    ["", ""],
    # Cheung Sha Wan Fire Station
    ["", ""],
    # Kowloon Tong Fire Station
    ["", ""],
    # Lai Chi Kok Fire Station
    ["", ""],
    # Mongkok Fire Station
    ["", ""],
    # Shek Kip Mei Fire Station
    ["", ""],

    # <------------------------------>
    # New Territories South Command
    # Airport South Fire Station
    ["", ""],
    # Airport Centre Fire Station
    ["", ""],
    # Airport North Fire Station
    ["", ""],
    # East Sea Rescue Berth
    ["", ""],
    # West Sea Rescue Berth
    ["", ""],
    # Kwai Chung Fire Station
    ["", ""],
    # Lai King Fire Station
    ["", ""],
    # Lei Muk Shue Fire Station
    ["", ""],
    # Sham Tseng Fire Station
    ["", ""],
    # Tsuen Wan Fire Station
    ["", ""],
    # Chek Lap Kok South Fire Station
    ["", ""],
    # Ma Wan Fire Station
    ["", ""],
    # Penny's Bay Fire Station
    ["", ""],
    # Tsing Yi Fire Station
    ["", ""],
    # Tsing Yi South Fire Station
    ["", ""],
    # Tung Chung Fire Station
    ["", ""],
    # Hong Kong-Zhuhai-Macao Bridge Fire Station
    ["", ""],


    # <------------------------------>
    # New Territories North Command
    # Ma On Shan Fire Station
    ["", ""],
    # Sha Tin Fire Station
    ["", ""],
    # Siu Lek Yuen Fire Station
    ["", ""],
    # Tai Po East Fire Station
    ["", ""],
    # Tai Po Fire Station
    ["", ""],
    # Tai Sum Fire Station
    ["", ""],
    # Fanling Fire Station
    ["", ""],
    # Mai Po Fire Station
    ["", ""],
    # Pat Heung Fire Station
    ["", ""],
    # Sha Tau Kok Fire Station
    ["", ""],
    # Sheung Shui Fire Station
    ["", ""],
    # Heung Yuen Wai Fire Station
    ["", ""],
    # Yuen Long Fire Station
    ["", ""],
    # Castle Peak Bay Fire Station
    ["", ""],
    # Fu Tei Fire Station
    ["", ""],
    # Lau Fau Shan Fire Station
    ["", ""],
    # Pillar Point Fire Station
    ["", ""],
    # Shenzhen Bay Fire Station
    ["", ""],
    # Tai Lam Chung Fire Station
    ["", ""],
    # Tin Shui Wai Fire Station
    ["", ""],
    # Tuen Mun Fire Station
    ["", ""],

    # <------------------------------>
    # <------------------------------>
    # <------------------------------>
    # Ambulance Depot
    # <------------------------------>


]


def move_zip_files():
    os.makedirs('./unzip', exist_ok=True)

    filenames = os.listdir('.')

    for filename in filenames:
        if filename.endswith('.zip'):
            if os.path.exists(f"./unzip/{filename}"):
                logging.info(
                    f'{datetime.datetime.now()}: The {filename} already exists in ./unzip, so skipping.')
            else:
                shutil.move(filename, './unzip')
                logging.info(
                    f'{datetime.datetime.now()}: The {filename} been moved into ./unzip')
                unzip_all_files()


def unzip_all_files():
    os.makedirs('./unzipped', exist_ok=True)
    zip_files = glob.glob('./unzip/*.zip')
    for zip_file in zip_files:
        folder_name = os.path.splitext(os.path.basename(zip_file))[0]
        os.makedirs(f'./unzipped/{folder_name}', exist_ok=True)
        with zipfile.ZipFile(zip_file, 'r') as zfile:
            zfile.extractall(f'./unzipped/{folder_name}')
            logging.info(f'{datetime.datetime.now()}: {zip_file} unzipped')
        os.remove(zip_file)
    if not zip_files:
        logging.error(
            f'{datetime.datetime.now()}: No ZIP files found in "./unzip" directory')
        return

# Function to find FSD in an XML file


def find_fsd(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    fsd_element = root.find('.//unit')
    if fsd_element is not None:
        return fsd_element.text
    return None

# Function to send email with a PDF file attached


def send_email(fsd_email, pdf_file, folder):
    email_address = 'efax_admin@hkfsd.hksarg'
    email_app_password = ''
    msg = EmailMessage()
    msg['Subject'] = 'Application for Visit to FSD Premises 申請參觀消防處單位 (參考編號：' + str(folder)+')'
    msg['From'] = email_address
    msg['To'] = fsd_email

    msg.set_content('Please find the PDF file attached.')

    with open(pdf_file, 'rb') as pdf:
        msg.add_attachment(pdf.read(), maintype='application',
                           subtype='pdf', filename=pdf_file)

    with smtplib.SMTP('10.18.11.64') as smtp:
        # smtp.login(email_address, email_app_password)
        smtp.send_message(msg)
        logging.info(
            f'{datetime.datetime.now()}: {pdf_file} sent to {fsd_email}')


# def move_folders_to_sent():
#     os.makedirs(sent_folder, exist_ok=True)
#     sent_folder = "./sent"
#     unzipped_folders = os.listdir('./unzipped')
#     for folder in unzipped_folders:
#         shutil.move(f'./unzipped/{folder}', sent_folder)
#     logging.info(
#         f'{datetime.datetime.now()}: All folders moved to "sent" directory')


def main():
    move_zip_files()
    unzipped_folders = os.listdir('./unzipped')
    for folder in unzipped_folders:
        xml_files = glob.glob(
            f'./unzipped/{folder}/{folder}/convertedData/*.xml')
        if not xml_files:
            logging.error(
                f'{datetime.datetime.now()}: No XML files found in {folder}. Skipping.')
            continue
        xml_file = xml_files[0]
        fsd_in_xml = find_fsd(xml_file)

        if fsd_in_xml:
            for fsd, fsd_email in fsd_list:
                # If FSD in the XML file matches one in the list, send email with PDF attached
                if fsd == fsd_in_xml:
                    pdf_files = glob.glob(
                        f'./unzipped/{folder}/{folder}/convertedData/*.pdf')
                    if not pdf_files:
                        status = 'PDF not found in folder'
                        logging.error(
                            f'{datetime.datetime.now()}: {folder}: {status}')
                        break
                    pdf_file = pdf_files[0]
                    send_email(fsd_email, pdf_file,folder)
                    status = 'Success, Email sent'
                    break
                else:
                    status = 'FSD unit not found in list'
            else:
                status = 'FSD not found in XML'
        logging.info(f'{datetime.datetime.now()}: {folder}: {status}')



if __name__ == "__main__":
    main()
