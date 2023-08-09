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
    ["Central Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Kotewall Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Kong Wan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sheung Wan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Victoria Peak Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Wan Chai Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Braemar Hill Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Chai Wan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["North Point Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sai Wan Ho Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Shau Kei Wan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tung Lo Wan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Aberdeen Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Ap Lei Chau Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Chung Hom Kok Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Kennedy Town Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Pok Fu Lam Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sandy Bay Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Aberdeen Fireboat Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Central Fireboat Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Cheung Chau Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Cheung Chau Fireboat Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Cheung Sha Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Discovery Bay Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Lamma Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Mui Wo Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["North Point Fireboat Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Peng Chau Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tai O (Old) Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tai O Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tsing Yi Fireboat Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tuen Mun Fireboat Station", "cjp_itmu_5@hkfsd.gov.hk"],
    # <----------------------------->
    # Kowloon Command
    ["Kai Tak Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Ma Tau Chung Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Ngau Chi Wan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sai Kung Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Shun Lee Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Wong Tai Sin Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Kowloon Bay Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Kwun Tong Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Lam Tin Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Po Lam Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tai Chik Sha Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Yau Tong Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Hung Hom Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tsim Sha Tsui Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tsim Tung Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Yau Ma Tei Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Cheung Sha Wan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Kowloon Tong Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Lai Chi Kok Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Mongkok Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Shek Kip Mei Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],

    # <------------------------------>
    # New Territories South Command

    ["Airport South Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Airport Centre Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Airport North Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["East Sea Rescue Berth", "cjp_itmu_5@hkfsd.gov.hk"],
    ["West Sea Rescue Berth", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Kwai Chung Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Lai King Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Lei Muk Shue Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sham Tseng Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tsuen Wan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Chek Lap Kok South Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Ma Wan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Penny's Bay Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tsing Yi Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tsing Yi South Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tung Chung Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Hong Kong-Zhuhai-Macao Bridge Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],


    # <------------------------------>
    # New Territories North Command

    ["Ma On Shan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sha Tin Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Siu Lek Yuen Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tai Po East Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tai Po Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tai Sum Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Fanling Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Mai Po Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Pat Heung Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sha Tau Kok Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sheung Shui Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Heung Yuen Wai Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Yuen Long Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Castle Peak Bay Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Fu Tei Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Lau Fau Shan Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Pillar Point Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Shenzhen Bay Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tai Lam Chung Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tin Shui Wai Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tuen Mun Fire Station", "cjp_itmu_5@hkfsd.gov.hk"],

    # <------------------------------>
    # <------------------------------>
    # <------------------------------>
    # Ambulance Depot
    # <------------------------------>

    # Hong Kong and Kowloon Region
    ["Braemar Hill Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Chai Wan Ambulance Station", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Morrison Hill Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sai Wan Ho Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Aberdeen Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Mount Davis Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Pok Fu Lam Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Lam Tin Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Ngau Tau Kok Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Po Lam Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tai Chik Sha Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Wong Tai Sin Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Ho Man Tin Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Kowloon Tong Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Ma Tau Chung Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Pak Tin Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Cheung Sha Wan Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Lai Chi Kok Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Mongkok Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tsim Tung Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Yau Ma Tei Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],

    # New Territories Region
    ["Ta Kwu Ling Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Fanling Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Lau Fau Shan Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sheung Shui Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tai Po Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tin Shui Wai Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Castle Peak Bay Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sham Tseng Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tuen Mun Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Yuen Long Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Lei Muk Shue Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Ma On Shan Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Sha Tin Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tin Sum Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Kwai Chung Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Penny's Bay Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tsing Yi Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tsuen Wan Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Tung Chung Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"],
    ["Hong Kong-Zhuhai-Macao Bridge Ambulance Depot", "cjp_itmu_5@hkfsd.gov.hk"]
]


def move_zip_files():
    # os.makedirs('./unzip', exist_ok=True)
    filenames = os.listdir('.')
    zip_found = False

    for filename in filenames:
        if filename.endswith('.zip'):
            if os.path.exists(f"./sent/{filename}"):
                zip_found = True
                logging.info(
                    f'{datetime.datetime.now()}: The {filename} already exists in ./sent, so skipping.')
            else:
                # shutil.move(filename, './unzip')
                # logging.info(
                #     f'{datetime.datetime.now()}: The {filename} been moved into ./unzip')
                unzip_all_files()
    if not zip_found:
        logging.error(f'{datetime.datetime.now()}: No ZIP files found in "./" directory. Program Halted.')
        return

def unzip_all_files():
    # zip_files = glob.glob('./unzip/*.zip')
    zip_files = glob.glob('./*.zip')
    for zip_file in zip_files:
        folder_name = os.path.splitext(os.path.basename(zip_file))[0]
        os.makedirs(f'./unzipped/{folder_name}', exist_ok=True)
        with zipfile.ZipFile(zip_file, 'r') as zfile:
            zfile.extractall(f'./unzipped/{folder_name}')
            logging.info(f'{datetime.datetime.now()}: {zip_file} unzipped')
        os.remove(zip_file)
    if not zip_files:
        logging.error(
            f'{datetime.datetime.now()}: No ZIP files found in "./" directory')
        return

def find_fsd(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()
    fsd_element = root.find('.//unit')
    if fsd_element is not None:
        return fsd_element.text
    return None


def send_email(fsd_email, pdf_file, folder):
    email_address = 'efax_admin@hkfsd.hksarg'
    email_app_password = ''
    msg = EmailMessage()
    msg['Subject'] = 'Application for Visit to FSD Premises 申請參觀消防處單位 (參考編號：' + re.split('_',str(folder),5)[5]+')'
    msg['From'] = email_address
    msg['To'] = fsd_email

    msg.set_content('Please find the PDF file attached. Submission date: '+re.split('_',str(folder),5)[0]+' '+re.split('_',str(folder),5)[1]+':'+re.split('_',str(folder),5)[2])

    with open(pdf_file, 'rb') as pdf:
        msg.add_attachment(pdf.read(), maintype='application', subtype='pdf', filename=pdf_file)

    with smtplib.SMTP('10.18.11.64') as smtp:
        try:
            smtp.send_message(msg)
            logging.info(f'{datetime.datetime.now()}: {pdf_file} sent to {fsd_email}')
        except smtplib.SMTPRecipientsRefused as e:
            logging.error(f'{datetime.datetime.now()}: {pdf_file} failed to send to {fsd_email}')
            for recipient, (code, errmsg) in e.recipients.items():
                logging.error(f'{recipient} was refused. Error code: {code}, Error message: {errmsg}')
            

# def move_folder_to_sent(folder):
#     os.makedirs('./sent', exist_ok=True)
#     sent_folder = "./sent"
#     # shutil.move(f'./sent/{folder}', sent_folder)
#     shutil.make_archive(f'./sent/{folder}','zip', sent_folder)
#     shutil.rmtree(f'./unzipped/{folder}')
#     os.makedirs('./sent', exist_ok=True)
#     shutil.move(f'./unzipped/{folder}', './sent')


def main():
    move_zip_files()
    os.makedirs('./unzipped', exist_ok=True)
    unzipped_folders = os.listdir('./unzipped')
    for folder in unzipped_folders:
        xml_files = glob.glob(
            f'./unzipped/{folder}/convertedData/*.xml')
        if not xml_files:
            logging.error(
                f'{datetime.datetime.now()}: No XML files found in {folder}. Skipping.')
            continue
        xml_file = xml_files[0]
        fsd_in_xml = find_fsd(xml_file)

        if fsd_in_xml:
            for fsd, fsd_email in fsd_list:
                if fsd == fsd_in_xml:
                    pdf_files = glob.glob(
                        f'./unzipped/{folder}/convertedData/*.pdf')
                    if not pdf_files:
                        status = 'PDF not found in folder'
                        logging.error(
                            f'{datetime.datetime.now()}: {folder}: {status}')
                        break
                    pdf_file = pdf_files[0]
                    send_email(fsd_email, pdf_file,folder)
                    status = 'Success, Email sent'
                    os.makedirs('./sent', exist_ok=True)
                    shutil.move(f'./unzipped/{folder}', './sent')
                    break
                else:
                    status = 'FSD unit not found in list'
            else:
                status = 'XML not found'
        else:
            status = 'No element found in XML'
        logging.info(f'{datetime.datetime.now()}: {folder}: {status}')



if __name__ == "__main__":
    main()
