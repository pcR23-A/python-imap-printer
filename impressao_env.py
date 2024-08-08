import logging #Lines 13, 40, 43, 51, 66, 70, 72, 91, 97, 99, 101, 107, 109, 112, 121 - Assists on debugging
import imaplib #Line 76 - Connects to IMAP server
import email #Line 90 - Parses email
import re #Line 21 - Extracts links
import requests #Line 34 - GET requests
import os #Lines 15-18, 37 - Calls environment variables / Manages file paths
import tempfile #Line 36 - Gets temporary files
import time #Line 125 - Makes the program wait until next check
from win32 import win32print, win32api #Lines 49, 50, 58, 62, 65, 68 - Manages printers and printing files
import platform #Line 47 - Checks which OS you're using

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

username = os.getenv('EMAIL_USERNAME')
password = os.getenv('EMAIL_PASSWORD')
imap_server = os.getenv('IMAP_SERVER')
from_address = os.getenv('FROM_ADDRESS')

def extract_links(text):
    matches = re.findall(r'(http[s]?://[^",\s]+)', text)
    if matches:
        return matches
    return []

def filter_links(links, keyword):
    # Filter links that contain the keyword in the filename
    filtered_links = [link for link in links if keyword in link]
    return filtered_links

def download_file(url):
    try:
        local_filename = url.split('/')[-1]
        response = requests.get(url)
        response.raise_for_status()
        temp_dir = tempfile.gettempdir()
        file_path = os.path.join(temp_dir, local_filename)
        with open(file_path, 'wb') as f:
            f.write(response.content)
        logging.debug(f'File downloaded: {file_path}')
        return file_path
    except Exception as e:
        logging.error(f'Error downloading file: {e}')
        return None

def print_file(file_path):
    if platform.system() == "Windows":
        try:
            printer_name = win32print.GetDefaultPrinter()
            hPrinter = win32print.OpenPrinter(printer_name)
            logging.debug(f'Printer opened: {printer_name}')
            
            # Define as constantes manualmente
            DMORIENT_LANDSCAPE = 2
            DM_IN_BUFFER = 8
            DM_OUT_BUFFER = 2

            properties = win32print.GetPrinter(hPrinter, 2)
            devmode = properties['pDevMode']
            devmode.Orientation = DMORIENT_LANDSCAPE
            
            win32print.DocumentProperties(None, hPrinter, printer_name, devmode, devmode, DM_IN_BUFFER | DM_OUT_BUFFER)
            
            # Start the print job
            win32api.ShellExecute(0, "print", file_path, None, ".", 0)
            logging.debug('File sent to printer')

            win32print.ClosePrinter(hPrinter)
        except Exception as e:
            logging.error(f'Error printing file: {e}')
    else:
        logging.error('Printing is only supported on Windows')

def mail_check():
    try:
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(username, password)
        status, messages = imap.select('INBOX')

        status, search_data = imap.search(None, '(UNSEEN FROM "{}")'.format(from_address))
        if status == 'OK':
            email_ids = search_data[0].split()
            print('Found {} unseen messages from {}.'.format(len(email_ids), from_address))

            for email_id in email_ids:
                status, msg_data = imap.fetch(email_id, '(RFC822)')
                if status == 'OK':
                    for response_part in msg_data:
                        if isinstance(response_part, tuple):
                            msg = email.message_from_bytes(response_part[1])
                            logging.debug(f'Processing email ID: {email_id}')
                            if msg.is_multipart():
                                for part in msg.walk():
                                    if part.get_content_type() == "text/plain":
                                        body = part.get_payload(decode=True).decode()
                                        links = extract_links(body)
                                        logging.debug(f'Email body: {body}')
                                        filtered_links = filter_links(links, 'ROTULO')
                                        logging.debug(f'Extracted links: {links}')
                                        for link in filtered_links:
                                            logging.debug(f'Found link: {link}')
                                            print('Found link:', link)
                                            file_path = download_file(link)
                                            print_file(file_path)
                            else:
                                body = msg.get_payload(decode=True).decode()
                                logging.debug(f'Email body: {body}')
                                links = extract_links(body)
                                logging.debug(f'Extracted links: {links}')
                                filtered_links = filter_links(links, 'ROTULO')
                                for link in filtered_links:
                                    logging.debug(f'Found link: {link}')
                                    print('Found link:', link)
                                    file_path = download_file(link)
                                    print_file(file_path)
                
                imap.store(email_id, '+FLAGS', '\\Seen')
        
        imap.logout()
    except Exception as e:
        logging.error(f'Error checking mail: {e}')

while True:
    mail_check()
    time.sleep(300)
