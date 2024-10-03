import pandas as pd
import os
import win32com.client as win32
from pathlib import Path
from datetime import datetime
import time
import re

# Setting Date
current_date = datetime.now().strftime("%d %B %Y")

# Define the folder paths
folder_date = "03" # Change it
folder_month = "Oktober" # Change it
format_date = "20241003" # Change it
base_path = Path(f"D:\\Daily MOXA\\blackup kirim dealer\\2024\\{folder_month}\\{folder_date}")
path_file = Path(f"D:\\Daily MOXA\\blackup kirim dealer\\2024\\{folder_month}\\{folder_date}\\DATA GABUNGAN LEADS FIFGROUP {format_date}.xlsx")

# Load the email list
file = "D:\\Daily MOXA\\Automate Send to MD\\Email list.xlsx"
file_DaaS = "D:\Daily MOXA\DAAS\Rekap DAAS Februari 2023.xlsx"
email_list = pd.read_excel(file)

# Load data daily
dealer = pd.read_excel(path_file)

# Moxa
main_dealer_filtered = dealer['Main Dealer'].unique()
filter_email = email_list[email_list['Main Dealer'].isin(main_dealer_filtered)]
filter_email_MD = filter_email['Main Dealer'].unique()

# DaaS
dealer_DaaS = pd.read_excel(file_DaaS)
filter_DaaS = dealer_DaaS[dealer_DaaS['Dispatch Date'] == pd.to_datetime(current_date)] # Change the date to the current date
DaaS_main_dealer = filter_DaaS['Main Dealer'].unique()
filter_email_DaaS = email_list[email_list['Main Dealer'].isin(DaaS_main_dealer)]
filter_email_DaaS_MD = filter_email_DaaS['Main Dealer'].unique()

# Moxa DaaS
over_lapping_maindealer = set(dealer['Main Dealer']).intersection(filter_DaaS["Main Dealer"])
over_lapping_maindealer_list = list(over_lapping_maindealer)
filter_email_DaaS_MOXA = email_list[email_list['Main Dealer'].isin(over_lapping_maindealer)]

# Set up Outlook
outlook = win32.Dispatch("outlook.application")

def attach_files(mail, attachment_filenames):
    # Attach files if they exist
    for filename, path_base in attachment_filenames:
        if path_base:  # Ensure path_base is not None
            attachment_path = path_base / filename
            print(f"Checking for file: {attachment_path}")
            if attachment_path.exists():
                print(f"Attaching file: {attachment_path}")
                try:
                    mail.Attachments.Add(str(attachment_path))
                except Exception as e:
                    print(f"Failed to attach {attachment_path}: {e}")
                time.sleep(1)  # Add a short delay before attaching the next file
            else:
                continue
        else:
            continue

def send_email(row, main_dealer_name, project_type, base_path):
    if project_type == 'DaaS & MOXA':
        subject = f"Data leads FIFGROUP {current_date} {row['Main Dealer']} (DaaS & MOXA)"
        attachment_filenames = [(f"Data leads FIFGROUP {format_date} {row['Main Dealer']}.xlsx", base_path),
                                (f"Data Leads FIFGROUP {format_date} {row['Main Dealer']} DaaS.xlsx", base_path)]
    elif project_type == 'DaaS':
        subject = f"Data leads FIFGROUP {current_date} {row['Main Dealer']} (DaaS)"
        attachment_filenames = [(f"Data Leads FIFGROUP {format_date} {row['Main Dealer']} DaaS.xlsx", base_path)]
    elif project_type == 'MOXA':
        subject = f"Data leads FIFGROUP {current_date} {row['Main Dealer']} (MOXA)"
        attachment_filenames = [(f"Data leads FIFGROUP {format_date} {row['Main Dealer']}.xlsx", base_path)]
    else:
        print(f"Invalid project type: {project_type}")
        return
    
    # create email
    mail = outlook.CreateItem(0)
    mail.To = row["to"]
    mail.CC = row["cc"]
    mail.Subject = subject
    mail.HTMLBody = f"""
<html>
<body style="font-family: Calibri, sans-serif; font-size: 11pt; color: black;">
<p>Dear Bapak & Ibu Yth,</p>

<p>Berikut terlampir data leads untuk pembiayaan motor baru dari aplikasi {project_type}.</p>

<p>Kami telah menambahkan waktu customer ingin dihubungi kembali, melalui channel apa customer ingin dihubungi kembali dan customer yang ingin melakukan pengajuan Syariah 
pada kolom remarks.</p>

<p>Terima kasih atas bantuan dan kerjasamanya,</p>

<p>Best Regards,<br>
Riyadh Akhdan Syafi<br>
<strong>CRM Data Mining</strong><br>
<a href="mailto:riyadh.asyafi@fifgroup.astra.co.id">riyadh.asyafi@fifgroup.astra.co.id</a>
</p>
</body>
</html>
"""
    # running attach
    attach_files(mail, attachment_filenames)

    # Display
    # mail.Display()

    # Sending
    mail.Send()

def extract_dealer_name(filename):
    match = re.search(r'FIFGROUP \d+ (.+)\.xlsx', filename)
    if match :
        dealer_name = match.group(1)
        dealer_name = dealer_name.replace('DaaS', '').strip()
        return dealer_name
    return None

processed_dealers = set()

# Iterate through files in the base_path directory
with os.scandir(base_path) as entries:
    for entry in entries:
        if entry.is_file() and entry.name.endswith('.xlsx'):  # Only check files with .xlsx extension
            dealer_name = extract_dealer_name(entry.name)

            # Skip if the dealer_name is missing or if it has already been processed
            if dealer_name is None or dealer_name in processed_dealers:
                continue

            # Check if the dealer is in the overlapping dealers set
            if dealer_name in over_lapping_maindealer_list:
                # Send DaaS & MOXA email
                for index, row in filter_email_DaaS_MOXA.iterrows():
                    print(f"Sending DaaS & MOXA email to {dealer_name}")
                    if dealer_name == row['Main Dealer']:
                        send_email(row, dealer_name, 'DaaS & MOXA', base_path)
                        # Add dealer to the processed set to avoid reprocessing
                        processed_dealers.add(dealer_name)
                    else:
                        print(f"Dealer not found in overlapping main dealer list: {dealer_name}")

            # If the dealer is not in the overlapping maindealer list
            elif dealer_name not in over_lapping_maindealer_list:
                # Check if the dealer is in the DaaS filtered email list
                if dealer_name in filter_email_DaaS_MD:
                    for index, row in filter_email_DaaS.iterrows():
                        print(f"Sending DaaS email to {dealer_name}")
                        if dealer_name == row['Main Dealer']:
                            send_email(row, dealer_name, 'DaaS', base_path)
                            processed_dealers.add(dealer_name)
                        else:
                            print(f"Dealer not found in DaaS main dealer list: {dealer_name}")
                
                # If the dealer is in the regular MOXA email list
                elif dealer_name in filter_email_MD:
                    for index, row in filter_email.iterrows():
                        print(f"Sending MOXA email to {dealer_name}")
                        if dealer_name == row['Main Dealer']:
                            send_email(row, dealer_name, 'MOXA', base_path)
                            processed_dealers.add(dealer_name)
                        else:
                            print(f"Dealer not found in MOXA main dealer list: {dealer_name}")
                    
                    
                    
