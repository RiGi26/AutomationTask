import pandas as pd
import win32com.client as win32
from pathlib import Path
from datetime import datetime
import time

# Setting Date
current_date = datetime.now().strftime("%d %B %Y")

# Define the folder paths
base_path = Path("D:\\Daily MOXA\\Data Reminder Moxa")
base_path_DaaS = Path("D:\\Daily MOXA\\Data Reminder DaaS")

# Load the email list
file = "D:\\Daily MOXA\\Automate Send to MD\\Email list reminder .xlsx"
email_list = pd.read_excel(file)

# Load the master reminder MD MOXA or DaaS
path = "D:\\Daily MOXA\\Data Reminder Moxa\\Remainder Data Leads Master.xlsx"
main_dealer = pd.read_excel(path)
path_DaaS = "D:\\Daily MOXA\\Data Reminder DaaS\\Remainder Data Leads Master DaaS.xlsx"
main_dealer_DaaS = pd.read_excel(path_DaaS)

# Get unique "Main Dealer" values
main_dealer_ids = main_dealer["Main Dealer"].unique()
main_dealer_for_DaaS = main_dealer_DaaS["Main Dealer"].unique()
filtered_email_list = email_list[email_list['Main Dealer'].isin(main_dealer_ids)]
filtered_email_list_DaaS = email_list[email_list['Main Dealer'].isin(main_dealer_for_DaaS)]

# Get overlapping main dealers
overlapping_main_dealers = set(filtered_email_list['Main Dealer']).intersection(filtered_email_list_DaaS['Main Dealer'])

# Set up Outlook
outlook = win32.Dispatch("outlook.application")

def attach_files(mail, attachment_filenames):
    for filename, path_base in attachment_filenames:
        if path_base:
            attachment_path = path_base / filename
            print(f"Checking for file: {attachment_path}")
            if attachment_path.exists():
                print(f"Attaching file: {attachment_path}")
                try:
                    mail.Attachments.Add(str(attachment_path))
                except Exception as e:
                    print(f"Failed to attach {attachment_path}: {e}")
                time.sleep(1)
            else:
                print(f"Attachment not found: {attachment_path}")
        else:
            print(f"Invalid path base for attachment: {filename}")

def send_email(row, main_dealer_name, project_type, base_path=None):
    if project_type == 'DaaS & MOXA':
        subject = f"Reminder Data Leads {main_dealer_name} (DaaS & MOXA)"
        attachment_filenames = [
            (f"Remainder Data Leads {main_dealer_name} DaaS.xlsx", base_path_DaaS),
            (f"Remainder Data Leads {main_dealer_name}.xlsx", base_path)
        ]
    elif project_type == 'DaaS':
        subject = f"Reminder Data Leads {main_dealer_name} (DaaS)"
        attachment_filenames = [(f"Remainder Data Leads {main_dealer_name} DaaS.xlsx", base_path_DaaS)]
    else:
        subject = f"Reminder Data Leads {main_dealer_name} (MOXA)"
        attachment_filenames = [(f"Remainder Data Leads {main_dealer_name}.xlsx", base_path)]
    
    mail = outlook.CreateItem(0)
    mail.To = row["to"]
    mail.CC = row["cc"]
    mail.Subject = subject
    mail.HTMLBody = f"""
    <html>
    <body style="font-family: Calibri, sans-serif; font-size: 11pt; color: black;">
    <p>Dear Bapak & Ibu PIC Main Dealer,</p>

    <p>Sehubungan dengan adanya Project {project_type},
    Kami ingin mengingatkan kembali untuk pemberian hasil follow up data leads  diberikan 1 hari setelah data diberikan oleh FIFGROUP.
    Untuk format feedback menyesuaikan dengan format yang diambil dari monitorku.
    </p>

    <p>Berikut kami lampirkan kembali data leads yang berasal dari {project_type} dan belum ada status feedback > 3 hari.</p>

    <p>Bila ada hal yang tidak sesuai atau ada pertanyaan bisa langsung menghubungi saya.</p>

    <p>Terima kasih atas bantuan dan kerjasamanya,</p>

    <p>Best Regards,<br>
    Riyadh Akhdan Syafi<br>
    <strong>CRM Data Mining</strong><br>
    <a href="mailto:riyadh.asyafi@fifgroup.astra.co.id">riyadh.asyafi@fifgroup.astra.co.id</a>
    </p>
    </body>
    </html>
    """
    
    # Attach files
    attach_files(mail, attachment_filenames)
    
    # Display the email
    # mail.Display()
    
    print(f"Email has been generated for {main_dealer_name} ({project_type})")
    # Uncomment the line below to send the email
    mail.Send()

for index, row in email_list.iterrows():
    main_dealer_name = row["Main Dealer"]
    if main_dealer_name in overlapping_main_dealers:
        print(f"Main dealer {main_dealer_name} is in both MOXA and DaaS lists")
        send_email(row, main_dealer_name, 'DaaS & MOXA', base_path)
    elif main_dealer_name not in overlapping_main_dealers:
        if main_dealer_name in main_dealer_ids:
            send_email(row, main_dealer_name, 'MOXA', base_path)
        elif main_dealer_name in main_dealer_for_DaaS:
            send_email(row, main_dealer_name, 'DaaS', base_path_DaaS)
