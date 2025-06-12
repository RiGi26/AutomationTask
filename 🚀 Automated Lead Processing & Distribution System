import os
import re
import time
from datetime import datetime
from pathlib import Path
import logging

import numpy as np
import openpyxl as op
import pandas as pd
import win32com.client as win32
from openpyxl.styles import Alignment, Font, Border, Side

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuration Constants
FOLDER_DATE = "30"  # change
FOLDER_YEAR = "2025"  # change
FOLDER_MONTH = "Mei"  # change

# Date configurations
current_date = datetime.now().strftime("%Y%m%d")
dispatch_date = datetime.now().strftime("%d/%m/%Y")
filter_date = datetime.now().strftime("%d %B %Y")

# Path configurations
BASE_PATHS = {
    'base': Path(f"D:\\Daily MOXA\\blackup kirim dealer\\{FOLDER_YEAR}\\{FOLDER_MONTH}\\{FOLDER_DATE}"),
    'data_file': Path(
        f"D:\\Daily MOXA\\blackup kirim dealer\\{FOLDER_YEAR}\\{FOLDER_MONTH}\\{FOLDER_DATE}\\DATA GABUNGAN LEADS FIFGROUP {current_date}.xlsx"),
    'master': Path("D:\\Daily MOXA\\Master Leads Interest 2025.xlsx"),
    'email_list': "D:\\Daily MOXA\\Automate Send to MD\\Email list.xlsx",
    'daas_file': "D:\\Daily MOXA\\DAAS\\Rekap DAAS Februari 2023.xlsx",
    'folder': Path("D:\\Daily MOXA"),
    'recap': Path("D:\\Daily MOXA\\Leads FIFGROUP Compile all MD v2.xlsx")
}

# Column mappings
COLUMN_MAPPING = {
    "Id Leads Data User": "Id Leads Data User",
    "Nama": "Nama",
    "Gender": "Gender",
    "Alamat": "Alamat",
    "Kelurahan": "Kelurahan",
    "Kecamatan": "Kecamatan",
    "Propinsi": "Propinsi",
    "Kota/Kabupaten": "Kota/Kabupaten",
    "No HP": "No HP",
    "MD (3 DIGIT)": "Main Dealer",
    "Pendidikan": "Pendidikan",
    "Tanggal Lahir": "Tanggal Lahir",
    "E-MAIL": "E-MAIL",
    "Dealer Sebelumnya (Jika Ada)": "Dealer Sebelumnya (Jika Ada)",
    "remarks": "Remarks/Keterangan"
}

FINAL_COLUMNS = [
    "Id Leads Data User", "Nama", "Gender", "Alamat", "Kelurahan", "Kecamatan",
    "Kota/Kabupaten", "Propinsi", "No HP", "No Hp-2", "Sales Date", "Varian Motor",
    "Main Dealer", "Assign Dealer Code (5 DIGIT)", "Propensity", "Pekerjaan",
    "Pendidikan", "Pengeluaran", "Agama", "Tanggal Lahir", "Frame No Terakhir",
    "Jenis Penjualan", "Sales ID", "Nama Leasing Sebelumnya", "Nama salesman",
    "Source Leads", "Platform Data", "Dealer Sebelumnya (Jika Ada)",
    "Remarks/Keterangan", "Rekomendasi DP/Angsuran (Tenure)",
    "Varian motor yang diinginkan", "Warna varian motor", "E-MAIL",
    "FACEBOOK", "INSTAGRAM", "TWITTER"
]

MPM_COLUMNS = [
    "Id Leads Data User", "Nama", "No HP", "Kota/Kabupaten",
    "Kecamatan", "Alamat", "Main Dealer"
]

HIDDEN_COLUMNS = [
    'C', 'D', 'E', 'F', 'G', 'H', 'J', 'K', 'L', 'N', 'O', 'P', 'Q', 'R', 'S',
    'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG',
    'AH', 'AI', 'AJ'
]


class DataProcessor:
    def __init__(self):
        self.processed_dealers = set()
        self.outlook = None
        self.master_df = None
        self.email_list_df = None
        self.dealer_daas_df = None

    def initialize_data(self):
        """Initialize all required data sources"""
        try:
            # Read master data
            self.master_df = pd.read_excel(BASE_PATHS['master'], sheet_name=FOLDER_MONTH,
                                           parse_dates=['Tanggal Lahir'])
            self.email_list_df = pd.read_excel(BASE_PATHS['email_list'])
            self.dealer_daas_df = pd.read_excel(BASE_PATHS['daas_file'])

            # Initialize Outlook
            self.outlook = win32.Dispatch("outlook.application")

            logger.info("Data initialization completed successfully")
            return True

        except Exception as e:
            logger.error(f"Failed to initialize data: {e}")
            return False

    def filter_data(self):
        """Filter data based on date criteria"""
        try:
            df_filtered = self.master_df[
                self.master_df["tgl"] == pd.to_datetime(filter_date)
                ]
            filter_daas = self.dealer_daas_df[
                self.dealer_daas_df['Dispatch Date'] == pd.to_datetime(filter_date)
                ]

            logger.info(f"Filtered data: {len(df_filtered)} MOXA records, {len(filter_daas)} DaaS records")
            return df_filtered, filter_daas

        except Exception as e:
            logger.error(f"Failed to filter data: {e}")
            return None, None

    def normalize_phone_number(self, phone_str):
        """Normalize phone number format"""
        phone_str = str(phone_str)
        return "0" + phone_str if not phone_str.startswith("0") else phone_str

    def process_main_data(self, df_filtered):
        """Process and transform main data"""
        try:
            # Apply column mapping
            df_pindah = df_filtered[list(COLUMN_MAPPING.keys())].rename(columns=COLUMN_MAPPING)

            # Normalize phone numbers
            df_pindah["No HP"] = df_pindah["No HP"].apply(self.normalize_phone_number)

            # Add missing columns
            for kolom in FINAL_COLUMNS:
                if kolom not in df_pindah.columns:
                    df_pindah[kolom] = np.nan

            df_pindah = df_pindah[FINAL_COLUMNS]
            logger.info("Main data processing completed")
            return df_pindah

        except Exception as e:
            logger.error(f"Failed to process main data: {e}")
            return None

    def save_excel_with_formatting(self, df, output_path, sheet_name='Sheet1'):
        """Save Excel file with consistent formatting"""
        try:
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                df_clean = df.fillna('')
                df_clean.to_excel(writer, sheet_name=sheet_name, index=False)

                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                border_format = workbook.add_format({'border': 1})
                date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})

                # Apply formatting
                for row_num in range(len(df_clean) + 1):
                    for col_num, col in enumerate(df_clean.columns):
                        if row_num == 0:
                            worksheet.write(row_num, col_num, col, border_format)
                        else:
                            value = df_clean.iloc[row_num - 1, col_num]
                            format_to_use = date_format if col == 'Tanggal Lahir' else border_format
                            worksheet.write(row_num, col_num, value, format_to_use)

                # Auto-adjust column widths
                for idx, col in enumerate(df_clean.columns):
                    max_len = max(df_clean[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(idx, idx, max_len)

            logger.info(f"File saved successfully: {output_path}")
            return True

        except Exception as e:
            logger.error(f"Failed to save file {output_path}: {e}")
            return False

    def split_data_by_dealer(self, df, output_dir):
        """Split data by dealer and save individual files"""
        unique_dealers = df["Main Dealer"].unique()

        for dealer in unique_dealers:
            try:
                df_dealer = df[df["Main Dealer"] == dealer].copy()
                df_dealer["No HP"] = df_dealer["No HP"].apply(self.normalize_phone_number)

                # Special handling for MPM dealers
                if dealer in ["PT MPM - MALANG", "PT MPM - SURABAYA"]:
                    df_dealer = df_dealer[MPM_COLUMNS]

                output_path = output_dir / f"Data Leads FIFGROUP {current_date} {dealer}.xlsx"
                self.save_excel_with_formatting(df_dealer, output_path)
                logger.info(f"File split completed for {dealer}")

            except Exception as e:
                logger.error(f"Failed to split data for {dealer}: {e}")

    def extract_dealer_name(self, filename):
        """Extract dealer name from filename"""
        match = re.search(r'FIFGROUP \d+ (.+)\.xlsx', filename)
        if match:
            dealer_name = match.group(1).replace('DaaS', '').strip()
            return dealer_name
        return None

    def attach_files(self, mail, attachment_filenames):
        """Attach files to email"""
        for filename, path_base in attachment_filenames:
            if path_base:
                attachment_path = path_base / filename
                if attachment_path.exists():
                    mail.Attachments.Add(str(attachment_path))
                    time.sleep(1)

    def create_email_body(self, project_type):
        """Create standardized email body"""
        return f"""
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

    def send_email(self, row, main_dealer_name, project_type, base_path):
        """Send email with appropriate attachments"""
        email_configs = {
            'DaaS & MOXA': {
                'subject': f"Data leads FIFGROUP {filter_date} {row['Main Dealer']} (DaaS & MOXA)",
                'attachments': [
                    (f"Data leads FIFGROUP {current_date} {row['Main Dealer']}.xlsx", base_path),
                    (f"Data Leads FIFGROUP {current_date} {row['Main Dealer']} DaaS.xlsx", base_path)
                ]
            },
            'DaaS': {
                'subject': f"Data leads FIFGROUP {filter_date} {row['Main Dealer']} (DaaS)",
                'attachments': [(f"Data Leads FIFGROUP {current_date} {row['Main Dealer']} DaaS.xlsx", base_path)]
            },
            'MOXA': {
                'subject': f"Data leads FIFGROUP {filter_date} {row['Main Dealer']} (MOXA)",
                'attachments': [(f"Data leads FIFGROUP {current_date} {row['Main Dealer']}.xlsx", base_path)]
            }
        }

        if project_type not in email_configs:
            logger.error(f"Invalid project type: {project_type}")
            return

        config = email_configs[project_type]

        try:
            mail = self.outlook.CreateItem(0)
            mail.To = row["to"]
            mail.CC = row["cc"]
            mail.Subject = config['subject']
            mail.HTMLBody = self.create_email_body(project_type)

            self.attach_files(mail, config['attachments'])
            mail.Display()
            mail.Send()

            logger.info(f"Email sent successfully to {main_dealer_name} for {project_type}")

        except Exception as e:
            logger.error(f"Failed to send email to {main_dealer_name}: {e}")

    def process_email_sending(self, base_path, dealer_df, filter_daas):
        """Process email sending logic"""
        try:
            # Get unique dealers
            main_dealer_filtered = dealer_df['Main Dealer'].unique()
            daas_main_dealer = filter_daas['Main Dealer'].unique()

            # Filter email lists
            filter_email = self.email_list_df[self.email_list_df['Main Dealer'].isin(main_dealer_filtered)]
            filter_email_daas = self.email_list_df[self.email_list_df['Main Dealer'].isin(daas_main_dealer)]

            # Find overlapping dealers
            overlapping_dealers = set(dealer_df['Main Dealer']).intersection(filter_daas["Main Dealer"])
            filter_email_both = self.email_list_df[self.email_list_df['Main Dealer'].isin(overlapping_dealers)]

            # Process files and send emails
            with os.scandir(base_path) as entries:
                for entry in entries:
                    if entry.is_file() and entry.name.endswith('.xlsx'):
                        dealer_name = self.extract_dealer_name(entry.name)

                        if dealer_name is None or dealer_name in self.processed_dealers:
                            continue

                        self._send_appropriate_email(dealer_name, overlapping_dealers,
                                                     filter_email_both, filter_email_daas,
                                                     filter_email, base_path)

        except Exception as e:
            logger.error(f"Error in email processing: {e}")

    def _send_appropriate_email(self, dealer_name, overlapping_dealers,
                                filter_email_both, filter_email_daas,
                                filter_email, base_path):
        """Send appropriate email based on dealer type"""
        if dealer_name in overlapping_dealers:
            # Send DaaS & MOXA email
            matching_rows = filter_email_both[filter_email_both['Main Dealer'] == dealer_name]
            for _, row in matching_rows.iterrows():
                self.send_email(row, dealer_name, 'DaaS & MOXA', base_path)
                self.processed_dealers.add(dealer_name)
                break
        else:
            # Check if DaaS only
            daas_rows = filter_email_daas[filter_email_daas['Main Dealer'] == dealer_name]
            if not daas_rows.empty:
                for _, row in daas_rows.iterrows():
                    self.send_email(row, dealer_name, 'DaaS', base_path)
                    self.processed_dealers.add(dealer_name)
                    break
            else:
                # MOXA only
                moxa_rows = filter_email[filter_email['Main Dealer'] == dealer_name]
                for _, row in moxa_rows.iterrows():
                    self.send_email(row, dealer_name, 'MOXA', base_path)
                    self.processed_dealers.add(dealer_name)
                    break

    def process_recap_file(self):
        """Process and update recap file"""
        try:
            df_recap = pd.read_excel(BASE_PATHS['recap'], parse_dates=['Tanggal Lahir'])
            df_daily = pd.read_excel(BASE_PATHS['data_file'], parse_dates=['Tanggal Lahir'])

            # Add required columns
            df_daily['Source Leads'] = 'FIF'
            df_daily['Platform Data'] = 'MOXA'

            # Filter and merge
            filter_daily = df_daily[df_daily['Main Dealer'] != 'blacklist']
            id_daily = filter_daily['Id Leads Data User'].tolist()
            df_merge = pd.concat([df_recap, df_daily])

            # Normalize phone numbers and update dispatch date
            df_merge['No HP'] = df_merge['No HP'].apply(self.normalize_phone_number)
            df_merge.loc[df_merge['Id Leads Data User'].isin(id_daily), "Dispatch Date"] = dispatch_date

            # Save merged data
            output_path = BASE_PATHS['folder'] / "Leads FIFGROUP Compile all MD v2.xlsx"
            self._save_recap_with_formatting(df_merge, output_path)

            logger.info(f"Recap file processed successfully: {output_path}")

        except Exception as e:
            logger.error(f"Error processing recap file: {e}")

    def _save_recap_with_formatting(self, df_merge, output_path):
        """Save recap file with specific formatting"""
        # Save with pandas first
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df_merge.to_excel(writer, sheet_name="concate", index=False)

        # Apply openpyxl formatting
        wb = op.load_workbook(output_path)
        ws = wb.active

        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        left_alignment = Alignment(horizontal='left')

        # Format cells and adjust column widths
        for col in range(1, 46):
            max_length = 0
            col_letter = ws.cell(row=1, column=col).column_letter

            for row in range(1, ws.max_row + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.alignment = left_alignment

                if isinstance(cell.value, datetime):
                    cell.number_format = 'DD/MM/YYYY'

                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))

            # Set column width
            ws.column_dimensions[col_letter].width = max_length + 2

        # Hide specified columns
        for col in HIDDEN_COLUMNS:
            ws.column_dimensions[col].hidden = True

        wb.save(output_path)


def main():
    """Main execution function"""
    # Create output directory
    os.makedirs(BASE_PATHS['base'], exist_ok=True)

    # Initialize processor
    processor = DataProcessor()

    if not processor.initialize_data():
        logger.error("Failed to initialize data. Exiting.")
        return

    # Filter data
    df_filtered, filter_daas = processor.filter_data()
    if df_filtered is None:
        logger.error("Failed to filter data. Exiting.")
        return

    # Process main data
    df_processed = processor.process_main_data(df_filtered)
    if df_processed is None:
        logger.error("Failed to process main data. Exiting.")
        return

    # Save main combined file
    main_output_path = BASE_PATHS['base'] / f"DATA GABUNGAN LEADS FIFGROUP {current_date}.xlsx"
    if not processor.save_excel_with_formatting(df_processed, main_output_path):
        logger.error("Failed to save main file. Exiting.")
        return

    # Split data by dealer
    processor.split_data_by_dealer(df_processed, BASE_PATHS['base'])

    # Process email sending
    processor.process_email_sending(BASE_PATHS['base'], df_processed, filter_daas)

    # Process recap file
    processor.process_recap_file()

    logger.info("All processing completed successfully!")


if __name__ == "__main__":
    main()
