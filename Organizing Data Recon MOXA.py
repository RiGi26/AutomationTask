import pandas as pd
import openpyxl as op
from openpyxl.styles import Alignment, Font, Border, Side

# File path and sheets to be processed
file_path = "D:\\Cross Selling\\Moxa\\Booking\\recap leads all New v7.xlsx"
sheets = ['NMC', 'NMC SY', 'REFI', 'REFI SY']
folder_month = 'September'

# mapping REFI SY
mapp_refi = ["Id User Profile", "Id Leads Data User", "Name", "Phone", "Nomor KTP", "Transaction", "LOB"]

# Output file path
output_file_path = f"D:\\Cross Selling\\Moxa\\Booking\\2024\\{folder_month}\\Rekap Moxa.xlsx"

def adjust_column_width_and_format(filepath, *sheet_names, font_name='Calibri', font_size=11):
    # Membuka file Excel
    workbook = op.load_workbook(filepath)
    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]

        # Font yang akan digunakan
        font_style = Font(name=font_name, size=font_size)
        alignment_style = Alignment(horizontal='left')  # Rata kiri

        # Border yang akan digunakan
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Loop untuk setiap kolom
        for column_cells in sheet.columns:
            max_length = 0
            column = column_cells[0].column_letter  # Mendapatkan huruf kolom
            
            for cell in column_cells:
                try:
                    # Mendapatkan panjang maksimal dari setiap cell pada kolom tersebut
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                        # Mengatur perataan rata kiri, font, dan border untuk setiap cell
                        cell.font = font_style
                        cell.alignment = alignment_style
                        cell.border = thin_border
                except:
                    pass

            # Mengatur lebar kolom sesuai panjang maksimal
            adjusted_width = (max_length + 2)  # Tambahkan margin
            sheet.column_dimensions[column].width = adjusted_width

    # Simpan perubahan ke file yang sama
    workbook.save(filepath)

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    # Loop over each sheet, read, filter, and save to new sheet in output file
    for sheet in sheets:
        # Read each sheet into a DataFrame
        df = pd.read_excel(file_path, sheet_name=sheet)

        # Ensure 'Bulan Booking' column exists and filter for NaN values
        if 'Bulan Booking' in df.columns:
            df = df[df['Bulan Booking'].isna()]

        # Ensure 'Transaction' column is in datetime format before filtering
        if 'Transaction' in df.columns:
            df['Transaction'] = pd.to_datetime(df['Transaction'], errors='coerce')
            df = df[df['Transaction'] >= pd.to_datetime("2024-06-01")]
        
        # Ensure 'Phone' column is in string format and starts with '0'
        if 'Phone' in df.columns:
            df['Phone'] = df['Phone'].astype(str).apply(lambda x: '0' + x if not x.startswith('0') else x)
            
        # Ensure 'No. KTP' column is in string format
        if 'Nomor KTP' in df.columns:
            df['Nomor KTP'] = df['Nomor KTP'].astype(str).apply(lambda x: x if x else "")
        
        df['LOB'] = sheet
        
        if sheet == 'NMC':    
            df = df.drop_duplicates(subset=['Id User Profile', 'Phone', 'Nomor KTP'])
        elif sheet == 'NMC SY':
            df = df.drop_duplicates(subset=['Id User Profile', 'Phone', 'Nomor KTP'])
        elif sheet == 'REFI':
            df = df.drop_duplicates(subset=['Id User Profile', 'Phone', 'Nomor KTP'])
        elif sheet == 'REFI SY':
            df = df[mapp_refi]
        else:
            print(f"not found")

        # Write the filtered DataFrame to a new sheet in the output Excel file
        df.to_excel(writer, sheet_name=sheet, index=False)
        
        print(f"Sheet '{sheet}' processed and saved on {output_file_path}")

adjust_column_width_and_format(output_file_path, sheets)