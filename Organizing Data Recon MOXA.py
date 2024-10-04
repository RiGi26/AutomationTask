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

def adjust_column_width_and_format(filepath, *sheets, font_name='Calibri', font_size=11):
    workbook = op.load_workbook(filepath)
    for sheet_name in sheets:
        sheet = workbook[sheet_name]

        font_style = Font(name=font_name, size=font_size)
        alignment_style = Alignment(horizontal='left')

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Loop untuk setiap kolom
        for column_cells in sheet.columns:
            max_length = 0
            column = column_cells[0].column_letter
            
            for cell in column_cells:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                        cell.font = font_style
                        cell.alignment = alignment_style
                        cell.border = thin_border
                except:
                    pass

            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column].width = adjusted_width

    workbook.save(filepath)
    
    print(f"columns was adjusted")

with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    for sheet in sheets:
        df = pd.read_excel(file_path, sheet_name=sheet)

        if 'Bulan Booking' in df.columns:
            df = df[df['Bulan Booking'].isna()]

        if 'Transaction' in df.columns:
            df['Transaction'] = pd.to_datetime(df['Transaction'], errors='coerce')
            df = df[df['Transaction'] >= pd.to_datetime("2024-06-01")]
        
        if 'Phone' in df.columns:
            df['Phone'] = df['Phone'].astype(str).apply(lambda x: '0' + x if not x.startswith('0') else x)
            
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

        df.to_excel(writer, sheet_name=sheet, index=False)
        
        print(f"Sheet '{sheet}' processed and saved on {output_file_path}")

adjust_column_width_and_format(output_file_path, 'NMC', 'NMC SY', 'REFI', 'REFI SY')
