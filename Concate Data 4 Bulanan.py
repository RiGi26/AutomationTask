import pandas as pd
import numpy as np
import openpyxl as op
from openpyxl.styles import Alignment, Font, Border, Side
from xlsxwriter import Workbook
import os

# Change it
month_path = "20240930"
folder_month = "September"

# path
file_month = f"D:\\Cross Selling\\Moxa\\Booking\\2024\\{folder_month}\\MOXA {month_path}.xlsx"
file_3_months = f"D:\\Cross Selling\\Moxa\\Booking\\2024\\{folder_month}\\Rekap Moxa.xlsx"
sheets = ['NMC', 'NMC SY', 'REFI', 'REFI SY']

# location
output_dir = f"D:\\Cross Selling\\Moxa\\Booking\\2024\\{folder_month}"
output_file_path = os.path.join(output_dir, "Data4Bulan.xlsx")
os.makedirs(output_dir, exist_ok=True)

# mapping
mapping = ["Id User Profile", "Id Leads Data User", "Name", "Phone", "Nomor KTP", "Transaction", "LOB"]
mapping_phone_ktp = ["Phone", "Nomor KTP"]

# combine
nmc_1 = pd.read_excel(file_3_months, sheet_name="NMC")
nmc_2 = pd.read_excel(file_month, sheet_name="NMC")
nmcsy_1 = pd.read_excel(file_3_months, sheet_name="NMC SY")
nmcsy_2 = pd.read_excel(file_month, sheet_name="NMC SY")
refi_1 = pd.read_excel(file_3_months, sheet_name="REFI")
refi_2 = pd.read_excel(file_month, sheet_name="REFI")
refisy_1 = pd.read_excel(file_3_months, sheet_name="REFI SY")

def form(data):
    if 'Phone' in data.columns:
        data['Phone'] = data['Phone'].astype(str).apply(lambda x: '0' + x if not x.startswith('0') else x)
    if 'Nomor KTP' in data.columns:
        data['Nomor KTP'] = data['Nomor KTP'].astype(str).apply(lambda x: x if x else "")
    return data

combine_nmc = pd.concat([nmc_1, nmc_2], ignore_index=True)[mapping]
combine_nmc = form(combine_nmc)

combine_nmcsy = pd.concat([nmcsy_1, nmcsy_2], ignore_index=True)[mapping]
combine_nmcsy = form(combine_nmcsy)

combine_refi = pd.concat([refi_1, refi_2], ignore_index=True)[mapping]
combine_refi = form(combine_refi)

refisy_1 = refisy_1[mapping]
refisy_1 = form(refisy_1)

combine_all = pd.concat([combine_nmc, combine_nmcsy, combine_refi, refisy_1], ignore_index=True)
combine_all = form(combine_all)

data_matching = combine_all[mapping_phone_ktp]

with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    combine_all = combine_all.drop_duplicates(subset=["Id User Profile", "Phone", "Nomor KTP"])
    combine_all.to_excel(writer, index=False, sheet_name="Gabungan")
    data_matching.to_excel(writer, index=False, sheet_name="Data Matching")

print(f"File telah berhasil dibuat di {output_file_path}")
    
def adjust_column_width_and_format(filepath, *sheet_names, font_name='Calibri', font_size=11):
    workbook = op.load_workbook(filepath)
    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]
        
        font_style = Font(name=font_name, size=font_size)
        alignment_style = Alignment(horizontal='left')  # Rata kiri

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

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

adjust_column_width_and_format(output_file_path, 'Gabungan', 'Data Matching')    
