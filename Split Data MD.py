import pandas as pd
import os
from datetime import datetime
import numpy as np
from pathlib import Path

# Path where folder has to change
folder_date = "30"
folder_month = "September"
current_date = datetime.now().strftime("%Y%m%d")
filter_date = datetime.now().strftime("%d %B %Y")

# Path to the Excel file Master and Extracted File
path = Path("D:\\Daily MOXA\\Master Leads Interest 2024.xlsx")
try:
    master = pd.read_excel(path, sheet_name="September")
except Exception as e:
    print(f"Failed to read master file: {e}")
    exit()

# Filtering data master
df_filtered = master[master["tgl"] == pd.to_datetime(filter_date)]# Double Check when daily task not sended on Time

# Path where the output files will be saved
output_dir = f"D:\\Daily MOXA\\blackup kirim dealer\\2024\\{folder_month}\\{folder_date}"
os.makedirs(output_dir, exist_ok=True) 

# Pemetaan kolom lama ke kolom baru
pemetaan_kolom = {
    "Id Leads Data User": "id",
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

# Kolom akhir yang diinginkan
kolom_akhir = [
    "id", "Nama", "Gender", "Alamat", "Kelurahan", "Kecamatan", "Kota/Kabupaten",
    "Propinsi", "No HP", "No Hp-2", "Sales Date", "Varian Motor", "Main Dealer",
    "Assign Dealer Code (5 DIGIT)", "Propensity", "Pekerjaan", "Pendidikan",
    "Pengeluaran", "Agama", "Tanggal Lahir", "Frame No Terakhir", "Jenis Penjualan",
    "Sales ID", "Nama Leasing Sebelumnya", "Nama salesman", "Source Leads",
    "Platform Data", "Dealer Sebelumnya (Jika Ada)", "Remarks/Keterangan",
    "Rekomendasi DP/Angsuran (Tenure)", "Varian motor yang diinginkan",
    "Warna varian motor", "E-MAIL", "FACEBOOK", "INSTAGRAM", "TWITTER"
]

# Pilih dan ganti nama kolom sesuai pemetaan
try:
    df_pindah = df_filtered[list(pemetaan_kolom.keys())].rename(columns=pemetaan_kolom)
    df_pindah["No HP"] = df_pindah["No HP"].astype(str).apply(lambda x: "0" + x if not x.startswith("0") else x)
except:
    print("Failed to select and rename columns")

try:
    # Tambahkan kolom yang tidak ada di DataFrame asli dan isi dengan NaN
    for kolom in kolom_akhir:
        if kolom not in df_pindah.columns:
            df_pindah[kolom] = np.nan
            
    # Urutkan kolom sesuai dengan kolom_akhir       
    df_pindah = df_pindah[kolom_akhir]
except:
    print("Failed to add missing columns")

# Construct the full file path and filename
output_file_path = os.path.join(output_dir, f"DATA GABUNGAN LEADS FIFGROUP {current_date}.xlsx")

# Save the resulting DataFrame to a new Excel file without formatting
try:
    with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
        df_pindah = df_pindah.fillna('')
        
        df_pindah.to_excel(writer, sheet_name='Sheet1',index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        border_format = workbook.add_format({'border': 1})
        date_format = workbook.add_format({'num_format': 'dd-mm-yyyy', 'border': 1})
        
        for row_num in range(len(df_pindah) + 1):
            for col_num, col in enumerate(df_pindah.columns):
                if row_num == 0:
                    value = col  # Header
                    worksheet.write(row_num, col_num, value, border_format)
                else:
                    value = df_pindah.iloc[row_num - 1, col_num]
                    
                    if col == 'Dispatch Date':
                        worksheet.write(row_num, col_num, value, date_format)
                    else:
                        worksheet.write(row_num, col_num, value, border_format)
        
        for idx, col in enumerate(df_pindah.columns):
            max_len = max(df_pindah[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(idx, idx, max_len)
            
    print(f"File disimpan dengan format yang sama di {output_file_path}")
except Exception as e:
    print(f"Failed to save the output file: {e}")

# Reformating extracted file

file = f"D:\\Daily MOXA\\blackup kirim dealer\\2024\\{folder_month}\\{folder_date}\\DATA GABUNGAN LEADS FIFGROUP {current_date}.xlsx"
try:
    extracted_file = pd.read_excel(file)
except Exception as e:
    print(f"Failed to read extracted file: {e}")
    exit()

# Get unique values from the specified column
unique_values = extracted_file["Main Dealer"].unique()

# Loop through each unique value
for unique in unique_values:
    if unique in ["PT MPM - MALANG", "PT MPM - SURABAYA"]:
        kolom_mpm = {
            "id": "id",
            "Nama": "Nama",
            "No HP": "No HP",
            "Kota/Kabupaten": "Kota/Kabupaten",
            "Kecamatan": "Kecamatan",
            "Alamat": "Alamat",
        }
        
        df_pindah = extracted_file[list(kolom_mpm.values())].rename(columns=kolom_mpm).copy()
        df_final = df_pindah.reindex(columns=kolom_akhir, fill_value=np.nan)
        df_final["No HP"] = df_final["No HP"].astype(str).apply(lambda x: "0" + x if not x.startswith("0") else x)

        output_path = os.path.join(output_dir, f"Data Leads FIFGROUP {current_date} {unique}.xlsx")
        try:
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                df_final.to_excel(writer, sheet_name="Sheet1", index=False)
                worksheet = writer.sheets["Sheet1"]
                for idx, col in enumerate(df_final.columns):
                    max_len = max(df_final[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(idx, idx, max_len)
            print(f"File has been Splitting {unique}")
        except Exception as e:
            print(f"Failed to save the file for {unique}: {e}")
    else:
        df_output = extracted_file[extracted_file["Main Dealer"] == unique].copy()
        df_output["No HP"] = df_output["No HP"].astype(str).apply(lambda x: "0" + x if not x.startswith("0") else x)

        output_path = os.path.join(output_dir, f"Data Leads FIFGROUP {current_date} {unique}.xlsx")
        try:
            with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
                df_output.to_excel(writer, sheet_name="Sheet1", index=False)
                worksheet = writer.sheets["Sheet1"]
                for idx, col in enumerate(df_output.columns):
                    max_len = max(df_output[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(idx, idx, max_len)
            print(f"File has been Splitting {unique}")
        except Exception as e:
            print(f"Failed to save the file for {unique}: {e}")

