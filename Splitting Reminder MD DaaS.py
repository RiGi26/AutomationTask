import os
from datetime import datetime
import pandas as pd
import numpy as np

# setting
pd.set_option("display.max_columns", None)

# Path where folder has to change
current_date = datetime.now().strftime("%Y%m%d")

# Path where the output files will be saved
output_dir = f"D:\\Daily MOXA\\Data Reminder DaaS"

file = "D:\\Daily MOXA\\DAAS\\Rekap DAAS Februari 2023.xlsx"
recap_data = pd.read_excel(file)

# filtering
filter_feedback = recap_data[
    recap_data["Update Status Date"].isna() & recap_data["Dispatch Date"].notna()
]
# fds = filter_feedback.query("`Dispatch Date` >= '2023-01-01'")
fds = filter_feedback[
    (filter_feedback["Dispatch Date"] >= pd.to_datetime("2024/01/01"))
    & (filter_feedback["Dispatch Date"] <= pd.to_datetime("2024/09/25"))
]
column_to_include = [
    "id",
    "Nama",
    "Gender",
    "Alamat",
    "Kelurahan",
    "Kecamatan",
    "Kota/Kabupaten",
    "Propinsi",
    "No HP",
    "No Hp-2",
    "Sales Date",
    "Varian Motor",
    "Main Dealer",
    "Assign Dealer Code (5 DIGIT)",
    "Propensity",
    "Pekerjaan",
    "Pendidikan",
    "Pengeluaran",
    "Agama",
    "Tanggal Lahir",
    "Frame No Terakhir",
    "Jenis Penjualan",
    "Sales ID",
    "Nama Leasing Sebelumnya",
    "Nama salesman",
    "Source Leads",
    "Platform Data",
    "Dealer Sebelumnya (Jika Ada)",
    "Remarks/Keterangan",
    "Rekomendasi DP/Angsuran (Tenure)",
    "Varian motor yang diinginkan",
    "Warna varian motor",
    "E-MAIL"
]
final_data = fds[column_to_include].copy()
final_data["No HP"] = (
            final_data["No HP"]
            .astype("str")
            .apply(lambda x: "0" + x if not x.startswith("0") else x)
        )
output_path = os.path.join(output_dir, f"Remainder Data Leads Master DaaS.xlsx")
final_data.to_excel(output_path, index=False)

unique_values = final_data["Main Dealer"].unique()

# Column mappings and final column orders
kolom_mpm = {
    "id": "id",
    "Nama": "Nama",
    "No HP": "No HP",
    "Kota/Kabupaten": "Kota/Kabupaten",
    "Kecamatan": "Kecamatan",
    "Alamat": "Alamat",
}

kolom_akhir = [
    "id",
    "Nama",
    "No HP",
    "Kota/Kabupaten",
    "Kode Dealer Refrensi",
    "Alamat",
    "Kelurahan",
    "Kecamatan",
]

def process_data_for_dealer_DaaS(dealer_name, final_data):
    if dealer_name in ["PT MPM - MALANG", "PT MPM - SURABAYA"]:
        df_pindah = final_data[list(kolom_mpm.values())].rename(columns=kolom_mpm).copy()
        df_final = df_pindah.reindex(columns=kolom_akhir, fill_value=np.nan)
    else:
        df_final = final_data[final_data["Main Dealer"] == dealer_name].copy()

    df_final["No HP"] = df_final["No HP"].astype(str).apply(lambda x: "0" + x if not x.startswith("0") else x)
    output_path = os.path.join(output_dir, f"Remainder Data Leads {dealer_name} DaaS.xlsx")
    write_to_excel(df_final, output_path)
    print(f"File has been created for {dealer_name} DaaS")
    
def write_to_excel(dataframe, path):
    # Replace NaN values with an empty string or another placeholder
    dataframe = dataframe.fillna('')

    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        dataframe.to_excel(writer, sheet_name="Sheet1", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]
        border_format = workbook.add_format({"border": 1})
        date_format = workbook.add_format({'num_format':'dd-mm-yyyy', 'border': 1})

        # Apply border and adjust column width
        for row_num in range(len(dataframe) + 1):
            for col_num, col in enumerate(dataframe.columns):
                if row_num == 0:
                    value = col  # Header
                    worksheet.write(row_num, col_num, value, border_format)
                else:
                    value = dataframe.iloc[row_num - 1, col_num]
                    
                    if col == 'Dispatch Date':
                        worksheet.write(row_num, col_num, value, date_format)
                    else:
                        worksheet.write(row_num, col_num, value, border_format)
                
        
        # Adjust column width
        for idx, col in enumerate(dataframe.columns):
            max_len = max(dataframe[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(idx, idx, max_len)

for unique in unique_values:
    process_data_for_dealer_DaaS(unique, final_data)

