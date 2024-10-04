import pandas as pd
import os
from datetime import datetime
import numpy as np


# Path where folder has to change
folder_date = "25"
folder_month = "September"
current_date = datetime.now().strftime("%Y%m%d")
filter_date = datetime.now().strftime("%d %B %Y")

# Path to the Excel file
file = f"D:\Daily MOXA\DAAS\Rekap DAAS Februari 2023.xlsx"
df = pd.read_excel(file)

# Path where the output files will be saved
output_dir = (
    f"D:\\Daily MOXA\\blackup kirim dealer\\2024\\{folder_month}\\{folder_date}"
)

# Get unique values from the specified column
df_filter = df[
    df["Dispatch Date"] == pd.to_datetime(filter_date)
]  # Double Check when daily late
unique_values = df_filter["Main Dealer"].unique()

# Create the output directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

# Column mappings and final column orders
columns = [
    "CUST_NO",
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
    "E-MAIL",
]

kolom_mpm = {
    "CUST_NO": "CUST_NO",
    "Nama": "Nama",
    "No HP": "No HP",
    "Kota/Kabupaten": "Kota/Kabupaten",
    "Kecamatan": "Kecamatan",
    "Alamat": "Alamat",
    "Kelurahan": "Kelurahan",
}

kolom_akhir = [
    "CUST_NO",
    "Nama",
    "No HP",
    "Kota/Kabupaten",
    "Kode Dealer Refrensi",
    "Alamat",
    "Kelurahan",
    "Kecamatan",
]

df_final = df_filter[columns]

# Loop through each unique value
for unique in unique_values:
    if unique in ["PT MPM - MALANG", "PT MPM - SURABAYA"]:
        # Select and rename columns for MPM dealers
        df_pindah = df_final[list(kolom_mpm.values())].rename(columns=kolom_mpm).copy()

        # Add missing columns if necessary and reorder
        df_final = df_pindah.reindex(columns=kolom_akhir, fill_value=np.nan)
        df_final["No HP"] = (
            df_final["No HP"]
            .astype("str")
            .apply(lambda x: "0" + x if not x.startswith("0") else x)
        )

        # Save to Excel
        output_path = os.path.join(
            output_dir, f"Data Leads FIFGROUP {current_date} {unique} DaaS.xlsx"
        )
        df_final.to_excel(output_path, index=False)

        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df_final = df_final.fillna("")

            df_final.to_excel(writer, sheet_name="Sheet1", index=False)
            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets["Sheet1"]
            border_format = workbook.add_format({"border": 1})
            date_format = workbook.add_format({"num_format": "dd-mm-yyyy", "border": 1})

            for row_num in range(len(df_final) + 1):
                for col_num, col in enumerate(df_final.columns):
                    if row_num == 0:
                        value = col  # Header
                        worksheet.write(row_num, col_num, value, border_format)
                    else:
                        value = df_final.iloc[row_num - 1, col_num]

                        if col == "Dispatch Date":
                            worksheet.write(row_num, col_num, value, date_format)
                        else:
                            worksheet.write(row_num, col_num, value, border_format)

            # Adjust the column width to fit the content
            for idx, col in enumerate(df_final.columns):
                # Find the maximum length of the column values
                max_len = (
                    max(df_final[col].astype(str).map(len).max(), len(col)) + 2
                )  # Add some extra space for readability
                worksheet.set_column(idx, idx, max_len)

            print(f"Files saved to: {output_dir}")

    else:
        # Filter data for other main dealers
        df_output = df_final[df_filter["Main Dealer"] == unique].copy()

        # Ensure No HP starts with '0'
        df_output["No HP"] = (
            df_output["No HP"]
            .astype(str)
            .apply(lambda x: "0" + x if not x.startswith("0") else x)
        )

        # Save to Excel
        output_path = os.path.join(
            output_dir, f"Data Leads FIFGROUP {current_date} {unique} DaaS.xlsx"
        )
        df_output.to_excel(output_path, index=False)

        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df_output = df_output.fillna("")

            df_output.to_excel(writer, sheet_name="Sheet1", index=False)
            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets["Sheet1"]
            border_format = workbook.add_format({"border": 1})
            date_format = workbook.add_format({"num_format": "dd-mm-yyyy", "border": 1})

            for row_num in range(len(df_output) + 1):
                for col_num, col in enumerate(df_output.columns):
                    if row_num == 0:
                        value = col  # Header
                        worksheet.write(row_num, col_num, value, border_format)
                    else:
                        value = df_output.iloc[row_num - 1, col_num]

                        if col == "Dispatch Date":
                            worksheet.write(row_num, col_num, value, date_format)
                        else:
                            worksheet.write(row_num, col_num, value, border_format)

            # Adjust the column width to fit the content
            for idx, col in enumerate(df_output.columns):
                # Find the maximum length of the column values
                max_len = (
                    max(df_output[col].astype(str).map(len).max(), len(col)) + 2
                )  # Add some extra space for readability
                worksheet.set_column(idx, idx, max_len)

        print(f"Files saved to: {output_dir}")
