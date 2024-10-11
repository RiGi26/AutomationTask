import pandas as pd
import numpy as np
import os
from pathlib import Path
from datetime import datetime
import datetime as dt

# setting display
pd.set_option("display.max_columns", None)

# set date
filter_date = datetime.now().strftime("%d %B %Y")
format_date = "20241011" # Change it

# directory
output_dir = Path("D:\\Daily MOXA\DL x HAYATI")

mapping_hayati = [
    "Kode Mitra",
    "Source Input",
    "Nama Lengkap",
    "Nomor KTP",
    "Tanggal Lahir",
    "Tempat Lahir",
    "Jenis Kelamin",
    "Nama Ibu Kandung",
    "Alamat KTP",
    "RT KTP",
    "RW KTP",
    "Kode Provinsi KTP",
    "Provinsi KTP",
    "Kode Kota KTP",
    "Kota KTP",
    "Kode Kecamatan KTP",
    "Kecamatan KTP",
    "Kode Kelurahan KTP",
    "Kelurahan KTP",
    "Kode Pos KTP",
    "Subzip KTP",
    "Alamat Domisili",
    "RT Domisili",
    "RW Domisili",
    "Kode Provinsi Domisili",
    "Provinsi Domisili",
    "Kode Kota Domisili",
    "Kota Domisili",
    "Kode Kecamatan Domisili",
    "Kecamatan Domisili",
    "Kode Kelurahan Domisili",
    "Kelurahan Domisili",
    "Kode Pos Domisili",
    "Subzip Domisili",
    "No HP 1",
    "No HP 2",
    "Email",
    "Status Pernikahan",
    "Status Rumah",
    "Pendidikan Terakhir",
    "Kepemilikan NPWP",
    "Nomor NPWP",
    "Tipe Pekerjaan",
    "Status Pekerjaan",
    "Jabatan Pekerjaan",
    "Tipe Pekerjaan Pasangan",
    "Status Pekerjaan Pasangan",
    "Jabatan Pekerjaan Pasangan",
    "Penghasilan per Bulan",
    "Pengeluaran per Bulan",
    "Kepemilikan Buku Tabungan",
    "Bank",
    "Nomor Rekening",
    "Kepemilikan Kartu Kredit",
    "Bank Penerbit Kartu Kredit",
    "Nomor Kartu Kredit",
    "Tipe Unit",
    "Tenor",
    "Pengiriman Kendaraan",
    "Nama yang Akan Tercantum Sebagai Pemilik BPKB",
    "Nama Pemilik Kendaraan",
    "Alamat Pemilik Kendaraan",
    "RT Pemilik Kendaraan",
    "RW Pemilik Kendaraan",
    "Kode Provinsi Pemilik Kendaraan",
    "Provinsi Pemilik Kendaraan",
    "Kode Kota Pemilik Kendaraan",
    "Kota Pemilik Kendaraan",
    "Kode Kecamatan Pemilik Kendaraan",
    "Kecamatan Pemilik Kendaraan",
    "Kode Kelurahan Pemilik Kendaraan",
    "Kelurahan Pemilik Kendaraan",
    "Kode Pos Pemilik Kendaraan",
    "Subzip Pemilik Kendaraan",
    "Call Via",
    "Contact Detail",
    "Interest? (Yes/No)",
]

source = {
    "Nama": "Nama Lengkap",
    "NIK": "Nomor KTP",
    "Tanggal Lahir_x": "Tanggal Lahir",
    "Tempat Lahir": "Tempat Lahir",
    "Gender": "Jenis Kelamin",
    "Alamat KTP": "Alamat KTP",
    "Kode Provinsi KTP": "Kode Provinsi KTP",
    "Provinsi KTP": "Provinsi KTP",
    "Kode Kota KTP": "Kode Kota KTP",
    "Kota KTP": "Kota KTP",
    "Kode Kecamatan KTP": "Kode Kecamatan KTP",
    "Kecamatan KTP": "Kecamatan KTP",
    "Kode Kelurahan KTP": "Kode Kelurahan KTP",
    "Kelurahan KTP": "Kelurahan KTP",
    "Kode Pos KTP": "Kode Pos KTP",
    "Subzip KTP": "Subzip KTP",
    "Alamat Domisili": "Alamat Domisili",
    "Kode Provinsi Domisili": "Kode Provinsi Domisili",
    "Provinsi Domisili": "Provinsi Domisili",
    "Kode Kota Domisili": "Kode Kota Domisili",
    "Kota Domisili": "Kota Domisili",
    "Kode Kecamatan Domisili": "Kode Kecamatan Domisili",
    "Kecamatan Domisili": "Kecamatan Domisili",
    "Kode Kelurahan Domisili": "Kode Kelurahan Domisili",
    "Kelurahan Domisili": "Kelurahan Domisili",
    "Kode Pos Domisili": "Kode Pos Domisili",
    "Subzip Domisili": "Subzip Domisili",
    "No HP 1": "No HP 1",
    "E-MAIL": "Email",
    "Status Pernikahan": "Status Pernikahan",
    "Status Kepemilikan Rumah": "Status Rumah",
    "Pendidikan Terakhir": "Pendidikan Terakhir",
    "Tipe Motor": "Tipe Unit",
    "Nama Pemilik Kendaraan": "Nama Pemilik Kendaraan",
    "Alamat": "Alamat Pemilik Kendaraan",
    "PROVINSICODE": "Kode Provinsi Pemilik Kendaraan",
    "PROVINSIDESC": "Provinsi Pemilik Kendaraan",
    "CITYCODE": "Kode Kota Pemilik Kendaraan",
    "CITYDESC": "Kota Pemilik Kendaraan",
    "KECAMATANCODE": "Kode Kecamatan Pemilik Kendaraan",
    "KECAMATANDESC": "Kecamatan Pemilik Kendaraan",
    "KELURAHANCODE": "Kode Kelurahan Pemilik Kendaraan",
    "KELURAHANDESC": "Kelurahan Pemilik Kendaraan",
    "ZIPCODE": "Kode Pos Pemilik Kendaraan",
    "Sub Zip": "Subzip Pemilik Kendaraan",
    "No HP": "Contact Detail"
}

status_mapping = {
    "Sd": "SD",
    "Smp": "SP",
    "Sma": "SA",
    "Sarjana": "S1",
    "Master": "S2",
    "Doctor": "S3",
    "Lain-Lain": "ZZ",
    "Diploma/Politehnik": "SO",
    "Rumah Sendiri": "H01",
    "Rumah Keluarga": "H02",
    "Rumah Dinas": "H03",
    "Kontrak": "H04",
    "Kost": "H05",
    "Kredit": "H06",
    "Lajang": "S",
    "Menikah": "M",
    "Cerai": "D",
    "Duda/Janda": "D",
    "Pria": "M",
    "Wanita": "F"
}

mapping = [
    "Id Leads Data User",
    "Tanggal Lahir",
    "Status Pernikahan",
    "Status Kepemilikan Rumah",
    "Pendidikan Terakhir",
    "Tipe Motor"
]

# path 
data_interest = Path("D:\\Daily MOXA\\Master Leads Interest 2024.xlsx")
database = Path("D:\\Daily MOXA\\Data Leads 2023.xlsx")
data_lookup = Path("D:\\Daily MOXA\DL x HAYATI\\FIFASTRA MOXA 2023 template.xlsx")

# reading data
df_i = pd.read_excel(data_interest, sheet_name="Oktober")
df_d = pd.read_excel(database, sheet_name="NMC")
df_l = pd.read_excel(data_lookup, sheet_name="LOOKUP - Alamat")

# data interest hayati
hayati = df_i[df_i["MD (3 DIGIT)"] == "CV HAYATI"].copy()
unique_id = hayati["Id Leads Data User"].tolist()

# data raw leads
data_personal = df_d[df_d["Id Leads Data User"].isin(unique_id)][mapping]

# data template
zipcode = hayati["Kode Pos"].tolist()
alamat = df_l[df_l["ZIPCODE"].isin(zipcode)].drop_duplicates(
    subset="ZIPCODE", keep="first"
)

# merging data
result_1 = hayati.merge(alamat, left_on="Kode Pos", right_on="ZIPCODE", how="left")
result_2 = result_1.merge(
    data_personal,
    how="left",
    left_on="Id Leads Data User",
    right_on="Id Leads Data User",
)
result_2[["Status Pernikahan", "Status Kepemilikan Rumah", "Pendidikan Terakhir"]] = (
    result_2[
        ["Status Pernikahan", "Status Kepemilikan Rumah", "Pendidikan Terakhir"]
    ].apply(lambda x: x.str.title() if x.dtype == "object" else x)
)
result_2[["Gender", "Status Pernikahan", "Status Kepemilikan Rumah", "Pendidikan Terakhir"]] = (
    result_2[
        ["Gender", "Status Pernikahan", "Status Kepemilikan Rumah", "Pendidikan Terakhir"]
    ].replace(status_mapping)
)

# duplicating data
result_2['Alamat KTP'] = result_2['Alamat']
result_2['Kode Provinsi KTP'] = result_2['PROVINSICODE']
result_2['Provinsi KTP'] = result_2['PROVINSIDESC']
result_2['Kode Kota KTP'] = result_2['CITYCODE']
result_2['Kota KTP'] = result_2['CITYDESC']
result_2['Kode Kecamatan KTP'] = result_2["KECAMATANCODE"]
result_2['Kecamatan KTP'] = result_2['KECAMATANDESC']
result_2['Kode Kelurahan KTP'] = result_2['KELURAHANCODE']
result_2['Kelurahan KTP'] = result_2['KELURAHANDESC']
result_2['Kode Pos KTP'] = result_2['ZIPCODE']
result_2['Subzip KTP'] = result_2['Sub Zip']

result_2['Alamat Domisili'] = result_2['Alamat']
result_2['Kode Provinsi Domisili'] = result_2['PROVINSICODE']
result_2['Provinsi Domisili'] = result_2['PROVINSIDESC']
result_2['Kode Kota Domisili'] = result_2['CITYCODE']
result_2['Kota Domisili'] = result_2['CITYDESC']
result_2['Kode Kecamatan Domisili'] = result_2["KECAMATANCODE"]
result_2['Kecamatan Domisili'] = result_2['KECAMATANDESC']
result_2['Kode Kelurahan Domisili'] = result_2['KELURAHANCODE']
result_2['Kelurahan Domisili'] = result_2['KELURAHANDESC']
result_2['Kode Pos Domisili'] = result_2['ZIPCODE']
result_2['Subzip Domisili'] = result_2['Sub Zip']
result_2['Nama Pemilik Kendaraan'] = result_2['Nama']
result_2['No HP 1'] = result_2['No HP']
result_2['Tempat Lahir'] = result_2['Kota KTP']
result_2['Tanggal Lahir_x'] = pd.to_datetime(result_2['Tanggal Lahir_x'], errors='coerce').dt.strftime('%d-%m-%Y')

data_final = result_2[list(source.keys())].rename(columns=source)

for kolom in mapping_hayati:
    if kolom not in data_final.columns:
        data_final[kolom] = np.nan
        
data_final = data_final[mapping_hayati]

# setting default 
data_final[["RT KTP","RW KTP","RT Domisili","RW Domisili", "RT Pemilik Kendaraan", "RW Pemilik Kendaraan"]] = "001"
data_final["Tipe Pekerjaan"] = "1"
data_final["Status Pekerjaan"] = "04"
data_final["Jabatan Pekerjaan"] = "002"
data_final["Penghasilan per Bulan"] = "0"
data_final["Pengeluaran per Bulan"] = "0"
data_final["Tenor"] = "12"
data_final["Pengiriman Kendaraan"] = "01"
data_final['Source Input'] = 'MXA'
data_final['Nama Ibu Kandung'] = 'IBU'
data_final['Kepemilikan Kartu Kredit'] = "False"
data_final['Kepemilikan Buku Tabungan'] = "False"
data_final['Call Via'] = 'WA'
data_final['Main Dealer Name'] = "Hayati"
data_final['Dealer Code'] = "2140030"
data_final['No HP 2'] = "-" 

column_edit = ['No HP 1', 'Contact Detail', 'Kode Provinsi KTP', 'Kode Kota KTP', 'Kode Kecamatan KTP', 'Kode Kelurahan KTP', 'Kode Provinsi Domisili', 'Kode Kota Domisili', 'Kode Kecamatan Domisili', 'Kode Kelurahan Domisili' ,'Kode Provinsi Pemilik Kendaraan', 'Kode Kota Pemilik Kendaraan', 'Kode Kecamatan Pemilik Kendaraan', 'Kode Kelurahan Pemilik Kendaraan']

# setting Nomor HP dan KTP
data_final['Nomor KTP'] = data_final['Nomor KTP'].astype(str).apply(lambda x: x if x else "")
for edit in column_edit:
    data_final[edit] = data_final[edit].astype(str).apply(lambda x : "0" + x if not x.startswith("0") else x)

output_path_file = os.path.join(output_dir, f"FIFASTRA MOXA 2024 - HAYATI {format_date}.xlsx")

print(data_final)

with pd.ExcelWriter(output_path_file, engine='xlsxwriter') as writer:
    data_final.to_excel(writer, sheet_name="HAYATI", index=False)
    workbook = writer.book
    worksheet = writer.sheets['HAYATI']
    border_format = workbook.add_format({'border': 1})
    
    for row_num in range(len(data_final) + 1):
        data_final = data_final.fillna(' ')
        for col_num, col in enumerate(data_final.columns):
            if row_num == 0:
                value = col  # Header
                worksheet.write(row_num, col_num, value, border_format)
            else:
                value = data_final.iloc[row_num - 1, col_num]
                worksheet.write(row_num, col_num, value, border_format)
    
    for idx, col in enumerate(data_final.columns):
        max_len = max(data_final[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(idx, idx, max_len)

    print(f"Data Leads HAYATI has been created at {output_path_file}")