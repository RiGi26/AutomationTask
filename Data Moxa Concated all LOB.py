import pandas as pd
import os

folder_month = "September"

# Paths
file_refi = "C:\\Users\\61140\\Downloads\\exportDanastra (58).xlsx"
file_nmc = "D:\\Daily MOXA\\Data Leads 2023.xlsx"
sheets = ["NMC", "NMC SY"]

# Output file path
output_directory = f"D:\\Cross Selling\\Moxa\\Booking\\2024\\{folder_month}"
output_file_path = os.path.join(output_directory, "MOXA 20240930.xlsx")

# Ensure the output directory exists
os.makedirs(output_directory, exist_ok=True)

# Column mapping for the refi file
Mapping_column = {
    "Lead ID": "Id User Profile",
    "Digital Lead Id": "Id Leads Data User",
    "Fullname": "Name",
    "Mobile Phone1": "Phone",
    "No KTP": "Nomor KTP",
    "Submit Date": "Transaction"
    }

with pd.ExcelWriter(output_file_path, engine="xlsxwriter") as writer:
    for sheet in sheets:
        df_nmc = pd.read_excel(file_nmc, sheet_name=sheet)

        if "Transaction" in df_nmc.columns:
            df_nmc["Transaction"] = pd.to_datetime(
                df_nmc["Transaction"], errors="coerce"
            )
            df_nmc = df_nmc[
                (df_nmc["Transaction"] >= pd.to_datetime("2024-09-01"))
                & (df_nmc["Transaction"] <= pd.to_datetime("2024-09-30"))
            ]

        if "Phone" in df_nmc.columns:
            df_nmc["Phone"] = (
                df_nmc["Phone"]
                .astype(str)
                .apply(lambda x: "0" + x if not x.startswith("0") else x)
            )

        if "Nomor KTP" in df_nmc.columns:
            df_nmc["Nomor KTP"] = df_nmc["Nomor KTP"].astype(str).fillna("")

        df_nmc = df_nmc.drop_duplicates(
            subset=["Id User Profile", "Phone", "Nomor KTP"]
        )
        
        df_nmc['LOB'] = sheet

        df_nmc.to_excel(writer, sheet_name=sheet, index=False)
        
    df_refi = pd.read_excel(file_refi)
    df_refi = df_refi[list(Mapping_column.keys())].rename(columns=Mapping_column)

    if "Transaction" in df_refi.columns:
        df_refi["Transaction"] = pd.to_datetime(df_refi["Transaction"], errors="coerce")

    if "Phone" in df_refi.columns:
        df_refi["Phone"] = (
            df_refi["Phone"]
            .astype(str)
            .apply(lambda x: "0" + x if not x.startswith("0") else x)
        )

    if "Nomor KTP" in df_refi.columns:
        df_refi["Nomor KTP"] = df_refi["Nomor KTP"].astype(str).fillna("")

    df_refi = df_refi.drop_duplicates(
        subset=["Id Leads Data User", "Phone", "Nomor KTP"]
    )
    
    df_refi['LOB'] = "REFI"

    df_refi.to_excel(writer, sheet_name="REFI", index=False)

    print(f"File has been generated at {output_directory} {output_file_path}")

    writer.book.add_worksheet("SUMMARY")
    writer.book.add_worksheet("BOOKING")
