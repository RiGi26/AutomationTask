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

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter(output_file_path, engine="xlsxwriter") as writer:
    # Loop over each sheet in the NMC file, process, and save to new sheet in output file
    for sheet in sheets:
        df_nmc = pd.read_excel(file_nmc, sheet_name=sheet)

        # If filtering is needed, apply it here (example filter on 'Transaction' column)
        if "Transaction" in df_nmc.columns:
            df_nmc["Transaction"] = pd.to_datetime(
                df_nmc["Transaction"], errors="coerce"
            )
            df_nmc = df_nmc[
                (df_nmc["Transaction"] >= pd.to_datetime("2024-09-01"))
                & (df_nmc["Transaction"] <= pd.to_datetime("2024-09-30"))
            ]

        # Ensure 'Phone' and 'Nomor KTP' columns are properly formatted
        if "Phone" in df_nmc.columns:
            df_nmc["Phone"] = (
                df_nmc["Phone"]
                .astype(str)
                .apply(lambda x: "0" + x if not x.startswith("0") else x)
            )

        if "Nomor KTP" in df_nmc.columns:
            df_nmc["Nomor KTP"] = df_nmc["Nomor KTP"].astype(str).fillna("")

        # Remove duplicates based on specified columns
        df_nmc = df_nmc.drop_duplicates(
            subset=["Id User Profile", "Phone", "Nomor KTP"]
        )
        
        df_nmc['LOB'] = sheet

        # Write each processed DataFrame to its respective sheet in the output Excel file
        df_nmc.to_excel(writer, sheet_name=sheet, index=False)

    # Process the Refi file and write to a new sheet
    df_refi = pd.read_excel(file_refi)

    # Select and rename columns based on mapping
    df_refi = df_refi[list(Mapping_column.keys())].rename(columns=Mapping_column)

    if "Transaction" in df_refi.columns:
        df_refi["Transaction"] = pd.to_datetime(df_refi["Transaction"], errors="coerce")

    # Ensure 'Phone' and 'Nomor KTP' columns are properly formatted
    if "Phone" in df_refi.columns:
        df_refi["Phone"] = (
            df_refi["Phone"]
            .astype(str)
            .apply(lambda x: "0" + x if not x.startswith("0") else x)
        )

    if "Nomor KTP" in df_refi.columns:
        df_refi["Nomor KTP"] = df_refi["Nomor KTP"].astype(str).fillna("")

    # Remove duplicates from the Refi data if necessary
    df_refi = df_refi.drop_duplicates(
        subset=["Id Leads Data User", "Phone", "Nomor KTP"]
    )
    
    df_refi['LOB'] = "REFI"

    # Write the processed Refi DataFrame to the output Excel file
    df_refi.to_excel(writer, sheet_name="REFI", index=False)

    print(f"File has been generated at {output_directory} {output_file_path}")

    # Add blank sheets 'SUMMARY' and 'BOOKING'
    writer.book.add_worksheet("SUMMARY")
    writer.book.add_worksheet("BOOKING")
