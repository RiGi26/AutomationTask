from pathlib import Path
import time
import os
from google.cloud import bigquery
import pandas as pd

# Set the Google Cloud authentication environment variable
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = (
    "C:\\Users\\61140\\.vscode\\Website\\Automate MOXA\\DE Project\\data-engineer-project-443203-769b801de721.json"
)

path = Path("D:\\Cross Selling\\Moxa\\Booking\\2024\\November\\MOXA 20241130.xlsx")
sheets = ['NMC', 'NMC SY', 'REFI', 'AMITRA']

# Function to create a table reference
def table_reference(project_id, dataset_id, table_id):
    dataset_ref = bigquery.DatasetReference(project_id, dataset_id)
    table_ref = bigquery.TableReference(dataset_ref, table_id)
    return table_ref

# Function to delete all tables in a dataset
def delete_dataset_tables(client, project_id, dataset_id):
    tables = client.list_tables(f'{project_id}.{dataset_id}')
    for table in tables:
        client.delete_table(table)
    print('Tables deleted.')

# Function to upload a CSV file to a BigQuery table
def upload_csv(client, table_ref, csv_file):
    # Delete the table if it exists
    client.delete_table(table_ref, not_found_ok=True)

    # Configure the load job
    load_job_configuration = bigquery.LoadJobConfig()
    load_job_configuration.autodetect = True
    load_job_configuration.source_format = bigquery.SourceFormat.CSV
    load_job_configuration.skip_leading_rows = 1
    load_job_configuration.allow_quoted_newlines = True

    # Load the CSV into BigQuery
    with open(csv_file, 'rb') as source_file:
        upload_job = client.load_table_from_file(
            source_file,
            destination=table_ref,
            job_config=load_job_configuration
        )

    while upload_job.state != 'DONE':
        time.sleep(2)
        upload_job.reload()
        print(f"Job state: {upload_job.state}")
    
    print("Upload completed.")
    print(upload_job.result())
    print()

# Main execution logic
project_id = 'data-engineer-project-443203'
dataset_id = 'Leads'

# Initialize the BigQuery client
client = bigquery.Client()
data_file_folder = Path("C:\\Users\\61140\\.vscode\\Website\\Automate MOXA")

try:
    for sheet in sheets:
        df = pd.read_excel(path, sheet_name=sheet)
        df_filter = df.rename(columns=lambda x:x.replace(' ', '_'))
        df_filter.to_csv(f'{sheet} MOXA Leads.csv', index=False)
        print(f"file {sheet} success to saved")
        print("")
    for file in os.listdir(data_file_folder):
        if file.endswith('.csv'):
            print(f'Processing file: {file}')
            table_name = '_'.join(file.split()[:-1]) 
            csv_file = data_file_folder / file
            table_ref = table_reference(project_id, dataset_id, table_name)
            upload_csv(client, table_ref, csv_file)  # Upload the CSV
            # delete_dataset_tables(client, project_id, dataset_id)
            
except Exception as e:
    print(f"Error occurred: {e}")
