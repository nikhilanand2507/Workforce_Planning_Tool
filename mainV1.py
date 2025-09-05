import streamlit as st
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.lists.list import List
import io
from office365.sharepoint.files.file import File
import pandas as pd
import os
from datetime import datetime

# Load environment variables
load_dotenv()

    
def list_folders_and_subfolders(context, folder_path):
    """
    List all folders and subfolders in the given SharePoint folder.
    Returns a list of tuples: (folder_name, folder_server_relative_url)
    """
    folders_info = []
    try:
        folder = context.web.get_folder_by_server_relative_url(folder_path)
        folders = folder.folders.get().execute_query()
        for subfolder in folders:
            folders_info.append((subfolder.name, subfolder.serverRelativeUrl))
            # Recursively list subfolders
            subfolders = list_folders_and_subfolders(context, subfolder.serverRelativeUrl)
            folders_info.extend(subfolders)
    except Exception as e:
        print(f"Error listing folders: {e}")
    return folders_info


def fetch_file_from_sharepoint_folder(site_url, folder_path, client_id, client_secret):
    """
     Parameters:
        site_url (str): The URL of the SharePoint site.
        folder_path (str): The path to the SharePoint folder.
        file_name (str): The name of the file to fetch.
        client_id (str): The Client ID for authentication.
        client_secret (str): The Client Secret for authentication.
    """
    try:
        credentials = ClientCredential(client_id, client_secret)
        context = ClientContext(site_url).with_credentials(credentials)
        # List folders and subfolders
        folders_info = list_folders_and_subfolders(context, folder_path)
        print("Folders and subfolders in the folder:")
        for name, url in folders_info:
            print(f"Name: {name}, Link: {url}")
        # Access the SharePoint folder
        folder = context.web.get_folder_by_server_relative_url(folder_path)
        files = folder.files.get().execute_query()
        if not files:
            print(f"No files found in folder: {folder_path}")
            return None, pd.DataFrame()
       
        # pick first Excel file (.xlsx, .xlsm, .xls)
        for f in files:
            fname = f.properties["Name"]
            if fname.endswith((".xlsx", ".xlsm", ".xls")):
                file_content = File.open_binary(context, f.serverRelativeUrl)

                # choose engine depending on extension
                if fname.endswith((".xlsx", ".xlsm")):
                    df = pd.read_excel(io.BytesIO(file_content.content), engine="openpyxl")
                else:  # .xls
                    df = pd.read_excel(io.BytesIO(file_content.content), engine="xlrd")

                return fname, df

        print("No Excel files in folder")
        return None, pd.DataFrame()

    except Exception as e:
        print(f"Error accessing SharePoint folder or files: {e}")
        return None, pd.DataFrame()
    
def first_time_run_pfp(df):
    print("First 5 rows:")
    print(df.head(5))
    print("\nColumn names:")
    print(df.columns)
    df['Unique Code'] = df['Project Number'].astype(str) + ' - ' + df['Employee Name'].astype(str)
    cols = ['Unique Code'] + [col for col in df.columns if col != 'Unique Code']
    df = df[cols]
    print("\nNew column 'Unique Code':")
    print(df[['Unique Code']].head(5))
    print(df.head(5))
    ## Save and upload the updated DataFrame to SharePoint
    #Unique_code_date_str = datetime.now().strftime('%Y-%m-%d')
    #output_file_name = f"Project Plan Analysis-continuous-Unique code-{Unique_code_date_str}.xlsx"
    #upload_dataframe_to_sharepoint_folder(site_url, folder_path, output_file_name, df, client_id, client_secret)
	
    # Remove duplicate rows based on 'Unique Code', keeping only the first occurrence
    df_unique = df.drop_duplicates(subset=['Unique Code'], keep='first')
	
    ## Save and upload the deduplicated DataFrame to OLD PFP folder in SharePoint
    #date_str = datetime.now().strftime('%Y-%m-%d')
    #dedup_output_file_name = f"Project Plan Analysis-continuous-unique values-{date_str}.xlsx"
    #upload_dataframe_to_sharepoint_folder(site_url, old_pfp_folder, dedup_output_file_name, df_unique, client_id, client_secret)
	
    # Remove rows with missing values in 'Employee Name' column
    df_no_missing = df_unique.dropna(subset=['Employee Name'])
    # Remove rows where 'Employee Name' is 'Labor Cost, Conversion Employee'
    df_final = df_no_missing[df_no_missing['Employee Name'] != 'Labor Cost, Conversion Employee']
    return df_final

    ## Save and upload the final cleaned DataFrame to OLD PFP folder in SharePoint
    #final_date_str = datetime.now().strftime('%Y-%m-%d')
    #final_output_file_name = f"Project Plan Analysis-continuous-final-{final_date_str}.xlsx"
    #upload_dataframe_to_sharepoint_folder(site_url, old_pfp_folder, final_output_file_name, df_final, client_id, client_secret)
    #print("\nAll files processed and uploaded to SharePoint.")

def upload_dataframe_to_sharepoint_folder(site_url, folder_path, file_name, df, client_id, client_secret):
	try:
		credentials = ClientCredential(client_id, client_secret)
		context = ClientContext(site_url).with_credentials(credentials)
		folder = context.web.get_folder_by_server_relative_url(folder_path)
		excel_buffer = io.BytesIO()
		with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
			df.to_excel(writer, index=False)
		excel_buffer.seek(0)
		folder.upload_file(file_name, excel_buffer.read()).execute_query()
		print(f"File '{file_name}' uploaded successfully to '{folder_path}'")
		return True
	except Exception as e:
		print(f"Error uploading DataFrame: {e}")
		return False


def run_streamlit_app():
    st.set_page_config(page_title="PFP Data Processor", layout="centered")
    st.title("Project Financial Plan (PFP) Data Processor")
    st.markdown("""
        <style>
        .stButton>button {
            background-color: #0078D4;
            color: white;
            font-size: 18px;
            padding: 0.5em 2em;
            border-radius: 8px;
        }
        </style>
    """, unsafe_allow_html=True)

    st.info("Click the button below to fetch, clean, and upload the PFP data to SharePoint.")

    if st.button("Process PFP Data for First Time Run"):
        with st.spinner("Processing and uploading data. Please wait..."):
            site_url = "https://arcadiso365.sharepoint.com/teams/CPW_Testing"
            folder_path = "/teams/CPW_Testing/Shared Documents/PLA CAN CPW Tool/CPW FINAL PACKAGE/01 Data Processing/Project Financial Plan (PFP)"
            old_pfp_folder = folder_path + "/OLD PFP"
            client_id = os.getenv('CLIENT_ID')
            client_secret = os.getenv('CLIENT_SECRET')
            file_name, df = fetch_file_from_sharepoint_folder(site_url, folder_path, client_id, client_secret)
            if file_name and not df.empty:
                cleaned_df = first_time_run_pfp(df)
                final_date_str = datetime.now().strftime('%Y-%m-%d')
                output_file_name = f"Project Plan Analysis-continuous-{final_date_str}.xlsx"
                success = upload_dataframe_to_sharepoint_folder(site_url, old_pfp_folder, output_file_name, cleaned_df, client_id, client_secret)
                if success:
                    st.success(f"Cleaned PFP data uploaded to SharePoint folder: {old_pfp_folder}")
                else:
                    st.error("Error uploading cleaned data to SharePoint.")
            else:
                st.error("No file fetched or DataFrame is empty.")

if __name__ == "__main__":
    run_streamlit_app()