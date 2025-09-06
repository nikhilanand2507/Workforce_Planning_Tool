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
import re

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

def fetch_and_clean_checker_from_sharepoint(sheet):
    """
    Fetches a macro Excel file from SharePoint using fetch_file_from_sharepoint_folder,
    loads all sheets into DataFrames, cleans the 'Checker' sheet by dropping rows with blank 'Person Number (from Department Tab)',
    prints df.head() for each sheet, and returns the cleaned DataFrame.
    """
    
    sheets = {name: pd.read_excel(sheet, sheet_name=name, header=0 if name.lower() == "dropdown" else 1) for name in sheet.sheet_names}
    for name, df in sheets.items():
        print(f"\nSheet: {name}")
        print(df.head(15))
    # Clean 'Checker' sheet
    if "Checker" in sheets:
        df_checker = sheets["Checker"]
        df_checker_no_blank = df_checker.dropna(subset=["Person Number\n(from Department Tab)"])
        print("\nOriginal 'Checker' sheet (may contain blanks):")
        print(df_checker.head(15))
        print("\n'Checker' sheet with no blank 'Person Number':")
        print(df_checker_no_blank.head(15))
        return sheets, df_checker_no_blank
    else:
        print("'Checker' sheet not found in macro file.")
        return sheets, None

def fetch_latest_pfp_for_employee_remapping_to_create_gba_from_old_pfp(site_url, folder_path, client_id, client_secret):
        """
        Fetches the latest 'Project Plan Analysis-continuous-YYYY-MM-DD.xlsx' file from the specified SharePoint folder.
        Returns the DataFrame of the latest file, or None if not found.
        """
        credentials = ClientCredential(client_id, client_secret)
        context = ClientContext(site_url).with_credentials(credentials)
        folder = context.web.get_folder_by_server_relative_url(folder_path)
        files = folder.files.get().execute_query()
        pattern = r"Project Plan Analysis-continuous-(\d{4}-\d{2}-\d{2})\.xlsx"
        excel_files = []
        file_dates = []

        for f in files:
            fname = f.properties["Name"]
            match = re.match(pattern, fname)
            if match:
                excel_files.append(f)
                file_dates.append(datetime.strptime(match.group(1), "%Y-%m-%d"))

        if not excel_files:
            print("No matching Excel files found in folder.")
            return None
        else:
            latest_idx = file_dates.index(max(file_dates))
            latest_file = excel_files[latest_idx]
            latest_fname = latest_file.properties["Name"]
            print(f"Latest file: {latest_fname}")
            file_content = File.open_binary(context, latest_file.serverRelativeUrl)
            df = pd.read_excel(io.BytesIO(file_content.content), engine="openpyxl")
            print(f"First 5 rows of '{latest_fname}':")
            print(df.head(5))
            print("-" * 40)
            return df
        
def process_pfp_and_workbook_structure_checker_tab_and_merge_for_first_run(pfp_df, gba_df):
    # Merge based on resource name/person
    merged_df = pd.merge(
        pfp_df,
        gba_df[['Person Number\n(from Department Tab)', 'File Name','Department Name', 'Department Manager','Name']],
        left_on='Resource',
        right_on='Name',
        how='left'
    )
    # Select only the required columns
    merged_df = merged_df[
        [
            'Unique Code',
            'Project Number',
            'Project Name',
            'Resource',
            'Expenditure Organization Name',
            'Person Number\n(from Department Tab)',
            'File Name',
            'Department Manager'
        ]
    ]
    # Extract unique values before ':' in 'Expenditure Organization Name' column
    if 'Expenditure Organization Name' in merged_df.columns:
        unique_prefixes = merged_df['Expenditure Organization Name'].dropna().apply(lambda x: str(x).split(':')[0].strip()).unique()
    else:
        unique_prefixes = []
    # Extract unique suffixes
    unique_suffixes = set()
    for val in unique_prefixes:
        parts = val.rsplit(' ', 1)
        if len(parts) > 1:
            unique_suffixes.add(parts[-1])
        else:
            unique_suffixes.add(parts[0])
    print("Unique suffixes:", unique_suffixes)
    # Define mapping for suffix categories
    suffix_map = {
        "MOB": ["MOB", "MOBILITY", "Mobility"],
        "PLA": ["PLA", "PLACES", "Places"],
        "RES": ["RES", "RESILIENCE"],
        "EF": ["EF", "Enabling Function", "ENABLING FUNCTION"],
        "SSC": ["SSC", "SHARED SERVICES", "Shared Services"]
    }
    
    def normalize_suffix(s):
        """This code takes a set of suffixes extracted from organization names and checks each one against a predefined mapping of categories (like MOB, PLA, etc.).
        It normalizes each suffix (removes spaces, makes uppercase) and compares it to all possible values in each category.
        If a match is found, the suffix is added to a set of selected suffixes. 
        This helps you identify which suffixes in your data belong to specific business categories."""
        return str(s).strip().upper()
    selected_suffixes = set()
    for suffix in unique_suffixes:
        norm = normalize_suffix(suffix)
        for key, values in suffix_map.items():
            if any(norm == v.upper() for v in values):
                selected_suffixes.add(suffix)
                break
    # Create separate DataFrames for each suffix category
    def get_suffix(org_name):
        val = str(org_name).split(':')[0].strip()
        parts = val.rsplit(' ', 1)
        return parts[-1] if len(parts) > 1 else parts[0]
    filtered_dfs = {}
    for key, values in suffix_map.items():
        mask = merged_df['Expenditure Organization Name'].dropna().apply(
            lambda x: any(normalize_suffix(get_suffix(x)) == v.upper() for v in values)
        )
        filtered_df = merged_df[mask]
        filtered_dfs[key] = filtered_df
    return merged_df, filtered_dfs
    


def run_streamlit_app():
    st.set_page_config(page_title="Work Force Planning", layout="centered")
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

    st.info("Click the button below for GBA-wise data extraction and upload.")

    if st.button("GBA Wise Data Extraction"):
        with st.spinner("Extracting and uploading GBA-wise data. Please wait..."):
            # Fetch macro workbook and process sheets
            site_url = "https://arcadiso365.sharepoint.com/teams/CPW_Testing"
            folder_path = "/teams/CPW_Testing/Shared Documents/PLA CAN CPW Tool/CPW FINAL PACKAGE/01 Data Processing/Workbook Structure"
            client_id = os.getenv('CLIENT_ID')
            client_secret = os.getenv('CLIENT_SECRET')
            file_name, df = fetch_file_from_sharepoint_folder(site_url, folder_path, client_id, client_secret)
            if file_name and file_name.endswith(".xlsm"):
                file_content = File.open_binary(ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret)), folder_path + "/" + file_name).content
                xls = pd.ExcelFile(io.BytesIO(file_content), engine="openpyxl")
            else:
                st.error("No macro workbook (.xlsm) found for GBA extraction.")
                return

            sheets, df_checker_cleaned = fetch_and_clean_checker_from_sharepoint(xls)
            if df_checker_cleaned is None:
                st.error("Checker sheet not found or no valid rows.")
                return

            # Fetch latest cleaned PFP file
            old_pfp_folder = "/teams/CPW_Testing/Shared Documents/PLA CAN CPW Tool/CPW FINAL PACKAGE/01 Data Processing/Project Financial Plan (PFP)/OLD PFP"
            latest_pfp_df = fetch_latest_pfp_for_employee_remapping_to_create_gba_from_old_pfp(site_url, old_pfp_folder, client_id, client_secret)
            if latest_pfp_df is None:
                st.error("No latest PFP file found for GBA extraction.")
                return

            merged_df, filtered_dfs = process_pfp_and_workbook_structure_checker_tab_and_merge_for_first_run(latest_pfp_df, df_checker_cleaned)

            # Add 'Oracle Date' and 'Index' columns to each filtered DataFrame
            oracle_date = datetime.now().strftime('%Y-%m-%d')
            for key, df in filtered_dfs.items():
                df.insert(0, 'Oracle Date', oracle_date)
                df.insert(1, 'Index', range(1, len(df) + 1))
                filtered_dfs[key] = df

            # Save each filtered DataFrame to SharePoint GBA Workbooks folder
            gba_folder = "/teams/CPW_Testing/Shared Documents/PLA CAN CPW Tool/CPW FINAL PACKAGE/02 GBA Workbooks"
            success_count = 0
            for key, df in filtered_dfs.items():
                output_file_name = f"CPW_Tool_{key}_Main.xlsx"
                success = upload_dataframe_to_sharepoint_folder(site_url, gba_folder, output_file_name, df, client_id, client_secret)
                if success:
                    success_count += 1
            if success_count:
                st.success(f"{success_count} GBA-wise DataFrames uploaded to SharePoint GBA Workbooks folder.")
            else:
                st.error("Failed to upload GBA-wise DataFrames.")

if __name__ == "__main__":
    run_streamlit_app()