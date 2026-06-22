import os
import json
import pandas as pd
from pathlib import Path
import sys

# Configure UTF-8 stdout
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# Add parent directory to path so we can import config if run directly
sys.path.insert(0, str(Path(__file__).resolve().parent))
import config
from sharepoint import AuthProvider, SharePointClient
def initialize():
    target_file_name = Path(config.SHAREPOINT_TARGET_FILE_PATH).name
    excel_path = config.PATH_OUTPUT / target_file_name
    
    if not excel_path.exists():
        print(f"Excel file '{target_file_name}' not found locally. Checking SharePoint...")
        auth = AuthProvider()
        client = SharePointClient(auth)
        
        target_exists = client.check_file_exists(
            config.SHAREPOINT_TARGET_FILE_PATH, 
            drive_id=config.SHAREPOINT_TARGET_DRIVE_ID
        )
        if target_exists:
            print(f"Downloading target file from SharePoint...")
            excel_path = client.download_file(
                config.SHAREPOINT_TARGET_FILE_PATH, 
                excel_path, 
                drive_id=config.SHAREPOINT_TARGET_DRIVE_ID
            )
        else:
            print("Target file not found on SharePoint. Checking local CRM_merge.xlsx fallback...")
            local_merge = config.PATH_DATA / "CRM_merge.xlsx"
            if local_merge.exists():
                excel_path = local_merge
            else:
                print("Local CRM_merge.xlsx not found. Downloading source file from SharePoint...")
                excel_path = client.download_file(
                    config.SHAREPOINT_SOURCE_FILE_PATH, 
                    excel_path, 
                    drive_id=config.SHAREPOINT_SOURCE_DRIVE_ID
                )

    print(f"Reading Excel file from {excel_path}...")
    df = pd.read_excel(excel_path)
    
    print("Columns found in Excel:", list(df.columns))
    
    history_db = {}
    cols_to_extract = config.OUTPUT_COLUMNS
    
    sheet_cols = list(df.columns)
    mapping = {}
    for c in cols_to_extract:
        if c in sheet_cols:
            mapping[c] = c
        else:
            for sc in sheet_cols:
                if c.split("] ")[-1].lower() in sc.lower():
                    mapping[c] = sc
                    break

    print("Mapping config columns to sheet columns:")
    for k, v in mapping.items():
        print(f"  {k} -> {v}")

    count = 0
    for idx, row in df.iterrows():
        act_id = str(row.get("ActivityId", "")).strip()
        if not act_id or act_id.lower() in ("nan", "none", ""):
            continue
            
        has_data = False
        row_fills = {}
        for col_name, sheet_col in mapping.items():
            val = row.get(sheet_col)
            if pd.notna(val) and str(val).strip() != "" and str(val).strip().lower() != "nan":
                row_fills[col_name] = str(val).strip()
                has_data = True
            else:
                row_fills[col_name] = None
                
        if has_data:
            history_db[act_id] = row_fills
            count += 1

    output_path = config.DB_JSON_PATH
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(history_db, f, ensure_ascii=False, indent=2)
        
    print(f"Successfully initialized history DB at {output_path} with {count} items!")

if __name__ == "__main__":
    initialize()
