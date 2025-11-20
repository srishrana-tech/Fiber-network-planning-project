import pandas as pd
import os
from pathlib import Path

def merge_csvs_to_excel(folder_path, output_file='merged_output.xlsx'):
    """
    Merge all CSV files in a folder into a single Excel sheet.
    
    Args:
        folder_path: Path to folder containing CSV files
        output_file: Name of output Excel file (default: merged_output.xlsx)
    """
    # Convert to Path object
    folder = Path(folder_path)
    
    # Check if folder exists
    if not folder.exists():
        print(f"Error: Folder '{folder_path}' does not exist!")
        return
    
    # Get all CSV files in the folder and subfolders (recursive search)
    csv_files = list(folder.rglob('*.csv'))
    
    if not csv_files:
        print(f"No CSV files found in '{folder_path}'")
        return
    
    print(f"Found {len(csv_files)} CSV file(s)")
    
    # List to store dataframes
    dfs = []
    
    # Read each CSV and append to list
    for csv_file in csv_files:
        print(f"Reading: {csv_file.relative_to(folder)}")
        try:
            df = pd.read_csv(csv_file)
            dfs.append(df)
        except Exception as e:
            print(f"Error reading {csv_file.name}: {e}")
    
    if not dfs:
        print("No data to merge!")
        return
    
    # Merge all dataframes
    merged_df = pd.concat(dfs, ignore_index=True)
    
    # Save to Excel
    output_path = folder / output_file
    merged_df.to_excel(output_path, index=False, engine='openpyxl')
    
    print(f"\nSuccess! Merged {len(dfs)} CSV files into '{output_path}'")
    print(f"Total rows: {len(merged_df)}")
    print(f"Total columns: {len(merged_df.columns)}")

if __name__ == "__main__":
    # Get folder path from user
    folder_path = input("Enter the folder path containing CSV files: ").strip()
    
    # Optional: Get custom output filename
    custom_output = input("Enter output filename (press Enter for 'merged_output.xlsx'): ").strip()
    output_file = custom_output if custom_output else 'merged_output.xlsx'
    
    # Ensure .xlsx extension
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'
    
    # Merge CSVs
    merge_csvs_to_excel(folder_path, output_file)