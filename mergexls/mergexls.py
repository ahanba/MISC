import os
import pandas as pd
import zipfile
import shutil
from pathlib import Path

def extract_zip(zip_path, extract_path):
    """
    Extract zip file to specified path.
    Returns True if successful, False otherwise.
    """
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)
        return True
    except Exception as e:
        print(f"Error extracting zip file: {e}")
        return False

def merge_excel_files(folder_path):
    """
    Merge multiple Excel files, maintaining consistent first two columns 
    and adding unique column values from each file.
    """
    # List to store DataFrames from each Excel file
    dfs = []
    
    # Iterate through Excel files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith('.xls') or filename.endswith('.xlsx'):
            # Full path to the file
            file_path = os.path.join(folder_path, filename)
            
            # Read the Excel file
            try:
                # Read the entire file
                df = pd.read_excel(file_path)
                
                # Check if file has at least 3 columns
                if len(df.columns) < 3:
                    print(f"Skipping {filename}: Not enough columns")
                    continue
                
                # Create a DataFrame with first two columns unchanged
                processed_df = df.iloc[:, :2].copy()
                
                # Add the third column with filename prefix
                base_filename = os.path.splitext(filename)[0]
                processed_df[base_filename] = df.iloc[:, 2]
                
                dfs.append(processed_df)
                
            except Exception as e:
                print(f"Error processing file {filename}: {e}")
    
    # Merge all DataFrames 
    if dfs:
        # Start with the first DataFrame
        merged_df = dfs[0]
        
        # Merge subsequent DataFrames
        for df in dfs[1:]:
            # Merge on first two columns
            merged_df = pd.merge(merged_df, df, on=df.columns[:2].tolist(), how='outer')
        
        return merged_df
    else:
        print("No Excel files found in the specified folder.")
        return None

def main():
    # Get the script's directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Look for zip files in the script's directory
    zip_files = [f for f in os.listdir(script_dir) if f.endswith('.zip')]
    
    if not zip_files:
        print("No zip files found in the script's directory.")
        return
    
    if len(zip_files) > 1:
        print("Multiple zip files found. Using the first one:", zip_files[0])
    
    zip_path = os.path.join(script_dir, zip_files[0])
    
    # Create a temporary directory for extracted files
    temp_dir = os.path.join(script_dir, 'temp_excel_files')
    os.makedirs(temp_dir, exist_ok=True)
    
    try:
        # Extract zip file
        print(f"Extracting {zip_files[0]}...")
        if not extract_zip(zip_path, temp_dir):
            print("Failed to extract zip file. Exiting.")
            return
        
        # Merge Excel files
        print("Merging Excel files...")
        merged_result = merge_excel_files(temp_dir)
        
        # Save the merged result
        if merged_result is not None:
            output_path = os.path.join(script_dir, 'merged_output.xlsx')
            merged_result.to_excel(output_path, index=False)
            print(f"Merged file saved to {output_path}")
            
            # Print column names for verification
            print("\nColumns in merged file:")
            print(merged_result.columns.tolist())
    
    finally:
        # Clean up: remove temporary directory and its contents
        print("Cleaning up temporary files...")
        try:
            shutil.rmtree(temp_dir)
        except Exception as e:
            print(f"Error cleaning up temporary files: {e}")

if __name__ == "__main__":
    main()