import pandas as pd
from eppy.modeleditor import IDF
import os
import io
import re

def extract_idf_version(idf_file_object):
    """
    Extracts the EnergyPlus version from an IDF file.
    
    Args:
        idf_file_object: A file-like object for the IDF file
        
    Returns:
        str: Version string (e.g., '22.1.0') or None if not found
    """
    try:
        # Reset file pointer to beginning
        idf_file_object.seek(0)
        content = idf_file_object.read()
        
        # Look for Version object pattern: Version,22.1.0;
        version_pattern = r'Version\s*,\s*([\d\.]+)\s*;'
        match = re.search(version_pattern, content)
        
        if match:
            return match.group(1)
        else:
            return None
    except Exception as e:
        print(f"Error extracting IDF version: {e}")
        return None

def suggest_idd_file(idf_version, idd_files_dir):
    """
    Suggests the most appropriate IDD file based on IDF version.
    
    Args:
        idf_version (str): IDF version string (e.g., '22.1.0')
        idd_files_dir (str): Path to IDD_FILES directory
        
    Returns:
        str: Suggested IDD filename or None if no match found
    """
    if not idf_version:
        return None
        
    # Extract major.minor version (e.g., '22.1' from '22.1.0')
    version_parts = idf_version.split('.')
    if len(version_parts) >= 2:
        major_minor = f"{version_parts[0]}.{version_parts[1]}"
    else:
        major_minor = idf_version
    
    # Look for IDD files that match the version
    available_idd_files = []
    if os.path.exists(idd_files_dir):
        for file in os.listdir(idd_files_dir):
            if file.endswith('.idd'):
                # Extract version from filename (e.g., V22-1-0-Energy+.idd -> 22.1.0)
                idd_version_match = re.search(r'V(\d+)-(\d+)-(\d+)-Energy\+\.idd', file)
                if idd_version_match:
                    idd_version = f"{idd_version_match.group(1)}.{idd_version_match.group(2)}.{idd_version_match.group(3)}"
                    available_idd_files.append((file, idd_version))
    
    # Find exact match first
    for file, idd_version in available_idd_files:
        if idd_version == idf_version:
            return file
    
    # Find major.minor match
    for file, idd_version in available_idd_files:
        if idd_version.startswith(major_minor):
            return file
    
    # If no match, return the latest version
    if available_idd_files:
        # Sort by version (newest first)
        available_idd_files.sort(key=lambda x: tuple(map(int, x[1].split('.'))), reverse=True)
        return available_idd_files[0][0]
    
    return None

def extract_idf_data(idf_file_object, idd_path):
    """
    Reads an IDF file object and extracts all parameters.

    Args:
        idf_file_object: A file-like object for the IDF file (e.g., from open() or io.StringIO).
        idd_path (str): The full path to the IDD file.

    Returns:
        dict: A dictionary containing the extracted data, with object types as keys.
              Returns None if an error occurs.
    """
    try:
        IDF.setiddname(idd_path)
        idf = IDF(idf_file_object)
        
        all_sheets_data = {}
        
        for obj_type in idf.idfobjects.keys():
            objects = idf.idfobjects.get(obj_type)
            if objects:
                object_data = []
                for obj in objects:
                    obj_list = obj.obj
                    field_names = obj.fieldnames
                    
                    for i, field_value in enumerate(obj_list):
                        if i < len(field_names):
                            field_name = field_names[i]
                            object_data.append({
                                'ObjectName': obj.Name if hasattr(obj, 'Name') else 'N/A',
                                'FieldName': field_name,
                                'Value': str(field_value), # Ensure all values are strings
                                'Unit': ''  # Unit info needs to be parsed from IDD, placeholder
                            })
                
                if object_data:
                    all_sheets_data[obj_type] = object_data
        
        return all_sheets_data

    except Exception as e:
        print(f"An error occurred during IDF data extraction: {e}")
        return None

def write_data_to_excel(all_sheets_data, output_excel_path, include_individual_sheets=True):
    """
    Writes the extracted IDF data to a multi-sheet Excel file.

    Args:
        all_sheets_data (dict): The data dictionary from extract_idf_data.
        output_excel_path (str): The path to save the output Excel file.
        include_individual_sheets (bool): Whether to include individual object type sheets.
    """
    if not all_sheets_data:
        print("No data provided to write to Excel.")
        return

    try:
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            # 1. Create "ALL" worksheet
            all_data = []
            for obj_type, data in all_sheets_data.items():
                for record in data:
                    all_data.append({'ObjectType': obj_type, **record})
            
            if all_data:
                all_df = pd.DataFrame(all_data)
                all_df.to_excel(writer, index=False, sheet_name='ALL')
                print(f"  - ALL: {len(all_data)} records")

            # 2. Create separate worksheets for each object type (if requested)
            if include_individual_sheets:
                for obj_type, data in all_sheets_data.items():
                    df = pd.DataFrame(data)
                    # Sanitize worksheet name
                    sheet_name = ''.join(c for c in obj_type if c.isalnum() or c in (' ', '_')).rstrip()[:31]
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
                    print(f"  - {obj_type}: {len(data)} records")
        
        total_sheets = 1 + (len(all_sheets_data) if include_individual_sheets else 0)
        print(f"✅ Success! Data saved to: {os.path.abspath(output_excel_path)}")
        print(f"   Total {total_sheets} worksheets (including ALL).")

    except Exception as e:
        print(f"An error occurred while writing to Excel: {e}")


# --- Main execution block for standalone script usage ---
if __name__ == '__main__':
    # --- 1. Configuration for standalone run ---
    # Note: These paths are relative to the project root where the script is executed from.
    IDF_FILE_PATH = "2025_DOE_Prototype.idf" 
    # This assumes the IDD files are in a known location relative to the project root.
    # You might need to adjust this path.
    IDD_FILE_PATH = "IDD_FILES/V25-1-0-Energy+.idd" 
    OUTPUT_EXCEL_PATH = "idf_parameters_output_standalone.xlsx"

    print("--- Running IDF Processor as a standalone script ---")

    # --- 2. Check for files ---
    if not os.path.exists(IDF_FILE_PATH):
        print(f"Error: IDF file not found. Please check the path: {IDF_FILE_PATH}")
    elif not os.path.exists(IDD_FILE_PATH):
        print(f"Error: IDD file not found. Please check the path: {IDD_FILE_PATH}")
    else:
        try:
            # --- 3. Process the file ---
            print(f"Reading IDF file: {IDF_FILE_PATH}...")
            with open(IDF_FILE_PATH, 'r', encoding='utf-8') as idf_file:
                # The file needs to be passed as a file-like object
                idf_file_object = io.StringIO(idf_file.read())
                
            print("Extracting data...")
            extracted_data = extract_idf_data(idf_file_object, IDD_FILE_PATH)

            # --- 4. Write to Excel ---
            if extracted_data:
                print("\nWriting data to Excel...")
                write_data_to_excel(extracted_data, OUTPUT_EXCEL_PATH)
            else:
                print("No data was extracted, Excel file not generated.")

        except Exception as e:
            print(f"\nAn unexpected error occurred in the main block: {e}")
            print("Please ensure your EnergyPlus installation is correct and the IDF file is valid.")