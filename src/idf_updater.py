"""
IDF File Updater - Updates IDF files based on modified Excel files

This module provides functionality to update IDF files based on modifications
made to Excel files that were originally generated from IDF files.
"""

import pandas as pd
from eppy.modeleditor import IDF
import os
import io
from typing import Dict, List, Tuple


def read_excel_data(excel_file_path: str) -> Dict[str, pd.DataFrame]:
    """
    Reads the modified Excel file and extracts data from all worksheets.
    
    Args:
        excel_file_path (str): Path to the modified Excel file
        
    Returns:
        Dict[str, pd.DataFrame]: Dictionary with sheet names as keys and DataFrames as values
        
    Raises:
        FileNotFoundError: If Excel file doesn't exist
        ValueError: If Excel file format is invalid
    """
    if not os.path.exists(excel_file_path):
        raise FileNotFoundError(f"Excel file not found: {excel_file_path}")
    
    try:
        # Read all sheets from Excel
        excel_data = pd.read_excel(excel_file_path, sheet_name=None, engine='openpyxl')
        
        # Validate required columns in each sheet
        for sheet_name, df in excel_data.items():
            if sheet_name != 'ALL':
                required_columns = ['ObjectName', 'FieldName', 'Value']
                missing_columns = [col for col in required_columns if col not in df.columns]
                if missing_columns:
                    raise ValueError(f"Sheet '{sheet_name}' missing required columns: {missing_columns}")
        
        return excel_data
        
    except Exception as e:
        raise ValueError(f"Error reading Excel file: {str(e)}")


def update_idf_from_excel(original_idf_path: str, modified_excel_path: str, idd_path: str, 
                         output_idf_path: str) -> Dict[str, any]:
    """
    Updates an IDF file based on modifications in an Excel file.
    
    Args:
        original_idf_path (str): Path to the original IDF file
        modified_excel_path (str): Path to the modified Excel file
        idd_path (str): Path to the IDD file
        output_idf_path (str): Path to save the updated IDF file
        
    Returns:
        Dict[str, any]: Update statistics and logs
        
    Raises:
        FileNotFoundError: If input files don't exist
        ValueError: If file formats are invalid
    """
    # Input validation
    if not os.path.exists(original_idf_path):
        raise FileNotFoundError(f"Original IDF file not found: {original_idf_path}")
    if not os.path.exists(modified_excel_path):
        raise FileNotFoundError(f"Modified Excel file not found: {modified_excel_path}")
    if not os.path.exists(idd_path):
        raise FileNotFoundError(f"IDD file not found: {idd_path}")
    
    # Initialize statistics
    stats = {
        'total_updates': 0,
        'successful_updates': 0,
        'failed_updates': 0,
        'warnings': [],
        'errors': [],
        'updated_objects': []
    }
    
    try:
        # Load IDD and IDF
        IDF.setiddname(idd_path)
        
        with open(original_idf_path, 'r', encoding='utf-8') as f:
            idf_content = f.read()
        
        idf_file_object = io.StringIO(idf_content)
        idf = IDF(idf_file_object)
        
        # Read Excel data
        excel_data = read_excel_data(modified_excel_path)
        
        # Process each sheet (skip 'ALL' sheet)
        for sheet_name, df in excel_data.items():
            if sheet_name == 'ALL':
                continue
                
            # Get object type from sheet name
            object_type = sheet_name.strip()
            
            # Get objects of this type from IDF
            if object_type in idf.idfobjects:
                idf_objects = idf.idfobjects[object_type]
                
                # Group Excel data by object name
                object_groups = df.groupby('ObjectName')
                
                for object_name, group_df in object_groups:
                    # Find matching object in IDF
                    matching_object = None
                    for obj in idf_objects:
                        if hasattr(obj, 'Name') and obj.Name == object_name:
                            matching_object = obj
                            break
                    
                    if matching_object:
                        # Update object fields
                        for _, row in group_df.iterrows():
                            field_name = row['FieldName']
                            new_value = str(row['Value'])
                            
                            try:
                                # Find field index
                                field_names = matching_object.fieldnames
                                if field_name in field_names:
                                    field_index = field_names.index(field_name)
                                    
                                    # Update the field value
                                    old_value = str(matching_object.obj[field_index])
                                    
                                    if old_value != new_value:
                                        matching_object.obj[field_index] = new_value
                                        stats['total_updates'] += 1
                                        stats['successful_updates'] += 1
                                        
                                        # Record update
                                        update_record = {
                                            'object_type': object_type,
                                            'object_name': object_name,
                                            'field_name': field_name,
                                            'old_value': old_value,
                                            'new_value': new_value
                                        }
                                        stats['updated_objects'].append(update_record)
                                        
                                else:
                                    stats['warnings'].append(
                                        f"Field '{field_name}' not found in object '{object_name}' of type '{object_type}'"
                                    )
                                    stats['failed_updates'] += 1
                                    
                            except Exception as e:
                                stats['errors'].append(
                                    f"Error updating field '{field_name}' in object '{object_name}': {str(e)}"
                                )
                                stats['failed_updates'] += 1
                    else:
                        stats['warnings'].append(
                            f"Object '{object_name}' of type '{object_type}' not found in IDF"
                        )
                        stats['failed_updates'] += 1
        
        # Save updated IDF
        idf.saveas(output_idf_path)
        
        return stats
        
    except Exception as e:
        stats['errors'].append(f"Critical error during IDF update: {str(e)}")
        return stats


def validate_excel_for_update(excel_file_path: str) -> Tuple[bool, List[str]]:
    """
    Validates if an Excel file is suitable for IDF updates.
    
    Args:
        excel_file_path (str): Path to the Excel file
        
    Returns:
        Tuple[bool, List[str]]: (is_valid, validation_messages)
    """
    messages = []
    
    if not os.path.exists(excel_file_path):
        return False, [f"File not found: {excel_file_path}"]
    
    try:
        excel_data = pd.read_excel(excel_file_path, sheet_name=None, engine='openpyxl')
        
        # Check if there are sheets other than 'ALL'
        object_sheets = [name for name in excel_data.keys() if name != 'ALL']
        
        if not object_sheets:
            messages.append("No object-specific sheets found (only 'ALL' sheet present)")
            return False, messages
        
        # Validate each object sheet
        for sheet_name in object_sheets:
            df = excel_data[sheet_name]
            required_columns = ['ObjectName', 'FieldName', 'Value']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                messages.append(f"Sheet '{sheet_name}' missing columns: {missing_columns}")
            
            # Check for empty data
            if df.empty:
                messages.append(f"Sheet '{sheet_name}' is empty")
        
        if messages:
            return False, messages
        else:
            messages.append("Excel file format is valid for IDF updates")
            return True, messages
            
    except Exception as e:
        return False, [f"Error reading Excel file: {str(e)}"]


# Example usage and testing
if __name__ == "__main__":
    # Example configuration
    ORIGINAL_IDF = "ASHRAE901_Hospital_STD2022/ASHRAE901_Hospital_STD2022_Albuquerque.idf"
    MODIFIED_EXCEL = "modified_parameters.xlsx"
    IDD_FILE = "IDD_FILES/V25-1-0-Energy+.idd"
    OUTPUT_IDF = "updated_model.idf"
    
    print("--- Testing IDF Updater ---")
    
    # Validate Excel file first
    is_valid, messages = validate_excel_for_update(MODIFIED_EXCEL)
    print("Excel validation:")
    for msg in messages:
        print(f"  - {msg}")
    
    if is_valid:
        # Perform update
        stats = update_idf_from_excel(ORIGINAL_IDF, MODIFIED_EXCEL, IDD_FILE, OUTPUT_IDF)
        
        print(f"\nUpdate Statistics:")
        print(f"  - Total updates attempted: {stats['total_updates']}")
        print(f"  - Successful updates: {stats['successful_updates']}")
        print(f"  - Failed updates: {stats['failed_updates']}")
        
        if stats['warnings']:
            print(f"\nWarnings ({len(stats['warnings'])}):")
            for warning in stats['warnings'][:5]:  # Show first 5 warnings
                print(f"  - {warning}")
        
        if stats['errors']:
            print(f"\nErrors ({len(stats['errors'])}):")
            for error in stats['errors'][:5]:  # Show first 5 errors
                print(f"  - {error}")
        
        if stats['updated_objects']:
            print(f"\nSample updates ({min(3, len(stats['updated_objects']))}):")
            for update in stats['updated_objects'][:3]:
                print(f"  - {update['object_type']}.{update['object_name']}.{update['field_name']}: "
                      f"{update['old_value']} → {update['new_value']}")
    else:
        print("\nExcel file is not valid for IDF updates.")