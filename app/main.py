import streamlit as st
import sys
import os
import io
import pandas as pd

# Add the src directory to Python path to import idf_processor
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from idf_processor import extract_idf_data, write_data_to_excel, extract_idf_version, suggest_idd_file
from idf_updater import update_idf_from_excel, validate_excel_for_update

# Configure the page
st.set_page_config(
    page_title="IDF to Excel Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Header
    st.markdown('<div class="main-header">🏢 IDF to Excel Converter & Updater</div>', unsafe_allow_html=True)
    
    st.markdown("""
    This application provides two main functions:
    - **Convert IDF to Excel**: ✅ Extract all parameters from IDF files and organize them into Excel worksheets
    - **Update IDF from Excel**: 🔧 **Under Development** - Apply modifications from Excel files back to IDF files
    """)
    
    # Sidebar for file uploads
    with st.sidebar:
        st.header("📁 File Upload")
        
        # Operation selection
        operation = st.radio(
            "Select Operation:",
            ["Convert IDF to Excel", "Update IDF from Excel (Under Development 🔧)"],
            help="Choose whether to convert IDF to Excel or update IDF from Excel"
        )
        
        # IDF file upload
        uploaded_idf = st.file_uploader(
            "Upload IDF File", 
            type=['idf'],
            help="Upload your EnergyPlus IDF file"
        )
        
        # Excel file upload (for update operation)
        if operation == "Update IDF from Excel (Under Development 🔧)":
            uploaded_excel = st.file_uploader(
                "Upload Modified Excel File",
                type=['xlsx'],
                help="Upload the Excel file with your modifications"
            )
        
        # IDD file selection
        st.subheader("IDD File Selection")
        
        # Check if IDD_FILES directory exists
        idd_files_dir = os.path.join(os.path.dirname(__file__), '..', 'IDD_FILES')
        available_idd_files = []
        
        if os.path.exists(idd_files_dir):
            for file in os.listdir(idd_files_dir):
                if file.endswith('.idd'):
                    available_idd_files.append(file)
        
        # Detect IDF version and suggest IDD file
        suggested_idd = None
        if uploaded_idf:
            try:
                # Extract version from uploaded IDF
                idf_content = uploaded_idf.read().decode('utf-8')
                idf_file_object = io.StringIO(idf_content)
                idf_version = extract_idf_version(idf_file_object)
                
                if idf_version:
                    st.success(f"📋 Detected IDF Version: {idf_version}")
                    suggested_idd = suggest_idd_file(idf_version, idd_files_dir)
                    if suggested_idd:
                        st.info(f"💡 Suggested IDD file: {suggested_idd}")
                    else:
                        st.warning("⚠️ No matching IDD file found for detected version")
                else:
                    st.warning("⚠️ Could not detect IDF version from file")
                
                # Reset file pointer for later use
                uploaded_idf.seek(0)
                    
            except Exception as e:
                st.warning(f"⚠️ Could not detect IDF version: {str(e)}")
        
        if available_idd_files:
            # Set default to suggested IDD if available
            default_index = 0
            if suggested_idd and suggested_idd in available_idd_files:
                default_index = available_idd_files.index(suggested_idd)
            
            selected_idd = st.selectbox(
                "Select IDD File",
                available_idd_files,
                index=default_index,
                help="Choose the appropriate IDD file for your IDF version"
            )
            idd_file_path = os.path.join(idd_files_dir, selected_idd)
        else:
            st.warning("No IDD files found in IDD_FILES directory.")
            st.info("Please upload an IDD file:")
            uploaded_idd = st.file_uploader(
                "Upload IDD File", 
                type=['idd'],
                help="Upload your EnergyPlus IDD file"
            )
            if uploaded_idd:
                # Save uploaded IDD file temporarily
                idd_file_path = os.path.join('/tmp', uploaded_idd.name)
                with open(idd_file_path, 'wb') as f:
                    f.write(uploaded_idd.getbuffer())
            else:
                idd_file_path = None
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        if operation == "Convert IDF to Excel":
            st.header("🔄 IDF to Excel Conversion")
            
            if uploaded_idf and idd_file_path:
                # Display file info
                st.success(f"✅ IDF File: {uploaded_idf.name}")
                st.success(f"✅ IDD File: {os.path.basename(idd_file_path)}")
                
                # Sheet selection option
                sheet_option = st.radio(
                    "Excel Output Format:",
                    ["All sheets (ALL + individual object sheets)", "Single ALL sheet only"],
                    help="Choose whether to include individual object type sheets or just the consolidated ALL sheet"
                )
                include_individual_sheets = sheet_option == "All sheets (ALL + individual object sheets)"
                
                # Process button
                if st.button("🚀 Convert to Excel", type="primary"):
                    with st.spinner("Processing IDF file..."):
                        try:
                            # Convert uploaded file to file-like object
                            idf_content = uploaded_idf.read().decode('utf-8')
                            idf_file_object = io.StringIO(idf_content)
                            
                            # Extract data
                            extracted_data = extract_idf_data(idf_file_object, idd_file_path)
                            
                            if extracted_data:
                                # Create output file
                                output_filename = f"{os.path.splitext(uploaded_idf.name)[0]}_converted.xlsx"
                                output_path = os.path.join('/tmp', output_filename)
                                
                                # Write to Excel
                                write_data_to_excel(extracted_data, output_path, include_individual_sheets)
                                
                                # Read the file for download
                                with open(output_path, 'rb') as f:
                                    excel_data = f.read()
                                
                                # Display success message
                                st.markdown('<div class="success-box">✅ Conversion completed successfully!</div>', unsafe_allow_html=True)
                                
                                # Show statistics
                                total_records = sum(len(data) for data in extracted_data.values())
                                total_sheets = 1 + (len(extracted_data) if include_individual_sheets else 0)
                                
                                st.markdown(f"""
                                **Conversion Summary:**
                                - Total object types: {len(extracted_data)}
                                - Total records: {total_records}
                                - Excel worksheets: {total_sheets}
                                - Format: {'ALL + individual sheets' if include_individual_sheets else 'Single ALL sheet only'}
                                """)
                                
                                # Download button
                                st.download_button(
                                    label="📥 Download Excel File",
                                    data=excel_data,
                                    file_name=output_filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                                
                                # Preview first few records
                                st.subheader("📊 Data Preview")
                                preview_data = []
                                for obj_type, data in list(extracted_data.items())[:3]:  # Show first 3 object types
                                    for record in data[:5]:  # Show first 5 records per type
                                        preview_data.append({
                                            'ObjectType': obj_type,
                                            'ObjectName': record['ObjectName'],
                                            'FieldName': record['FieldName'],
                                            'Value': record['Value']
                                        })
                                
                                if preview_data:
                                    preview_df = pd.DataFrame(preview_data)
                                    st.dataframe(preview_df, use_container_width=True)
                            else:
                                st.error("❌ Failed to extract data from IDF file. Please check the file format.")
                                
                        except Exception as e:
                            st.error(f"❌ An error occurred during processing: {str(e)}")
                            st.info("Please ensure your IDF and IDD files are compatible and valid.")
            
            elif not uploaded_idf:
                st.info("👆 Please upload an IDF file to get started.")
            elif not idd_file_path:
                st.info("👆 Please select or upload an IDD file.")
        
        else:  # Update IDF from Excel
            st.header("🔄 Excel to IDF Update")
            
            # Clear warning that this feature is under development
            st.warning("⚠️ **This feature is currently under development and may not work as expected.**")
            st.info("The Excel to IDF update functionality requires further development and testing. Please use the IDF to Excel conversion feature for now.")
            
            if uploaded_idf and uploaded_excel and idd_file_path:
                # Display file info
                st.success(f"✅ Original IDF File: {uploaded_idf.name}")
                st.success(f"✅ Modified Excel File: {uploaded_excel.name}")
                st.success(f"✅ IDD File: {os.path.basename(idd_file_path)}")
                
                # Validate Excel file first
                with st.spinner("Validating Excel file..."):
                    try:
                        # Save uploaded Excel file temporarily
                        excel_path = os.path.join('/tmp', uploaded_excel.name)
                        with open(excel_path, 'wb') as f:
                            f.write(uploaded_excel.getbuffer())
                        
                        is_valid, validation_messages = validate_excel_for_update(excel_path)
                        
                        if not is_valid:
                            st.error("❌ Excel file validation failed:")
                            for msg in validation_messages:
                                st.error(f"  - {msg}")
                            st.info("Please ensure the Excel file was generated by this application and contains the required columns.")
                        else:
                            st.success("✅ Excel file format is valid for IDF updates")
                            
                            # Process button
                            if st.button("🚀 Update IDF File", type="primary"):
                                with st.spinner("Updating IDF file..."):
                                    try:
                                        # Save uploaded IDF file temporarily
                                        idf_path = os.path.join('/tmp', uploaded_idf.name)
                                        with open(idf_path, 'wb') as f:
                                            f.write(uploaded_idf.getbuffer())
                                        
                                        # Create output filename
                                        output_filename = f"{os.path.splitext(uploaded_idf.name)[0]}_updated.idf"
                                        output_path = os.path.join('/tmp', output_filename)
                                        
                                        # Perform update
                                        stats = update_idf_from_excel(idf_path, excel_path, idd_file_path, output_path)
                                        
                                        # Read the updated IDF file for download
                                        with open(output_path, 'rb') as f:
                                            updated_idf_data = f.read()
                                        
                                        # Display results
                                        st.markdown('<div class="success-box">✅ IDF update completed successfully!</div>', unsafe_allow_html=True)
                                        
                                        # Show statistics
                                        st.markdown(f"""
                                        **Update Summary:**
                                        - Total updates attempted: {stats['total_updates']}
                                        - Successful updates: {stats['successful_updates']}
                                        - Failed updates: {stats['failed_updates']}
                                        """)
                                        
                                        # Show warnings if any
                                        if stats['warnings']:
                                            st.warning(f"⚠️ Warnings ({len(stats['warnings'])}):")
                                            for warning in stats['warnings'][:5]:
                                                st.write(f"  - {warning}")
                                        
                                        # Show errors if any
                                        if stats['errors']:
                                            st.error(f"❌ Errors ({len(stats['errors'])}):")
                                            for error in stats['errors'][:5]:
                                                st.write(f"  - {error}")
                                        
                                        # Show sample updates
                                        if stats['updated_objects']:
                                            st.subheader("📝 Sample Updates")
                                            sample_updates = stats['updated_objects'][:5]
                                            update_data = []
                                            for update in sample_updates:
                                                update_data.append({
                                                    'Object Type': update['object_type'],
                                                    'Object Name': update['object_name'],
                                                    'Field Name': update['field_name'],
                                                    'Old Value': update['old_value'],
                                                    'New Value': update['new_value']
                                                })
                                            
                                            update_df = pd.DataFrame(update_data)
                                            st.dataframe(update_df, use_container_width=True)
                                        
                                        # Download button
                                        st.download_button(
                                            label="📥 Download Updated IDF File",
                                            data=updated_idf_data,
                                            file_name=output_filename,
                                            mime="text/plain"
                                        )
                                        
                                    except Exception as e:
                                        st.error(f"❌ An error occurred during IDF update: {str(e)}")
                                        st.info("Please ensure your files are compatible and valid.")
                        
                    except Exception as e:
                        st.error(f"❌ Error validating Excel file: {str(e)}")
            
            elif not uploaded_idf:
                st.info("👆 Please upload an IDF file to get started.")
            elif not uploaded_excel:
                st.info("👆 Please upload a modified Excel file.")
            elif not idd_file_path:
                st.info("👆 Please select or upload an IDD file.")
    
    with col2:
        st.header("ℹ️ Instructions")
        
        if operation == "Convert IDF to Excel":
            st.markdown("""
            **How to use (IDF to Excel):**
            1. Upload your IDF file
            2. Select the appropriate IDD file
            3. Click "Convert to Excel"
            4. Download the converted Excel file
            
            **Supported Formats:**
            - IDF files (.idf)
            - IDD files (.idd)
            
            **Output Features:**
            - Multiple worksheets for each object type
            - Comprehensive "ALL" worksheet
            - Structured parameter data
            - Easy-to-analyze format
            """)
        else:
            st.markdown("""
            **How to use (Excel to IDF):**
            1. Upload your original IDF file
            2. Upload the modified Excel file
            3. Select the appropriate IDD file
            4. Click "Update IDF File"
            5. Download the updated IDF file
            
            **Requirements:**
            - Excel file must be generated by this application
            - Contains required columns: ObjectName, FieldName, Value
            - Object identifiers must match the original IDF
            
            **Update Process:**
            - Exact object matching by name and type
            - Field-by-field value updates
            - Comprehensive error reporting
            - Update statistics and logs
            """)
        
        st.header("📋 Object Types")
        if uploaded_idf and idd_file_path and 'extracted_data' in locals():
            object_types = list(extracted_data.keys())
            for obj_type in object_types[:10]:  # Show first 10
                st.write(f"• {obj_type}")
            if len(object_types) > 10:
                st.write(f"... and {len(object_types) - 10} more")
        else:
            st.info("Object types will appear here after conversion.")

if __name__ == "__main__":
    main()