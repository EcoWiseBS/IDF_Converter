# IDF Converter

A web application for converting EnergyPlus IDF files to Excel format and updating IDF files from modified Excel data.

## 🚀 Features

### ✅ Available Features
- **IDF to Excel Conversion**: Extract all parameters from EnergyPlus IDF files and organize them into structured Excel worksheets
- **Multi-sheet Output**: Generate both consolidated "ALL" worksheets and individual object type sheets
- **Version Detection**: Automatically detect IDF file versions and suggest appropriate IDD files
- **Web Interface**: User-friendly Streamlit web application with drag-and-drop file uploads
- **Batch Processing**: Support for multiple IDF file versions (V8.0.0 to V25.1.0)

### 🔧 Under Development
- **Excel to IDF Update**: Apply modifications from Excel files back to IDF files (currently in development)

## 📋 Requirements

- Python 3.8+
- EnergyPlus IDF files (.idf)
- EnergyPlus IDD files (.idd)

## 🛠 Installation

1. **Clone the repository:**
   ```bash
   git clone <repository-url>
   cd IDF_Converter
   ```

2. **Create a virtual environment (recommended):**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Usage

### Web Application (Recommended)

1. **Start the Streamlit application:**
   ```bash
   streamlit run app/main.py
   ```

2. **Open your browser** to the displayed URL (typically `http://localhost:8501`)

3. **Convert IDF to Excel:**
   - Upload your IDF file
   - Select or upload an appropriate IDD file
   - Choose output format (all sheets or single ALL sheet)
   - Click "Convert to Excel"
   - Download the generated Excel file

### Command Line Usage

The application can also be used as a standalone script:

```python
# Edit the paths in src/idf_processor.py and run:
python src/idf_processor.py
```

## 📁 Project Structure

```
IDF_Converter/
├── app/
│   └── main.py              # Streamlit web application
├── src/
│   ├── idf_processor.py     # Core IDF to Excel conversion logic
│   ├── idf_updater.py       # Excel to IDF update functionality
│   └── __init__.py
├── IDD_FILES/               # EnergyPlus IDD files for different versions
│   ├── V22-1-0-Energy+.idd
│   ├── V23-1-0-Energy+.idd
│   └── ...
├── requirements.txt         # Python dependencies
└── README.md
```

## 🔧 Dependencies

- **streamlit**: Web application framework
- **pandas**: Data manipulation and Excel I/O
- **eppy**: EnergyPlus IDF file processing
- **openpyxl**: Excel file handling

## 📊 Output Format

### Excel File Structure
- **ALL Sheet**: Consolidated view of all object parameters
- **Individual Sheets**: Separate worksheets for each object type (e.g., `Building`, `Zone`, `Material`)

### Data Columns
- `ObjectType`: EnergyPlus object type
- `ObjectName`: Name of the specific object instance
- `FieldName`: Parameter field name
- `Value`: Parameter value
- `Unit`: Field unit (placeholder for future enhancement)

## 🔍 Supported EnergyPlus Versions

The application includes IDD files for:
- V8.0.0
- V22.1.0, V22.2.0
- V23.1.0, V23.2.0
- V24.1.0, V24.2.0
- V25.1.0

## 🐛 Troubleshooting

### Common Issues

1. **IDF Version Mismatch**
   - Ensure the IDD file matches your IDF file version
   - Use the automatic version detection feature

2. **File Encoding Issues**
   - IDF files should be UTF-8 encoded
   - Check for special characters in file paths

3. **Missing Dependencies**
   - Run `pip install -r requirements.txt` to ensure all dependencies are installed

### Error Messages

- **"IDF file not found"**: Check file path and permissions
- **"IDD file not found"**: Ensure IDD files are in the `IDD_FILES` directory
- **"Version detection failed"**: Manual IDD file selection may be required

## 🤝 Contributing

Contributions are welcome! Please feel free to submit pull requests or open issues for:
- Bug fixes
- New features
- Documentation improvements
- Test cases

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- **EnergyPlus**: Building energy simulation software
- **Eppy**: Python library for EnergyPlus IDF files
- **Streamlit**: For the web application framework

## 📞 Support

For questions or support:
1. Check the troubleshooting section above
2. Review the code documentation
3. Open an issue on the project repository

---

**Note**: The Excel to IDF update functionality is currently under development and may not work as expected. Use the IDF to Excel conversion feature for reliable results.
