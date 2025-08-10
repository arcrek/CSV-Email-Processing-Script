# CSV Email Processing Script

A Python script that processes CSV files to filter and extract email addresses based on login status and optional exclusion lists. Creates output files for account management and cleanup.

## Features

- **Interactive File Selection**: GUI dialogs for selecting input files
- **Login Status Filtering**: Filters users based on login activity
- **Email Exclusion**: Optional exclusion of specific emails from a text file (emails to keep)
- **Domain-based Output**: Automatically detects email domain and creates appropriately named output files
- **Data Validation**: Validates email formats and CSV structure
- **Error Handling**: Comprehensive error handling with user-friendly messages

## Requirements

- Python 3.6 or higher
- pandas library

## Installation

1. **Install Python dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

   Or install pandas directly:
   ```bash
   pip install pandas
   ```

2. **Verify installation:**
   ```bash
   python --version
   python -c "import pandas; print('pandas version:', pandas.__version__)"
   ```

## Usage

### Option 1: Command Line Script

```bash
python csv_email_processor.py
```

### Option 2: GUI Version (Recommended)

For a user-friendly graphical interface:

```bash
python csv_email_processor_gui.py
```

**GUI Features:**
- ✅ Simple button interface (Select/Clear for files, Start Processing)
- ✅ Clear buttons to remove file selections with smart enable/disable
- ✅ Real-time status log with progress updates
- ✅ No console commands needed
- ✅ Intuitive file selection with filename display
- ✅ Multithreaded processing to prevent UI freezing
- ✅ Success/error dialogs with detailed results

### Option 3: Standalone Executable (GUI)

**For the smallest, optimized executable:**

1. **Build the GUI executable:**
   ```bash
   python build_gui_exe.py
   ```

2. **Run the executable:**
   ```bash
   ./dist/CSVEmailProcessorGUI.exe
   ```

**Optimizations applied:**
- 🎯 Minimal file size (~8-12 MB)
- 🚀 Excluded unnecessary modules
- ⚡ Bytecode optimization (--optimize=2)
- 📦 Optional UPX compression for additional size reduction
- 🖥️ Pure GUI (no console window)
- 🔒 Debug symbols stripped

### Step-by-Step Process

#### GUI Version (Recommended):
1. **Launch**: Run `csv_email_processor_gui.py` or the GUI executable
2. **Select CSV**: Click "Select CSV" button to choose your input file
   - File must contain: `Email Address [Required]` and `Last Sign In [READ ONLY]` columns
   - Use "Clear" button to remove selection if needed
3. **Select TXT (Optional)**: Click "Select TXT" to choose emails to exclude
   - Optional file with emails to exclude from deletion (one per line)
   - Use "Clear" button to remove selection if needed
4. **Start Processing**: Click "Start Processing" button
5. **Monitor Progress**: Watch the status log for real-time updates
6. **Complete**: Success dialog shows output file location

#### Command Line Version:
1. **Select CSV File**: Dialog opens asking for input CSV file
2. **Optional TXT File**: Choose whether to select exclusion file
3. **Processing**: Automatic processing with console output
4. **Output**: Creates `[domain]-to-delete.csv` file

#### Processing Logic:
- ✅ **Filters**: Users who have logged in (NOT "Never logged in") for deletion
- ✅ **Excludes**: Emails found in optional TXT file from deletion  
- ✅ **Validates**: Email format and CSV structure
- ✅ **Outputs**: `[domain]-to-delete.csv` with `primaryEmail` column

### Input CSV Format

Expected CSV structure:
```csv
First Name [Required],Last Name [Required],Email Address [Required],Status [READ ONLY],Last Sign In [READ ONLY],Email Usage [READ ONLY]
John,Doe,john.doe@company.com,Active,2024-01-15,1.2GB
Jane,Smith,jane.smith@company.com,Active,Never logged in,0.0GB
```

### Output Format

Generated file: `company-to-delete.csv`
```csv
primaryEmail
john.doe@company.com
```

### Optional Exclusion File Format

TXT file with emails to exclude from deletion (one per line):
```
admin@company.com
support@company.com
noreply@company.com
```

## Script Logic

1. **Login Status Filtering**: 
   - **Includes** users who have logged in (NOT "Never logged in") in the deletion list
   - **Excludes** users with "Never logged in" status from deletion
   - Rationale: Active accounts that have been used are identified for potential deletion or management

2. **Email Exclusion**:
   - Removes any emails found in the optional TXT file from the deletion list
   - Case-insensitive matching

3. **Domain Detection**:
   - Automatically extracts domain from the first email address
   - Warns if multiple domains are detected

## Error Handling

The script handles various error conditions:

- **Missing Required Columns**: Validates CSV structure
- **Invalid Email Formats**: Skips malformed emails with warnings
- **File Access Errors**: Handles permission and file not found errors
- **Empty Results**: Warns when no data remains after filtering
- **Multiple Domains**: Warns and uses the primary domain

## Example Usage Scenarios

### Scenario 1: Basic Cleanup
- Input: CSV with mixed login statuses
- No exclusion file
- Output: All users who have logged in (active accounts)

### Scenario 2: Advanced Filtering
- Input: CSV with user data
- Exclusion file: Admin and service accounts
- Output: Active users excluding protected accounts

### Scenario 3: Domain Migration
- Process multiple CSV files from different domains
- Generate separate deletion lists per domain

## Troubleshooting

### Common Issues

1. **"Missing required columns" error**:
   - Ensure CSV has `Email Address [Required]` and `Last Sign In [READ ONLY]` columns
   - Check for extra spaces or different capitalization

2. **"No valid email addresses found"**:
   - Verify email format in CSV (must contain @ symbol)
   - Check if all users have "Never logged in" (no active users)

3. **Permission denied when saving**:
   - Ensure write permissions in the CSV file directory
   - Close the input CSV if it's open in Excel

4. **Empty output file**:
   - All users may have "Never logged in" (no active accounts)
   - All emails may be in the exclusion list

### Dependencies Not Found

If you get `ModuleNotFoundError: No module named 'pandas'`:
```bash
pip install pandas
```

If you get tkinter errors on Linux:
```bash
sudo apt-get install python3-tk
```

### Building Executable Issues

**If you get "No module named 'secrets'" error:**

1. **Update dependencies:**
   ```bash
   pip install --upgrade pandas numpy pyinstaller
   ```

**General build troubleshooting:**

1. **PyInstaller not found:**
   ```bash
   pip install pyinstaller
   ```

2. **Missing dependencies in executable:**
   - Update pandas/numpy to latest versions
   - Ensure all required modules are installed

3. **Build process fails:**
   - Check if antivirus is blocking PyInstaller
   - Try running as administrator (Windows)
   - Ensure all dependencies are up to date

4. **Executable file size:**
   - Standard build: ~8-12 MB
   - With UPX compression: ~4-8 MB (additional optimization)

## Expected Behavior

The script processes CSV files containing user data and creates a deletion list:
- **Filters**: Users who have logged in (excludes "Never logged in" users)
- **Processes**: Email validation and domain detection
- **Outputs**: Domain-specific CSV file with emails ready for account management
- **Excludes**: Any emails specified in the optional TXT exclusion file

The script identifies active users who have actually used their accounts for potential deletion or management.

## Technical Details

### Command Line Version (`csv_email_processor.py`)
- **Framework**: Pure Python with tkinter dialogs and pandas
- **Interface**: Console output with GUI file dialogs
- **Threading**: Single-threaded execution

### GUI Version (`csv_email_processor_gui.py`) 
- **Framework**: tkinter GUI with pandas for data processing
- **Interface**: Full graphical interface with Select/Clear buttons and real-time status log
- **File Management**: Smart file selection with clear buttons and filename display
- **Threading**: Multi-threaded processing to prevent UI freezing
- **User Experience**: Real-time status updates, success/error dialogs with detailed results
- **Layout**: Responsive grid layout (600x500 resizable window)
- **Status Logging**: Scrollable text area with automatic scrolling to latest updates

### File Size Optimization (Executable)
- **Optimized build**: ~8-12 MB (excluded unnecessary modules)
- **UPX compressed**: ~4-8 MB (additional compression if UPX available)
- **Excluded modules**: matplotlib, test, unittest, doctest, pdb, PIL, setuptools, pkg_resources
- **Included optimizations**: bytecode optimization, debug symbol stripping, windowed mode

### General
- **File Handling**: Supports various CSV encodings and formats
- **Memory Efficient**: Processes large CSV files using pandas
- **Cross-platform**: Works on Windows, macOS, and Linux
- **Error Handling**: Comprehensive validation and user feedback

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify your CSV file format matches the expected structure
3. Ensure all dependencies are installed correctly

