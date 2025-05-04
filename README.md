# TRAK File Builder

A cross-platform GUI application for creating and uploading TRAK delimited files for music inventory management.

## Features

- Create new item spreadsheets with required and optional fields
- Process existing Excel spreadsheets into the TRAK format
- Auto-calculate prices based on cost for CDs and LPs
- Preview and edit generated files before finalizing
- Cross-platform file upload capability (Windows, macOS, Linux)

## Requirements

- Python 3.7 or higher
- Required Python packages:
  - pandas
  - openpyxl
  - tkinter (usually included with Python)
  - paramiko (for file uploads)

## Installation

### Windows

1. Install Python from [python.org](https://python.org)
   - Make sure to check "Add Python to PATH" during installation
2. Open Command Prompt and run:
   ```
   pip install pandas openpyxl paramiko
   ```
3. Download or clone this repository to your computer

### macOS

1. Install Python using Homebrew:
   ```
   brew install python
   ```
   Or download from [python.org](https://python.org)
2. Open Terminal and run:
   ```
   pip3 install pandas openpyxl paramiko
   ```
3. Download or clone this repository to your computer

### Linux

1. Install Python and tkinter:
   ```
   sudo apt-get update
   sudo apt-get install python3 python3-pip python3-tk
   ```
   (Use your distribution's package manager if not using apt)
2. Install required packages:
   ```
   pip3 install pandas openpyxl paramiko
   ```
3. Download or clone this repository to your computer

## Usage

1. Run the application:
   - Windows: Double-click on `main.py` or run `python main.py` in Command Prompt
   - macOS/Linux: Run `python3 main.py` in Terminal

2. Using the application:
   - **Step 1**: Create a new spreadsheet or select an existing Excel file
   - **Step 2**: Review and edit the processed spreadsheet (opens in your default spreadsheet application)
   - **Step 3**: Generate and review the TRAK delimited file
   - **Step 4**: Upload the file to your TRAK server

3. Upload functionality:
   - Enter your server credentials (username, host, target path)
   - Click "Upload TRAK File"

## File Format

The application works with spreadsheets containing these columns:

### Required Columns (in red):
- UPC: Universal Product Code
- TITLE: Product title
- ARTIST: Artist name
- MANUF: Manufacturer
- GENRE: Music genre
- CONFIG: Configuration (CD, LP)
- COST: Wholesale cost

### Optional Columns:
- DEPT: Department code (automatically filled based on CONFIG)
- MISC: Miscellaneous info (defaults to MANUF if empty)
- LIST: List price (calculated from COST if empty)
- PRICE: Selling price (calculated from COST if empty)
- VENDOR: Vendor name (defaults to MANUF if empty)

## Troubleshooting

### Application Won't Start
- Ensure Python is installed and in your PATH
- Check that all required packages are installed

### Upload Issues
- For Windows users: The application will automatically handle SSH connections using paramiko without external tools
- If you see "The 'paramiko' package is required for file uploads" message, click "Yes" to install it automatically
- For password issues: Check that your server credentials are correct
- If upload fails, check your network connection and server availability

## License

This software is provided as-is without warranty.

## Contact

For support, please contact the developer of this application.