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

### Step 1: Open the Application
1. Launch **Trak File Builder** from your desktop.  
2. The main dashboard will appear with options for input, output, and building files.

---

### Step 2: Prepare the Input Spreadsheet
1. **Browse for Input File**  
   - Click **Browse** to select an existing spreadsheet.  
   - Ensure the spreadsheet contains all **required columns**.  

2. **Create a Blank Template (if needed)**  
   - Click **Create New** to generate a template spreadsheet.  
   - Choose a **file path** where the template will be created and saved.  
   - ⚠️ *Do not change the file location after selecting this path.*  

3. **Name the File** appropriately.

---

### Step 3: Fill in Spreadsheet Data
- **Required Columns (red):** Must be filled manually.  
- **Auto-populated Columns (black):** Automatically filled based on required data.  
   - You may overwrite these values if necessary by entering your own data.  
- Save the spreadsheet once all required information is complete. 

---

### Step 4: Build and Review the Spreadsheet
1. Change the **output directory** if desired (*recommended to keep the default*).  
2. Click **Build Spreadsheet**.  
3. Click **View/Edit Spreadsheet** to review.  
   - Verify auto-populated data.  
   - Make any final changes.  
   - Save again if changes are made.  

---

### Step 5: Build and Review the Trak File
1. Click **Build Trak File**.  
2. Click **View/Edit Trak File** to verify the file.  
3. Confirm that all information is correct before uploading.  

---

### Step 6: Upload the Trak File
1. Enter your **username, host, and target path**.  
   - These will auto-populate with default values.  
   - ⚠️ *Do not change unless you are certain of what you’re doing.*  
2. Click **Upload Trak File**.  
3. Enter the password: `trak`.  

---

### Step 7: Verify Upload in PuTTY
1. Open a **PuTTY session** into Trak.  
2. From the main page, navigate through the menu:  
   - `[k] Optional Modules` →  
   - `[d] Database Menu` →  
   - `[a] Alternate Database Posting` →  
   - `[t] Trak Delimited Database` →  
   - `[u] Update Database`  
3. Search for UPCs from the upload to verify the file was installed correctly.  

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