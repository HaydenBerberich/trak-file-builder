import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import os
import sys
import subprocess

# Function to determine the selling price for CDs based on cost
def get_cd_price(cost):
    if cost <= 0:
        return None
    elif cost <= 0.99:
        return 1.99
    elif cost <= 1.99:
        return 3.99
    elif cost <= 2.99:
        return 5.99
    elif cost <= 3.99:
        return 7.99
    elif cost <= 4.99:
        return 9.99
    elif cost <= 5.99:
        return 11.99
    elif cost <= 6.99:
        return 13.99
    elif cost <= 7.99:
        return 14.99
    elif cost <= 9.99:
        return 15.99
    elif cost <= 11.99:
        return 16.99
    elif cost <= 12.99:
        return 17.99
    elif cost <= 13.99:
        return 21.99
    elif cost <= 14.99:
        return 22.99
    elif cost <= 15.99:
        return 24.99
    elif cost <= 16.99:
        return 25.99
    elif cost <= 17.99:
        return 26.99
    elif cost <= 18.99:
        return 27.99
    elif cost <= 19.99:
        return 29.99
    elif cost <= 20.99:
        return 31.99
    else:
        return cost * 1.4  # If cost > 20.99, price = cost * 1.4

# Function to determine the selling price for LPs based on cost
def get_lp_price(cost):
    if cost <= 0:
        return None
    elif cost <= 0.99:
        return 1.99
    elif cost <= 1.99:
        return 3.99
    elif cost <= 2.99:
        return 5.99
    elif cost <= 3.99:
        return 7.99
    elif cost <= 4.99:
        return 9.99
    elif cost <= 5.99:
        return 11.99
    elif cost <= 6.99:
        return 13.99
    elif cost <= 7.99:
        return 14.99
    elif cost <= 9.99:
        return 15.99
    elif cost <= 10.99:
        return 19.99
    elif cost <= 11.99:
        return 22.99
    elif cost <= 12.99:
        return 22.99
    elif cost <= 13.99:
        return 23.99
    elif cost <= 14.99:
        return 24.99
    elif cost <= 15.99:
        return 25.99
    elif cost <= 16.99:
        return 27.99
    elif cost <= 17.99:
        return 29.99
    elif cost <= 18.99:
        return 30.99
    elif cost <= 19.99:
        return 31.99
    elif cost <= 20.99:
        return 33.99
    elif cost <= 21.99:
        return 34.99
    elif cost <= 22.99:
        return 35.99
    elif cost <= 23.99:
        return 36.99
    elif cost <= 24.99:
        return 38.99
    elif cost <= 25.99:
        return 39.99
    elif cost <= 26.99:
        return 41.99
    elif cost <= 27.99:
        return 44.99
    elif cost <= 28.99:
        return 45.99
    elif cost <= 29.99:
        return 46.99
    elif cost <= 30.99:
        return 47.99
    elif cost <= 31.99:
        return 48.99
    elif cost <= 32.99:
        return 49.99
    elif cost <= 33.99:
        return 50.99
    elif cost <= 34.99:
        return 52.99
    elif cost <= 35.99:
        return 54.99
    elif cost <= 36.99:
        return 55.99
    elif cost <= 37.99:
        return 58.99
    elif cost <= 38.99:
        return 59.99
    elif cost <= 39.99:
        return 61.99
    elif cost <= 40.99:
        return 62.99
    elif cost <= 41.99:
        return 64.99
    elif cost <= 42.99:
        return 65.99
    elif cost <= 43.99:
        return 66.99
    elif cost <= 44.99:
        return 68.99
    elif cost <= 45.99:
        return 69.99
    elif cost <= 46.99:
        return 71.99
    elif cost <= 47.99:
        return 73.99
    elif cost <= 48.99:
        return 74.99
    elif cost <= 49.99:
        return 76.99
    else:
        return cost * 1.4  # If cost > 49.99, price = cost * 1.4

# Process the input data and create a complete Excel file
def process_input_data(input_file_path, output_dir):
    # Read the Excel file into a DataFrame, ensuring the UPC column is read as a string
    df = pd.read_excel(input_file_path, dtype={'UPC': str})

    # Drop rows where all of the required columns are NaN (blank)
    df = df.dropna(subset=['UPC', 'TITLE', 'ARTIST', 'MANUF', 'GENRE', 'CONFIG', 'DEPT', 'MISC', 'PRICE', 'VENDOR', 'COST'], how='all')

    # Replace NaN values with an empty string
    df = df.fillna('')

    # Set MISC and VENDOR to MANUF if they're not provided
    df['MISC'] = df.apply(lambda row: row['MANUF'] if not row['MISC'] else row['MISC'], axis=1)
    df['VENDOR'] = df.apply(lambda row: row['MANUF'] if not row['VENDOR'] else row['VENDOR'], axis=1)

    # Remove dollar signs from PRICE, COST, and LIST columns
    df['PRICE'] = df['PRICE'].replace({r'\$': ''}, regex=True)
    df['COST'] = df['COST'].replace({r'\$': ''}, regex=True)
    df['LIST'] = df['LIST'].replace({r'\$': ''}, regex=True)

    # Process each row to fill in missing values
    processed_rows = []
    
    for index, row in df.iterrows():
        processed_row = row.copy()
        
        # Set department based on CONFIG if DEPT is empty
        if not processed_row['DEPT']:
            if processed_row['CONFIG'] == 'CD':
                processed_row['DEPT'] = '02'
            elif processed_row['CONFIG'] == 'LP':
                processed_row['DEPT'] = '01'
        else:
            # Ensure DEPT is formatted as two digits
            dept = str(int(processed_row['DEPT']))
            if len(dept) == 1:
                processed_row['DEPT'] = '0' + dept
            else:
                processed_row['DEPT'] = dept
                
        # Calculate LIST and PRICE based on CONFIG if missing
        cost_value = float(processed_row['COST']) if processed_row['COST'] else 0
        
        if processed_row['CONFIG'] == 'CD':
            # If price is missing, determine it from the cost
            if not processed_row['PRICE']:
                calculated_price = get_cd_price(cost_value)
                if calculated_price:
                    processed_row['PRICE'] = format(calculated_price, '.2f')
            
            # If list price is missing, use the same calculation as price
            if not processed_row['LIST']:
                calculated_list = get_cd_price(cost_value)
                if calculated_list:
                    processed_row['LIST'] = format(calculated_list, '.2f')
                    
        elif processed_row['CONFIG'] == 'LP':
            # If price is missing, determine it from the cost using the LP pricing table
            if not processed_row['PRICE']:
                calculated_price = get_lp_price(cost_value)
                if calculated_price:
                    processed_row['PRICE'] = format(calculated_price, '.2f')
            
            # If list price is missing, use the same calculation as price
            if not processed_row['LIST']:
                calculated_list = get_lp_price(cost_value)
                if calculated_list:
                    processed_row['LIST'] = format(calculated_list, '.2f')
        
        # Format cost value
        if processed_row['COST']:
            processed_row['COST'] = format(float(processed_row['COST']), '.2f')
            
        processed_rows.append(processed_row)
    
    # Create a new DataFrame with processed data
    processed_df = pd.DataFrame(processed_rows)
    
    # Reorder columns to match the desired output format
    ordered_columns = ['UPC', 'TITLE', 'ARTIST', 'MANUF', 'GENRE', 'CONFIG', 'DEPT', 'MISC', 'LIST', 'PRICE', 'VENDOR', 'COST']
    processed_df = processed_df[ordered_columns]
    
    # Create output file paths
    excel_output_path = os.path.join(output_dir, 'trakdelim.xlsx')
    
    # Save to a new Excel file
    processed_df.to_excel(excel_output_path, index=False)
    
    return processed_df, excel_output_path

# Generate the delimited text file from the processed data
def generate_delimited_file(df, output_dir):
    # Create output file path
    text_output_path = os.path.join(output_dir, 'trakdelim.txt')
    
    # Open the output file for writing in the specified directory
    with open(text_output_path, 'w') as file:
        
        # Iterate over each row in the DataFrame
        for index, row in df.iterrows():
            # Format the numeric values for the delimited file (no decimal points)
            # Ensure values are strings before calling replace
            list_price = str(row['LIST']).replace('.', '') if pd.notna(row['LIST']) else ''
            price = str(row['PRICE']).replace('.', '') if pd.notna(row['PRICE']) else ''
            cost = str(row['COST']).replace('.', '') if pd.notna(row['COST']) else ''
            
            # Format the row data according to the specified layout
            formatted_row = (
                f"C|{row['UPC']}|{row['TITLE']}|{row['ARTIST']}|{row['MANUF']}|||{row['GENRE']}|||{row['MISC']}|{row['CONFIG']}|||{row['DEPT']}|{list_price}||||||{row['VENDOR']}|{cost}|||||||||{price}"
            )

            # Write the formatted data to the output file
            file.write(formatted_row + '\n')
            
    return text_output_path

# Function to open a file with the default application
def open_with_default_app(file_path):
    if not os.path.exists(file_path):
        return False
        
    if sys.platform.startswith('darwin'):  # macOS
        subprocess.call(('open', file_path))
    elif os.name == 'nt':  # Windows
        os.startfile(file_path)
    elif os.name == 'posix':  # Linux
        # Redirect stderr to /dev/null to suppress warnings
        with open(os.devnull, 'w') as devnull:
            subprocess.call(('xdg-open', file_path), stderr=devnull)
    return True

# GUI Application
class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TRAK File Converter")
        self.root.geometry("700x600")
        
        # Variables
        self.input_file_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.output_dir.set(os.getcwd())  # Default to current directory
        self.excel_output_path = tk.StringVar()
        self.text_output_path = tk.StringVar()
        
        # SCP upload variables
        self.scp_username = tk.StringVar(value="trak")
        self.scp_host = tk.StringVar(value="192.168.12.99")
        self.scp_target_path = tk.StringVar(value="/trak/data/trakdelim.txt")
        
        # Status variable
        self.status_text = tk.StringVar()
        self.status_text.set("Ready to convert")
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create the notebook (tabbed interface)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create pages
        self.page1 = ttk.Frame(self.notebook)
        self.page2 = ttk.Frame(self.notebook)
        self.page3 = ttk.Frame(self.notebook)
        
        # Add pages to notebook
        self.notebook.add(self.page1, text="1. Select Input")
        self.notebook.add(self.page2, text="2. Edit Spreadsheet")
        self.notebook.add(self.page3, text="3. TRAK File")
        
        # Disable pages 2 and 3 initially
        self.notebook.tab(1, state="disabled")
        self.notebook.tab(2, state="disabled")
        
        # Create content for each page
        self.create_page1()
        self.create_page2()
        self.create_page3()
        
        # Log area (common to all pages)
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar for log
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Status bar (common to all pages)
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)
        
        ttk.Label(status_frame, textvariable=self.status_text, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X)
    
    def create_page1(self):
        """Create content for the first page (Input Selection)"""
        # Input file section
        input_frame = ttk.LabelFrame(self.page1, text="Select Input Excel File", padding="20")
        input_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(input_frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=15)
        ttk.Entry(input_frame, textvariable=self.input_file_path, width=50).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=15)
        ttk.Button(input_frame, text="Browse...", command=self.browse_input_file).grid(row=0, column=2, padx=5, pady=15)
        
        # Add required and optional columns information
        ttk.Label(input_frame, text="Required Columns: UPC, TITLE, ARTIST, MANUF, GENRE, CONFIG, COST", 
                 ).grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=5, pady=2)
        ttk.Label(input_frame, text="Optional Columns: DEPT, MISC, LIST, PRICE, VENDOR",
                 ).grid(row=2, column=0, columnspan=3, sticky=tk.W, padx=5, pady=2)
        
        # Output directory
        ttk.Label(input_frame, text="Output Directory:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=15)
        ttk.Entry(input_frame, textvariable=self.output_dir, width=50).grid(row=3, column=1, sticky=tk.EW, padx=5, pady=15)
        ttk.Button(input_frame, text="Browse...", command=self.browse_output_dir).grid(row=3, column=2, padx=5, pady=15)
        
        # Process button
        button_frame = ttk.Frame(self.page1)
        button_frame.pack(fill=tk.X, padx=10, pady=20)
        
        ttk.Button(
            button_frame, 
            text="Build Spreadsheet", 
            command=self.process_to_excel,
            width=20
        ).pack(side=tk.RIGHT, padx=5)
    
    def create_page2(self):
        """Create content for the second page (Spreadsheet View/Edit)"""
        # Spreadsheet section
        sheet_frame = ttk.LabelFrame(self.page2, text="Generated Spreadsheet", padding="20")
        sheet_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(sheet_frame, text="Spreadsheet Path:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=15)
        ttk.Entry(sheet_frame, textvariable=self.excel_output_path, width=50, state="readonly").grid(row=0, column=1, sticky=tk.EW, padx=5, pady=15)
        
        button_frame = ttk.Frame(sheet_frame)
        button_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, padx=5, pady=15)
        
        ttk.Button(
            button_frame, 
            text="View/Edit Spreadsheet", 
            command=self.open_excel_file,
            width=25
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            button_frame, 
            text="Build TRAK File", 
            command=self.process_to_text,
            width=25
        ).pack(side=tk.RIGHT, padx=5)
    
    def create_page3(self):
        """Create content for the third page (TRAK File View/Upload)"""
        # Text file section
        text_frame = ttk.LabelFrame(self.page3, text="Generated TRAK File", padding="20")
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(text_frame, text="TRAK File Path:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=15)
        ttk.Entry(text_frame, textvariable=self.text_output_path, width=50, state="readonly").grid(row=0, column=1, sticky=tk.EW, padx=5, pady=15)
        
        button_frame = ttk.Frame(text_frame)
        button_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, padx=5, pady=15)
        
        ttk.Button(
            button_frame, 
            text="View/Edit TRAK File", 
            command=self.open_text_file,
            width=25
        ).pack(side=tk.LEFT, padx=5)
        
        # SCP upload section
        scp_frame = ttk.LabelFrame(self.page3, text="Upload to TRAK Server", padding="10")
        scp_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(scp_frame, text="Username:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(scp_frame, textvariable=self.scp_username, width=15).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(scp_frame, text="Host:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(scp_frame, textvariable=self.scp_host, width=20).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        
        ttk.Label(scp_frame, text="Target Path:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Entry(scp_frame, textvariable=self.scp_target_path, width=50).grid(row=2, column=1, columnspan=2, sticky=tk.EW, padx=5, pady=5)
        
        ttk.Button(
            scp_frame, 
            text="Upload TRAK File", 
            command=self.upload_file,
            width=20
        ).grid(row=3, column=1, sticky=tk.E, padx=5, pady=10)
        
        # Navigation buttons
        nav_frame = ttk.Frame(self.page3)
        nav_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(
            nav_frame, 
            text="Start Over", 
            command=self.start_over,
            width=15
        ).pack(side=tk.LEFT, padx=5)
    
    def browse_input_file(self):
        filetypes = [("Excel files", "*.xlsx"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(
            title="Select Input Excel File",
            filetypes=filetypes
        )
        if filename:
            self.input_file_path.set(filename)
            self.log("Selected input file: " + filename)
    
    def browse_output_dir(self):
        directory = filedialog.askdirectory(
            title="Select Output Directory"
        )
        if directory:
            self.output_dir.set(directory)
            self.log("Selected output directory: " + directory)
    
    def process_to_excel(self):
        """Process the input file to create the Excel spreadsheet"""
        # Get the input file path
        input_file = self.input_file_path.get()
        output_dir = self.output_dir.get()
        
        if not input_file:
            messagebox.showerror("Error", "Please select an input Excel file.")
            return
            
        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"Input file does not exist: {input_file}")
            return
        
        if not output_dir:
            messagebox.showerror("Error", "Please select an output directory.")
            return
            
        if not os.path.exists(output_dir):
            messagebox.showerror("Error", f"Output directory does not exist: {output_dir}")
            return
        
        try:
            self.status_text.set("Processing input data...")
            self.root.update()
            
            # Process the input data and create Excel
            self.log("Processing input data...")
            processed_df, excel_path = process_input_data(input_file, output_dir)
            self.excel_output_path.set(excel_path)
            
            self.log(f"Created Excel file: {excel_path}")
            self.status_text.set("Spreadsheet created successfully!")
            
            # Enable page 2 and switch to it
            self.notebook.tab(1, state="normal")
            self.notebook.select(1)
            
        except Exception as e:
            error_message = str(e)
            self.status_text.set("Error: " + error_message)
            self.log("ERROR: " + error_message)
            messagebox.showerror("Error", f"An error occurred during processing:\n{error_message}")
    
    def open_excel_file(self):
        """Open the Excel file with the default application"""
        excel_path = self.excel_output_path.get()
        if excel_path:
            if open_with_default_app(excel_path):
                self.log(f"Opened spreadsheet: {excel_path}")
            else:
                self.log(f"Error: Cannot open file {excel_path}")
                messagebox.showerror("Error", f"Cannot open file: {excel_path}")
    
    def process_to_text(self):
        """Generate the delimited text file from the Excel spreadsheet"""
        excel_path = self.excel_output_path.get()
        output_dir = self.output_dir.get()
        
        if not excel_path or not os.path.exists(excel_path):
            messagebox.showerror("Error", "Spreadsheet file not found. Please build it first.")
            return
            
        try:
            self.status_text.set("Generating TRAK file...")
            self.root.update()
            
            # Read the Excel file
            df = pd.read_excel(excel_path)
            
            # Generate the delimited text file
            self.log("Generating TRAK file...")
            text_path = generate_delimited_file(df, output_dir)
            self.text_output_path.set(text_path)
            
            self.log(f"Created TRAK file: {text_path}")
            self.status_text.set("TRAK file created successfully!")
            
            # Enable page 3 and switch to it
            self.notebook.tab(2, state="normal")
            self.notebook.select(2)
            
        except Exception as e:
            error_message = str(e)
            self.status_text.set("Error: " + error_message)
            self.log("ERROR: " + error_message)
            messagebox.showerror("Error", f"An error occurred during text file generation:\n{error_message}")
    
    def open_text_file(self):
        """Open the text file with the default text editor"""
        text_path = self.text_output_path.get()
        if text_path:
            if open_with_default_app(text_path):
                self.log(f"Opened TRAK file: {text_path}")
            else:
                self.log(f"Error: Cannot open file {text_path}")
                messagebox.showerror("Error", f"Cannot open file: {text_path}")
    
    def upload_file(self):
        """Upload the TRAK file to the remote server using SCP"""
        try:
            # Determine the file path to upload
            text_path = self.text_output_path.get()
            
            # Check if the file exists
            if not os.path.exists(text_path):
                messagebox.showerror("Error", f"File not found: {text_path}")
                return
            
            # Get SCP parameters
            username = self.scp_username.get()
            host = self.scp_host.get()
            target_path = self.scp_target_path.get()
            
            # Prompt for password
            password = simpledialog.askstring("Password", "Enter password:", show='*')
            if password is None:
                # User cancelled the password dialog
                self.log("Upload cancelled - no password provided")
                return
            
            # Create the SCP command using sshpass to provide the password
            remote_target = f"{username}@{host}:{target_path}"
            
            # Check if sshpass is available - platform-specific approach
            sshpass_available = False
            is_windows = os.name == 'nt'
            is_posix = os.name == 'posix'
            
            if is_posix:  # Linux/Unix/macOS
                sshpass_available = subprocess.run(
                    ["which", "sshpass"], 
                    capture_output=True
                ).returncode == 0
            elif is_windows:  # Windows
                # On Windows, try to check if sshpass exists in PATH
                try:
                    sshpass_available = subprocess.run(
                        ["where", "sshpass"],
                        capture_output=True,
                        shell=True
                    ).returncode == 0
                except (FileNotFoundError, subprocess.SubprocessError):
                    # 'where' command might not be available or might fail
                    sshpass_available = False
            
            # Check if scp is available
            scp_available = False
            if is_posix:  # Linux/Unix/macOS
                scp_available = subprocess.run(
                    ["which", "scp"], 
                    capture_output=True
                ).returncode == 0
            elif is_windows:  # Windows
                try:
                    scp_available = subprocess.run(
                        ["where", "scp"],
                        capture_output=True,
                        shell=True
                    ).returncode == 0
                except (FileNotFoundError, subprocess.SubprocessError):
                    scp_available = False
            
            if not scp_available:
                self.log("Error: SCP is not available on this system")
                messagebox.showerror("Error", "SCP is not available on this system. Please install an SCP client.")
                return
            
            if sshpass_available:
                # Use sshpass for password authentication
                scp_command = ["sshpass", "-p", password, "scp", text_path, remote_target]
            else:
                # If sshpass is not available, will need to enter password manually
                scp_command = ["scp", text_path, remote_target]
                self.log("Note: 'sshpass' not found on system. You will need to enter password manually in terminal.")
            
            # Log the command (without showing the password)
            if sshpass_available:
                # Don't log the actual password
                log_command = ["sshpass", "-p", "*****", "scp", text_path, remote_target]
                self.log(f"Uploading file with command: {' '.join(log_command)}")
            else:
                self.log(f"Uploading file with command: scp {text_path} {remote_target}")
            
            self.status_text.set("Uploading file...")
            self.root.update()
            
            # Execute the SCP command
            process = subprocess.Popen(
                scp_command,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                shell=is_windows  # Use shell=True for Windows
            )
            
            stdout, stderr = process.communicate()
            
            if process.returncode == 0:
                self.log("Upload successful!")
                self.status_text.set("Upload completed successfully")
                messagebox.showinfo("Success", "File uploaded successfully!")
            else:
                error_message = stderr.strip()
                self.log(f"Upload failed: {error_message}")
                self.status_text.set("Upload failed")
                messagebox.showerror("Error", f"Upload failed:\n{error_message}")
                
        except Exception as e:
            error_message = str(e)
            self.log(f"Upload error: {error_message}")
            self.status_text.set("Upload error")
            messagebox.showerror("Error", f"An error occurred during upload:\n{error_message}")
    
    def start_over(self):
        """Reset the application to start a new conversion"""
        # Clear the output paths
        self.excel_output_path.set("")
        self.text_output_path.set("")
        
        # Disable pages 2 and 3
        self.notebook.tab(1, state="disabled")
        self.notebook.tab(2, state="disabled")
        
        # Switch to page 1
        self.notebook.select(0)
        
        self.status_text.set("Ready to convert")
        self.log("Started a new conversion process")
    
    def log(self, message):
        """Add a message to the log text area"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

# Main execution
if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()