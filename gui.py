"""
GUI Application for TRAK File Builder.
"""
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import subprocess
import pandas as pd

from utils import open_with_default_app
from file_processor import process_input_data, generate_delimited_file, create_new_spreadsheet

class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TRAK File Converter")
        self.root.geometry("700x800")  # Increased height from 600 to 700
        self.root.minsize(600, 600)  # Increased minimum height from 500 to 600
        
        # Configure root to be responsive
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
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
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")  # Use grid instead of pack for the main frame
        
        # Configure main_frame rows and columns
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=3)  # Notebook gets more space
        main_frame.rowconfigure(1, weight=1)  # Log area gets less space
        
        # Create the notebook (tabbed interface)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)  # Use grid instead of pack
        
        # Create pages
        self.page1 = ttk.Frame(self.notebook)
        self.page2 = ttk.Frame(self.notebook)
        self.page3 = ttk.Frame(self.notebook)
        
        # Configure page frames to be responsive
        for page in [self.page1, self.page2, self.page3]:
            page.columnconfigure(0, weight=1)
            page.rowconfigure(0, weight=1)
        
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
        log_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)  # Use grid instead of pack
        
        # Configure log_frame
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky="nsew")  # Use grid instead of pack
        
        # Scrollbar for log - place it properly
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.config(yscrollcommand=scrollbar.set)
    
    def create_page1(self):
        """Create content for the first page (Input Selection)"""
        # Input file section
        input_frame = ttk.LabelFrame(self.page1, text="Select Input Excel File", padding="20")
        input_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Configure input_frame grid weights
        input_frame.columnconfigure(0, weight=0)  # Labels don't need to expand
        input_frame.columnconfigure(1, weight=1)  # Entry fields should expand
        input_frame.columnconfigure(2, weight=0)  # Buttons don't need to expand
        input_frame.columnconfigure(3, weight=0)  # Buttons don't need to expand
        
        ttk.Label(input_frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=15)
        ttk.Entry(input_frame, textvariable=self.input_file_path, width=50).grid(row=0, column=1, sticky=tk.EW, padx=5, pady=15)
        browse_button = ttk.Button(input_frame, text="Browse...", command=self.browse_input_file)
        browse_button.grid(row=0, column=2, padx=5, pady=15)
        
        # Add Create New button
        create_new_button = ttk.Button(input_frame, text="Create New", command=self.create_new_spreadsheet)
        create_new_button.grid(row=0, column=3, padx=5, pady=15)
        
        # Add required and optional columns information
        ttk.Label(input_frame, text="Required Columns: UPC, TITLE, ARTIST, MANUF, GENRE, CONFIG, COST", 
                 foreground="red").grid(row=1, column=0, columnspan=4, sticky=tk.W, padx=5, pady=2)
        ttk.Label(input_frame, text="Optional Columns: DEPT, MISC, LIST, PRICE, VENDOR",
                 foreground="blue").grid(row=2, column=0, columnspan=4, sticky=tk.W, padx=5, pady=2)
        
        # Output directory
        ttk.Label(input_frame, text="Output Directory:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=15)
        ttk.Entry(input_frame, textvariable=self.output_dir, width=50).grid(row=3, column=1, sticky=tk.EW, padx=5, pady=15)
        ttk.Button(input_frame, text="Browse...", command=self.browse_output_dir).grid(row=3, column=2, padx=5, pady=15)
        
        # Process button
        button_frame = ttk.Frame(self.page1)
        button_frame.pack(fill=tk.X, padx=10, pady=20)
        
        # Configure button_frame to allow button to stay right-aligned
        button_frame.columnconfigure(0, weight=1)
        
        process_button = ttk.Button(
            button_frame, 
            text="Build Spreadsheet", 
            command=self.process_to_excel,
            width=20
        )
        process_button.grid(row=0, column=0, sticky=tk.E, padx=5)
    
    def create_page2(self):
        """Create content for the second page (Spreadsheet View/Edit)"""
        # Spreadsheet section
        sheet_frame = ttk.LabelFrame(self.page2, text="Generated Spreadsheet", padding="20")
        sheet_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Configure the sheet_frame to be responsive
        sheet_frame.columnconfigure(0, weight=0)  # Label column
        sheet_frame.columnconfigure(1, weight=1)  # Entry column
        sheet_frame.rowconfigure(0, weight=0)     # Fixed height for the path row
        sheet_frame.rowconfigure(1, weight=1)     # Button frame can expand
        
        ttk.Label(sheet_frame, text="Spreadsheet Path:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=15)
        ttk.Entry(sheet_frame, textvariable=self.excel_output_path, width=50, state="readonly").grid(row=0, column=1, sticky=tk.EW, padx=5, pady=15)
        
        button_frame = ttk.Frame(sheet_frame)
        button_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, padx=5, pady=15)
        
        # Make the button frame responsive
        button_frame.columnconfigure(0, weight=1)  # Left side gets space
        button_frame.columnconfigure(1, weight=1)  # Right side gets space
        
        ttk.Button(
            button_frame, 
            text="View/Edit Spreadsheet", 
            command=self.open_excel_file,
            width=25
        ).grid(row=0, column=0, sticky=tk.W, padx=5)
        
        ttk.Button(
            button_frame, 
            text="Build TRAK File", 
            command=self.process_to_text,
            width=25
        ).grid(row=0, column=1, sticky=tk.E, padx=5)
    
    def create_page3(self):
        """Create content for the third page (TRAK File View/Upload)"""
        # Text file section
        text_frame = ttk.LabelFrame(self.page3, text="Generated TRAK File", padding="20")
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Configure text_frame to be responsive
        text_frame.columnconfigure(0, weight=0)  # Label column
        text_frame.columnconfigure(1, weight=1)  # Entry column
        text_frame.rowconfigure(0, weight=0)     # Fixed height for path row
        text_frame.rowconfigure(1, weight=1)     # Button frame can expand
        
        ttk.Label(text_frame, text="TRAK File Path:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=15)
        ttk.Entry(text_frame, textvariable=self.text_output_path, width=50, state="readonly").grid(row=0, column=1, sticky=tk.EW, padx=5, pady=15)
        
        button_frame = ttk.Frame(text_frame)
        button_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, padx=5, pady=15)
        
        # Configure button_frame to be responsive
        button_frame.columnconfigure(0, weight=1)
        
        ttk.Button(
            button_frame, 
            text="View/Edit TRAK File", 
            command=self.open_text_file,
            width=25
        ).grid(row=0, column=0, sticky=tk.W, padx=5)
        
        # SCP upload section
        scp_frame = ttk.LabelFrame(self.page3, text="Upload to TRAK Server", padding="10")
        scp_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Configure scp_frame to be responsive
        scp_frame.columnconfigure(0, weight=0)  # Label column
        scp_frame.columnconfigure(1, weight=1)  # Entry/button column
        scp_frame.columnconfigure(2, weight=0)  # Extra column if needed
        
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
        
        # Configure nav_frame to be responsive
        nav_frame.columnconfigure(0, weight=1)
        nav_frame.columnconfigure(1, weight=1)
        
        ttk.Button(
            nav_frame, 
            text="Start Over", 
            command=self.start_over,
            width=15
        ).grid(row=0, column=0, sticky=tk.W, padx=5)
    
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
    
    def create_new_spreadsheet(self):
        """Create a new spreadsheet with required and optional columns"""
        # Ask user for the save location
        filetypes = [("Excel files", "*.xlsx"), ("All files", "*.*")]
        file_path = filedialog.asksaveasfilename(
            title="Create New Spreadsheet",
            defaultextension=".xlsx",
            filetypes=filetypes
        )
        
        if not file_path:
            return  # User cancelled
            
        try:
            # Create a new spreadsheet using the file_processor module
            if create_new_spreadsheet(file_path):
                # Set the input file path to the new file
                self.input_file_path.set(file_path)
                self.log(f"Created new spreadsheet: {file_path}")
                
                # Open the file with the default spreadsheet editor
                if open_with_default_app(file_path):
                    self.log(f"Opened new spreadsheet: {file_path}")
                else:
                    self.log(f"Error: Cannot open file {file_path}")
                    messagebox.showerror("Error", f"Cannot open file: {file_path}")
            else:
                self.log("Error creating new spreadsheet")
                messagebox.showerror("Error", "Failed to create new spreadsheet")
                
        except Exception as e:
            error_message = str(e)
            self.log(f"Error creating spreadsheet: {error_message}")
            messagebox.showerror("Error", f"An error occurred while creating the spreadsheet:\n{error_message}")
    
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
            self.root.update()
            
            # Process the input data and create Excel
            self.log("Processing input data...")
            processed_df, excel_path = process_input_data(input_file, output_dir)
            self.excel_output_path.set(excel_path)
            
            self.log(f"Created Excel file: {excel_path}")
            
            # Enable page 2 and switch to it
            self.notebook.tab(1, state="normal")
            self.notebook.select(1)
            
        except Exception as e:
            error_message = str(e)
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
            self.root.update()
            
            # Read the Excel file with UPC column as string to preserve leading zeros
            df = pd.read_excel(excel_path, dtype={'UPC': str})
            
            # Generate the delimited text file
            self.log("Generating TRAK file...")
            text_path = generate_delimited_file(df, output_dir)
            self.text_output_path.set(text_path)
            
            self.log(f"Created TRAK file: {text_path}")
            
            # Enable page 3 and switch to it
            self.notebook.tab(2, state="normal")
            self.notebook.select(2)
            
        except Exception as e:
            error_message = str(e)
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
        """Upload the TRAK file to the remote server using Paramiko (cross-platform SSH/SCP)"""
        try:
            # Check if paramiko is installed
            try:
                import paramiko
            except ImportError:
                self.log("The 'paramiko' package is required for file uploads.")
                if messagebox.askyesno("Install Required Package", 
                                      "The Python package 'paramiko' is required for file uploads.\nWould you like to install it now?"):
                    self.log("Installing paramiko package...")
                    self.root.update()
                    
                    # Use pip to install paramiko
                    try:
                        subprocess.check_call([sys.executable, "-m", "pip", "install", "paramiko"])
                        self.log("Paramiko installed successfully.")
                        # Import paramiko now that it's installed
                        import paramiko
                    except Exception as e:
                        self.log(f"Error installing paramiko: {e}")
                        messagebox.showerror("Installation Error", 
                                            f"Failed to install paramiko. Please install it manually using:\npip install paramiko")
                        return
                else:
                    self.log("Upload cancelled - required package not installed")
                    return

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
                
            self.log(f"Connecting to {host} as {username}...")
            self.root.update()
                
            # Create SSH client
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            try:
                # Connect to the server
                ssh.connect(hostname=host, username=username, password=password)
                self.log("Connected to server. Uploading file...")
                self.root.update()
                
                # Upload the file using SCP
                with ssh.open_sftp() as sftp:
                    sftp.put(text_path, target_path)
                    
                self.log("Upload successful!")
                messagebox.showinfo("Success", "File uploaded successfully!")
                
            except paramiko.AuthenticationException:
                self.log("Authentication failed. Check username and password.")
                messagebox.showerror("Authentication Error", "Failed to authenticate. Check username and password.")
            except paramiko.SSHException as e:
                self.log(f"SSH connection error: {str(e)}")
                messagebox.showerror("Connection Error", f"SSH connection error: {str(e)}")
            except paramiko.sftp.SFTPError as e:
                self.log(f"SFTP error: {str(e)}")
                messagebox.showerror("SFTP Error", f"SFTP error during upload: {str(e)}")
            except Exception as e:
                self.log(f"Error during upload: {str(e)}")
                messagebox.showerror("Upload Error", f"Error during upload: {str(e)}")
            finally:
                # Close the SSH connection
                ssh.close()
                    
        except Exception as e:
            error_message = str(e)
            self.log(f"Upload error: {error_message}")
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
        
        self.log("Started a new conversion process")
    
    def log(self, message):
        """Add a message to the log text area"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.update()