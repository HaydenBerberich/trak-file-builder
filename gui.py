"""
GUI Application for TRAK File Builder.
Supports optional multi-hop SSH upload (jump host) with site selection (St. Louis, Springfield, or Custom).
"""
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import pandas as pd

from utils import open_with_default_app
from file_processor import process_input_data, generate_delimited_file, create_new_spreadsheet


class FileConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TRAK File Converter")
        self.root.geometry("700x850")
        self.root.minsize(600, 700)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # ---------- Site Configurations ----------
        self.sites = {
            "St. Louis": {"jump": "100.107.39.118", "final": "192.168.12.99", "user": "trak", "password": "trak"},
            "Springfield": {"jump": "100.118.199.62", "final": "192.168.1.99", "user": "trak", "password": "trak"},
            "Custom": {"jump": "", "final": "", "user": "", "password": ""}
        }
        default_site = "St. Louis"

        # ---------- Variables ----------
        self.input_file_path = tk.StringVar()
        self.output_dir = tk.StringVar(value=os.getcwd())
        self.excel_output_path = tk.StringVar()
        self.text_output_path = tk.StringVar()

        self.site_selection = tk.StringVar(value=default_site)
        self.jump_host = tk.StringVar(value=self.sites[default_site]["jump"])
        self.final_host = tk.StringVar(value=self.sites[default_site]["final"])
        self.scp_username = tk.StringVar(value=self.sites[default_site]["user"])
        self.scp_password = tk.StringVar(value=self.sites[default_site]["password"])
        self.scp_target_path = tk.StringVar(value="/trak/data/trakdelim.txt")

        # ---------- Layout ----------
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=3)
        main_frame.rowconfigure(1, weight=1)

        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        self.page1, self.page2, self.page3 = ttk.Frame(self.notebook), ttk.Frame(self.notebook), ttk.Frame(self.notebook)
        for p in [self.page1, self.page2, self.page3]:
            p.columnconfigure(0, weight=1)
            p.rowconfigure(0, weight=1)

        self.notebook.add(self.page1, text="1. Select Input")
        self.notebook.add(self.page2, text="2. Edit Spreadsheet")
        self.notebook.add(self.page3, text="3. TRAK File")
        self.notebook.tab(1, state="disabled")
        self.notebook.tab(2, state="disabled")

        # Build UI pages
        self.create_page1()
        self.create_page2()
        self.create_page3()

        # ---------- Log Area ----------
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        log_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.config(yscrollcommand=scrollbar.set)

    # ---------- PAGE 1 ----------
    def create_page1(self):
        frame = ttk.LabelFrame(self.page1, text="Select Input Excel File", padding="20")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Excel File:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.input_file_path).grid(row=0, column=1, sticky=tk.EW, padx=5)
        ttk.Button(frame, text="Browse...", command=self.browse_input_file).grid(row=0, column=2, padx=5)
        ttk.Button(frame, text="Create New", command=self.create_new_spreadsheet).grid(row=0, column=3, padx=5)

        ttk.Label(frame, text="Required Columns: UPC, TITLE, ARTIST, MANUF, GENRE, CONFIG, COST", foreground="red")\
            .grid(row=1, column=0, columnspan=4, sticky=tk.W)
        ttk.Label(frame, text="Optional Columns: DEPT, MISC, LIST, PRICE, VENDOR", foreground="blue")\
            .grid(row=2, column=0, columnspan=4, sticky=tk.W)

        ttk.Label(frame, text="Output Directory:").grid(row=3, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.output_dir).grid(row=3, column=1, sticky=tk.EW, padx=5)
        ttk.Button(frame, text="Browse...", command=self.browse_output_dir).grid(row=3, column=2, padx=5)

        ttk.Button(self.page1, text="Build Spreadsheet", command=self.process_to_excel, width=20)\
            .pack(pady=20)

    # ---------- PAGE 2 ----------
    def create_page2(self):
        frame = ttk.LabelFrame(self.page2, text="Generated Spreadsheet", padding="20")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Spreadsheet Path:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.excel_output_path, state="readonly").grid(row=0, column=1, sticky=tk.EW, padx=5)
        ttk.Button(frame, text="View/Edit Spreadsheet", command=self.open_excel_file, width=25)\
            .grid(row=1, column=0, sticky=tk.W, padx=5, pady=10)
        ttk.Button(frame, text="Build TRAK File", command=self.process_to_text, width=25)\
            .grid(row=1, column=1, sticky=tk.E, padx=5, pady=10)

    # ---------- PAGE 3 ----------
    def create_page3(self):
        frame = ttk.LabelFrame(self.page3, text="Generated TRAK File", padding="20")
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="TRAK File Path:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.text_output_path, state="readonly").grid(row=0, column=1, sticky=tk.EW, padx=5)
        ttk.Button(frame, text="View/Edit TRAK File", command=self.open_text_file).grid(row=1, column=0, sticky=tk.W, pady=5)

        # --- Upload section ---
        upload_frame = ttk.LabelFrame(self.page3, text="Upload via SSH (Optional Jump Host)", padding="10")
        upload_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        upload_frame.columnconfigure(1, weight=1)

        ttk.Label(upload_frame, text="Select Location:").grid(row=0, column=0, sticky=tk.W)
        site_dropdown = ttk.Combobox(upload_frame, textvariable=self.site_selection,
                                     values=list(self.sites.keys()), state="readonly")
        site_dropdown.grid(row=0, column=1, sticky=tk.W, pady=3)
        site_dropdown.bind("<<ComboboxSelected>>", self.update_site_settings)

        ttk.Label(upload_frame, text="Jump Host:").grid(row=1, column=0, sticky=tk.W)
        self.jump_entry = ttk.Entry(upload_frame, textvariable=self.jump_host, width=25, state="readonly")
        self.jump_entry.grid(row=1, column=1, sticky=tk.W, pady=3)

        ttk.Label(upload_frame, text="Final Host:").grid(row=2, column=0, sticky=tk.W)
        self.final_entry = ttk.Entry(upload_frame, textvariable=self.final_host, width=25, state="readonly")
        self.final_entry.grid(row=2, column=1, sticky=tk.W, pady=3)

        ttk.Label(upload_frame, text="Username:").grid(row=3, column=0, sticky=tk.W)
        self.user_entry = ttk.Entry(upload_frame, textvariable=self.scp_username, width=25, state="readonly")
        self.user_entry.grid(row=3, column=1, sticky=tk.W, pady=3)

        ttk.Label(upload_frame, text="Password:").grid(row=4, column=0, sticky=tk.W)
        self.pass_entry = ttk.Entry(upload_frame, textvariable=self.scp_password, width=25, show="*", state="readonly")
        self.pass_entry.grid(row=4, column=1, sticky=tk.W, pady=3)

        ttk.Label(upload_frame, text="Target Path:").grid(row=5, column=0, sticky=tk.W)
        ttk.Entry(upload_frame, textvariable=self.scp_target_path, width=50).grid(row=5, column=1, sticky=tk.EW, pady=3)

        ttk.Button(upload_frame, text="Upload TRAK File", command=self.upload_file, width=20)\
            .grid(row=6, column=1, sticky=tk.E, pady=10)

        ttk.Label(upload_frame,
                  text="Next steps: [K] Optional Modules > [D]atabase Menu > [A]lternate Database Posting > "
                       "[T]rak Delimited Database > [U]pdate Database.",
                  wraplength=500).grid(row=7, column=0, columnspan=2, pady=10)

    def update_site_settings(self, event=None):
        site = self.site_selection.get()
        config = self.sites.get(site, self.sites["Custom"])
        self.jump_host.set(config["jump"])
        self.final_host.set(config["final"])
        self.scp_username.set(config["user"])
        self.scp_password.set(config["password"])
        state = "normal" if site == "Custom" else "readonly"
        self._set_entry_state(state)
        self.log(f"Site set to {site}: jump={self.jump_host.get()}, final={self.final_host.get()}")

    def _set_entry_state(self, state):
        self.jump_entry.config(state=state)
        self.final_entry.config(state=state)
        self.user_entry.config(state=state)
        self.pass_entry.config(state=state)

    # ---------- File Actions ----------
    def browse_input_file(self):
        f = filedialog.askopenfilename(title="Select Input Excel File", filetypes=[("Excel files", "*.xlsx")])
        if f:
            self.input_file_path.set(f)
            self.log(f"Selected input file: {f}")

    def browse_output_dir(self):
        d = filedialog.askdirectory(title="Select Output Directory")
        if d:
            self.output_dir.set(d)
            self.log(f"Selected output directory: {d}")

    def create_new_spreadsheet(self):
        f = filedialog.asksaveasfilename(title="Create New Spreadsheet", defaultextension=".xlsx",
                                         filetypes=[("Excel files", "*.xlsx")])
        if not f:
            return
        if create_new_spreadsheet(f):
            self.input_file_path.set(f)
            open_with_default_app(f)
            self.log(f"Created new spreadsheet: {f}")

    def process_to_excel(self):
        inp = self.input_file_path.get()
        out = self.output_dir.get()
        if not os.path.exists(inp):
            messagebox.showerror("Error", "Input file not found.")
            return
        self.log("Processing input data...")
        df, path = process_input_data(inp, out)
        self.excel_output_path.set(path)
        self.log(f"Created Excel: {path}")
        self.notebook.tab(1, state="normal")
        self.notebook.select(1)

    def open_excel_file(self):
        open_with_default_app(self.excel_output_path.get())

    def process_to_text(self):
        excel = self.excel_output_path.get()
        out = self.output_dir.get()
        if not os.path.exists(excel):
            messagebox.showerror("Error", "Spreadsheet not found.")
            return
        self.log("Generating TRAK file...")
        text_path = generate_delimited_file(excel, out)
        self.text_output_path.set(text_path)
        self.log(f"Created TRAK file: {text_path}")
        self.notebook.tab(2, state="normal")
        self.notebook.select(2)

    def open_text_file(self):
        open_with_default_app(self.text_output_path.get())

    # ---------- Upload ----------
    def upload_file(self):
        try:
            import paramiko
        except ImportError:
            if messagebox.askyesno("Missing Package", "Paramiko not installed. Install now?"):
                subprocess.check_call([sys.executable, "-m", "pip", "install", "paramiko"])
                import paramiko
            else:
                return

        local_path = self.text_output_path.get()
        if not os.path.exists(local_path):
            messagebox.showerror("Error", f"File not found: {local_path}")
            return

        jump_host = self.jump_host.get().strip()
        final_host = self.final_host.get().strip()
        username = self.scp_username.get().strip()
        password = self.scp_password.get().strip()
        target_path = self.scp_target_path.get().strip()

        from paramiko import SSHClient, AutoAddPolicy

        try:
            if jump_host:
                # --- Connect via jump host ---
                self.log(f"Connecting to jump host {jump_host}...")
                jump = SSHClient()
                jump.set_missing_host_key_policy(AutoAddPolicy())
                jump.connect(jump_host, username=username, password=password)
                self.log(f"Connected to jump host {jump_host}")

                trans = jump.get_transport()
                chan = trans.open_channel("direct-tcpip", (final_host, 22), ("127.0.0.1", 0))

                final = SSHClient()
                final.set_missing_host_key_policy(AutoAddPolicy())
                final.connect(final_host, username=username, password=password, sock=chan)
                self.log(f"Connected to final host {final_host} via {jump_host}")
            else:
                # --- Direct connection ---
                self.log(f"Connecting directly to {final_host}...")
                final = SSHClient()
                final.set_missing_host_key_policy(AutoAddPolicy())
                final.connect(final_host, username=username, password=password)
                self.log(f"Connected directly to {final_host}")

            # --- Upload file ---
            with final.open_sftp() as sftp:
                self.log(f"Uploading {os.path.basename(local_path)} to {target_path}...")
                sftp.put(local_path, target_path)

            messagebox.showinfo("Success", f"Uploaded to {final_host}:{target_path}")
            self.log("✅ Upload complete.")

        except Exception as e:
            self.log(f"Upload error: {e}")
            messagebox.showerror("Upload Error", str(e))
        finally:
            try:
                final.close()
            except Exception:
                pass
            try:
                if jump_host:
                    jump.close()
            except Exception:
                pass

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.update()


if __name__ == "__main__":
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()
