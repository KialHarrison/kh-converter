import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from docx2pdf import convert
import os
import requests

GITHUB_API_URL = "https://api.github.com/repos/KialHarrison/kh-converter/releases/latest"
CURRENT_VERSION = "1.0.0"

class DocxToPdfConverter(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("DOCX to PDF Converter")
        self.geometry("600x200")
        tabControl = ttk.Notebook(self)

        self.input_dir = ""
        self.output_dir = ""

        self.create_widgets()
        # self.check_for_updates()

    def create_widgets(self):
        tabControl = ttk.Notebook(self)
        
        tab1 = ttk.Frame(tabControl)
        tab2 = ttk.Frame(tabControl)
        
        tabControl.add(tab1, text='DOCX to PDF')
        tabControl.add(tab2, text='Testing')

        tabControl.grid(column=3, row=3)
        
        ttk.Label(tab1, text="Input Directory:").grid(row=1, column=1, padx=10, pady=10)
        self.input_entry = ttk.Entry(tab1, width=30)
        self.input_entry.grid(row=1, column=2, padx=10, pady=10)

        ttk.Button(tab1, text="Browse", command=self.browse_input).grid(row=1, column=3, padx=10, pady=10)

        ttk.Label(tab1, text="Output Directory:").grid(row=2, column=1, padx=10, pady=10)
        self.output_entry = ttk.Entry(tab1, width=30)
        self.output_entry.grid(row=2, column=2, padx=10, pady=10)

        ttk.Button(tab1, text="Browse", command=self.browse_output).grid(row=2, column=3, padx=10, pady=10)

        ttk.Button(tab1, text="Convert", command=self.convert_files).grid(row=3, column=2, padx=10, pady=10)

    def browse_input(self):
        self.input_dir = filedialog.askdirectory()
        self.input_entry.delete(0, tk.END)
        self.input_entry.insert(0, self.input_dir)

    def browse_output(self):
        self.output_dir = filedialog.askdirectory()
        self.output_entry.delete(0, tk.END)
        self.output_entry.insert(0, self.output_dir)

    def convert_files(self):
        if not self.input_dir or not self.output_dir:
            messagebox.showwarning("Missing Information", "Please select both input and output directories.")
            return

        if not os.path.isdir(self.input_dir):
            messagebox.showerror("Invalid Directory", "The specified input directory does not exist.")
            return

        if not os.path.isdir(self.output_dir):
            messagebox.showerror("Invalid Directory", "The specified output directory does not exist.")
            return

        # Check if Microsoft Word or LibreOffice is accessible
        try:
            # Convert all .docx files in the input directory to PDF and save them in the output directory
            convert(self.input_dir, self.output_dir)
            messagebox.showinfo("Success", "All .docx files have been converted to PDF.")
        except Exception as e:
            messagebox.showerror("Conversion Error", f"An error occurred during conversion: {e}")
            print(f"Error details: {e}")
            
    def check_for_updates(self):
        try:
            response = requests.get(GITHUB_API_URL)
            response.raise_for_status()
            latest_release = response.json()
            latest_version = latest_release["tag_name"]

            if self.is_newer_version(latest_version, CURRENT_VERSION):
                messagebox.showinfo("Update Available", f"A new version ({latest_version}) is available. Please update your application.")
            else:
                print("You are using the latest version.")
        except requests.RequestException as e:
            print(f"Error checking for updates: {e}")

    def is_newer_version(self, latest, current):
        latest_parts = [int(part) for part in latest.split('.')]
        current_parts = [int(part) for part in current.split('.')]
        return latest_parts > current_parts

if __name__ == "__main__":
    app = DocxToPdfConverter()
    app.mainloop()
