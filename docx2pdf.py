import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
import os

class DocxToPdfConverter(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("DOCX to PDF Converter")
        self.geometry("600x200")

        self.input_dir = ""
        self.output_dir = ""

        self.create_widgets()

    def create_widgets(self):
        tk.Label(self, text="Input Directory:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.input_entry = tk.Entry(self, width=30)
        self.input_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Button(self, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=10, pady=10)

        tk.Label(self, text="Output Directory:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.output_entry = tk.Entry(self, width=30)
        self.output_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Button(self, text="Browse", command=self.browse_output).grid(row=1, column=2, padx=10, pady=10)

        tk.Button(self, text="Convert", command=self.convert_files).grid(row=2, column=0, columnspan=3, pady=20)

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

if __name__ == "__main__":
    app = DocxToPdfConverter()
    app.mainloop()
