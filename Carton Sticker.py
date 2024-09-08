import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import serial
import serial.tools.list_ports
from barcode import EAN13
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch
import os

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Carton Sticker Printer")
        self.style = ttk.Style()
        self.configure_styles()
        self.geometry("1000x800")
        self.resizable(True, True)

        # Define StringVars
        self.printer_ip = tk.StringVar(value="10.1.10.110")
        self.weight_scale_port = tk.StringVar(value="Select Port")
        self.so_number = tk.StringVar(value="")
        self.job_number = tk.StringVar(value="")
        self.rbo = tk.StringVar(value="")
        self.weight_scale_weight = tk.StringVar(value="")
        self.item = tk.StringVar(value="")
        self.order_qty = tk.StringVar(value="1")  # Default order quantity to 1
        self.po_number = tk.StringVar(value="")
        self.customer = tk.StringVar(value="")
        self.print_to_pdf = tk.BooleanVar(value=False)
        self.show_total_weight = tk.BooleanVar(value=False)
        self.auto_weight = tk.BooleanVar(value=False)
        self.pdf_folder = tk.StringVar(value="")
        self.excel_data = None  # DataFrame to store Excel data

        self.create_widgets()

    def configure_styles(self):
        self.style.theme_use("clam")  # Use a modern theme

        # Configure dark theme
        self.configure(bg='#2e2e2e')  # Dark background color
        self.style.configure('TFrame', background='#2e2e2e')
        self.style.configure('TLabel', background='#2e2e2e', foreground='#e3e3e3')
        self.style.configure('TEntry', background='#3e3e3e', foreground='#e3e3e3', fieldbackground='#3e3e3e')
        self.style.configure('TButton', background='#4e4e4e', foreground='#e3e3e3', borderwidth=1)
        self.style.map('TButton', background=[('pressed', '#5e5e5e'), ('active', '#6e6e6e')])
        self.style.configure('TCheckbutton', background='#2e2e2e', foreground='#e3e3e3')

    def load_excel_data(self, file_path):
        try:
            self.excel_data = pd.read_excel(file_path, engine='openpyxl')
            messagebox.showinfo("Info", "Excel data loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while loading the Excel file: {e}")

    def on_so_number_change(self, *args):
        try:
            so_number = self.so_number.get()
            if self.excel_data is not None and so_number:
                data_row = self.excel_data.loc[self.excel_data['SO Number'] == so_number]
                if not data_row.empty:
                    data_row = data_row.iloc[0]
                    self.job_number.set(str(data_row['Job Number']))
                    self.rbo.set(str(data_row['RBO']))
                    self.weight_scale_weight.set(str(data_row['Weight']))
                    self.item.set(str(data_row['Item']))
                    self.order_qty.set(str(data_row['Order Qty']))
                    self.po_number.set(str(data_row['PO Number']))
                    self.customer.set(str(data_row['Customer']))
                else:
                    messagebox.showinfo("Info", f"SO Number '{so_number}' not found in the Excel data.")
            else:
                messagebox.showinfo("Info", "Please load the Excel data first.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
        
    def create_widgets(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill='both', expand=True)

        settings_frame = ttk.Frame(main_frame, padding="10")
        settings_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Configure grid layout
        settings_frame.grid_rowconfigure(0, weight=1)
        settings_frame.grid_columnconfigure(0, weight=1)
        settings_frame.grid_columnconfigure(1, weight=1)

        # Add Widgets with grid layout
        ttk.Label(settings_frame, text="Printer IP:", font=("Arial", 12)).grid(row=0, column=0, sticky='w', pady=5)
        ttk.Entry(settings_frame, textvariable=self.printer_ip, font=("Arial", 12)).grid(row=0, column=1, sticky='ew', pady=5)

        ttk.Label(settings_frame, text="Weight Scale Port:", font=("Arial", 12)).grid(row=1, column=0, sticky='w', pady=5)
        port_frame = ttk.Frame(settings_frame)
        port_frame.grid(row=1, column=1, sticky='ew', pady=5)
        ttk.Entry(port_frame, textvariable=self.weight_scale_port, font=("Arial", 12)).pack(side=tk.LEFT, fill='x', expand=True)
        ttk.Button(port_frame, text="Select Port", command=self.select_port, style="TButton").pack(side=tk.RIGHT)

        ttk.Label(settings_frame, text="SO Number:", font=("Arial", 12)).grid(row=2, column=0, sticky='w', pady=5)
        so_number_entry = ttk.Entry(settings_frame, textvariable=self.so_number, font=("Arial", 12))
        so_number_entry.grid(row=2, column=1, sticky='ew', pady=5)
        self.so_number.trace_add("write", self.on_so_number_change)

        ttk.Label(settings_frame, text="Job Number:", font=("Arial", 12)).grid(row=3, column=0, sticky='w', pady=5)
        ttk.Entry(settings_frame, textvariable=self.job_number, font=("Arial", 12)).grid(row=3, column=1, sticky='ew', pady=5)

        ttk.Label(settings_frame, text="RBO:", font=("Arial", 12)).grid(row=4, column=0, sticky='w', pady=5)
        ttk.Entry(settings_frame, textvariable=self.rbo, font=("Arial", 12)).grid(row=4, column=1, sticky='ew', pady=5)

        ttk.Label(settings_frame, text="Weight:", font=("Arial", 12)).grid(row=5, column=0, sticky='w', pady=5)
        ttk.Entry(settings_frame, textvariable=self.weight_scale_weight, font=("Arial", 12)).grid(row=5, column=1, sticky='ew', pady=5)

        ttk.Label(settings_frame, text="Item:", font=("Arial", 12)).grid(row=6, column=0, sticky='w', pady=5)
        ttk.Entry(settings_frame, textvariable=self.item, font=("Arial", 12)).grid(row=6, column=1, sticky='ew', pady=5)

        ttk.Label(settings_frame, text="Order Qty:", font=("Arial", 12)).grid(row=7, column=0, sticky='w', pady=5)
        ttk.Entry(settings_frame, textvariable=self.order_qty, font=("Arial", 12)).grid(row=7, column=1, sticky='ew', pady=5)

        ttk.Label(settings_frame, text="PO Number:", font=("Arial", 12)).grid(row=8, column=0, sticky='w', pady=5)
        ttk.Entry(settings_frame, textvariable=self.po_number, font=("Arial", 12)).grid(row=8, column=1, sticky='ew', pady=5)

        ttk.Label(settings_frame, text="Customer:", font=("Arial", 12)).grid(row=9, column=0, sticky='w', pady=5)
        ttk.Entry(settings_frame, textvariable=self.customer, font=("Arial", 12)).grid(row=9, column=1, sticky='ew', pady=5)

        print_options_frame = ttk.Frame(settings_frame)
        print_options_frame.grid(row=10, column=0, columnspan=2, pady=5, sticky='ew')

        ttk.Label(print_options_frame, text="Print Options:", font=("Arial", 12)).pack(anchor='w', pady=5)

        ttk.Checkbutton(print_options_frame, text="Print to PDF", variable=self.print_to_pdf, command=self.get_pdf_folder, style="TCheckbutton").pack(anchor='w', pady=5)
        ttk.Checkbutton(print_options_frame, text="Show Total Weight", variable=self.show_total_weight, style="TCheckbutton").pack(anchor='w', pady=5)
        ttk.Checkbutton(print_options_frame, text="Auto Weight", variable=self.auto_weight, style="TCheckbutton").pack(anchor='w', pady=5)

        # Button to load Excel file
        ttk.Button(settings_frame, text="Load Excel Data", command=self.load_excel_file, style="TButton").grid(row=11, column=0, sticky='sw', pady=10, padx=10)

        # Print Button at the bottom right
        ttk.Button(settings_frame, text="Print", command=self.print_preview, style="TButton").grid(row=11, column=1, sticky='se', pady=10, padx=10)

        preview_frame = ttk.Frame(main_frame, padding="10")
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.preview_text = tk.Text(preview_frame, wrap='word', font=("Arial", 12), bg='#2e2e2e', fg='#e3e3e3', insertbackground='white')
        self.preview_text.pack(fill=tk.BOTH, expand=True)

        self.preview_image = ttk.Label(preview_frame)
        self.preview_image.pack(fill=tk.BOTH, expand=True)

    def load_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.load_excel_data(file_path)

    def select_port(self):
        ports = serial.tools.list_ports.comports()
        port_names = [port.device for port in ports]
        port_name = tk.simpledialog.askstring("Select Port", "Available ports:\n" + "\n".join(port_names))
        if port_name in port_names:
            self.weight_scale_port.set(port_name)
        else:
            messagebox.showerror("Error", "Invalid port selected.")

    def get_pdf_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.pdf_folder.set(folder_path)

    def print_preview(self):
        # Dummy function for print preview (To be implemented)
        preview_text = self.preview_text.get("1.0", tk.END)
        messagebox.showinfo("Print Preview", preview_text)

if __name__ == "__main__":
    app = Application()
    app.mainloop()
