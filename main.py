import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def pdf_to_csv(pdf_path, csv_path, page_range=None):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            tables = []
            start_page = page_range[0] if page_range else 1
            end_page = page_range[1] if page_range else len(pdf.pages)
            
            for page_num in range(start_page - 1, end_page):
                page = pdf.pages[page_num]
                extracted_tables = page.extract_tables()
                
                for table_num, table in enumerate(extracted_tables, start=1):
                    df = pd.DataFrame(table[1:], columns=table[0])
                    tables.append(df)
                    print(f"Extracted table from page {page_num + 1}, table {table_num}")
            
            final_df = pd.concat(tables, ignore_index=True)
            final_df.to_csv(csv_path, index=False)
            messagebox.showinfo("Success", f"Tables saved to {csv_path}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def browse_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    pdf_path_entry.delete(0, tk.END)
    pdf_path_entry.insert(0, file_path)

def save_csv():
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    csv_path_entry.delete(0, tk.END)
    csv_path_entry.insert(0, file_path)

def run_extraction():
    pdf_path = pdf_path_entry.get()
    csv_path = csv_path_entry.get()
    
    try:
        start_page = int(start_page_entry.get())
        end_page = int(end_page_entry.get())
        page_range = (start_page, end_page)
    except ValueError:
        messagebox.showerror("Error", "Please enter valid page numbers.")
        return
    
    if not pdf_path or not csv_path:
        messagebox.showerror("Error", "Please select both PDF and CSV file paths.")
        return
    
    pdf_to_csv(pdf_path, csv_path, page_range)

# GUI setup
root = tk.Tk()
root.title("PDF to CSV Converter")

# PDF file selection
tk.Label(root, text="PDF File:").grid(row=0, column=0, padx=10, pady=5)
pdf_path_entry = tk.Entry(root, width=40)
pdf_path_entry.grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=browse_pdf).grid(row=0, column=2, padx=10, pady=5)

# CSV file save path
tk.Label(root, text="Save as CSV:").grid(row=1, column=0, padx=10, pady=5)
csv_path_entry = tk.Entry(root, width=40)
csv_path_entry.grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="Save As", command=save_csv).grid(row=1, column=2, padx=10, pady=5)

# Page range input
tk.Label(root, text="Page Range (Start - End):").grid(row=2, column=0, padx=10, pady=5)
start_page_entry = tk.Entry(root, width=5)
start_page_entry.grid(row=2, column=1, sticky="W", padx=10, pady=5)
tk.Label(root, text="to").grid(row=2, column=1)
end_page_entry = tk.Entry(root, width=5)
end_page_entry.grid(row=2, column=1, sticky="E", padx=10, pady=5)

# Run button
tk.Button(root, text="Convert to CSV", command=run_extraction, bg="lightblue").grid(row=3, column=0, columnspan=3, pady=20)

root.mainloop()
