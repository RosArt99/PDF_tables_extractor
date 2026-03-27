from pathlib import Path
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import re

def clean_table(df):
    
    remove_patterns = [
        "Boeing Proprietary", #in that case programm was used on Boeing docs, but you can make your patterns as needed
        "Page"
    ]

    cleaned_rows = []

    for row in df.values.tolist():
        row_text = " ".join([str(x) for x in row if x])

        if not row_text.strip():
            continue

        if any(p in row_text for p in remove_patterns):
            continue

        cleaned_rows.append(row)

    return pd.DataFrame(cleaned_rows)

def select_pdf():
    file_path = filedialog.askopenfilename(
        filetypes=[("PDF files", "*.pdf")]
    )
    if file_path:
        pdf_path_var.set(file_path)


def parse_pages(pages_text):

    if "-" in pages_text:
        start, end = pages_text.split("-")
        pages = list(range(int(start) - 1, int(end)))
    else:
        pages = [int(p.strip()) - 1 for p in pages_text.split(",")]

    return pages


def extract_tables():

    pdf_path = pdf_path_var.get()
    pages_input = pages_var.get()

    if not pdf_path:
        messagebox.showerror("Error", "Please select a PDF file")
        return

    if not pages_input:
        messagebox.showerror("Error", "Please enter page numbers")
        return

    try:
        pages = parse_pages(pages_input)

        tables_all = []

        with pdfplumber.open(pdf_path) as pdf:

            for page_number in pages:

                if page_number >= len(pdf.pages):
                    messagebox.showerror(
                        "Error",
                        f"Page {page_number+1} does not exist in PDF"
                    )
                    return

                page = pdf.pages[page_number]
                tables = page.extract_tables()

                for table in tables:
                    df = pd.DataFrame(table)
                    tables_all.append(df)

        if not tables_all:
            messagebox.showinfo("Result", "No tables found")
            return

        df_final = pd.concat(tables_all, ignore_index=True)
        df_final = clean_table(df_final)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")]
        )

        if save_path:

            with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:

                df_final.to_excel(writer, index=False)

                workbook = writer.book
                worksheet = writer.sheets["Sheet1"]

                format_text = workbook.add_format({'num_format': '@'})

                worksheet.set_column(0, len(df_final.columns), None, format_text)

            messagebox.showinfo("Success", "Tables extracted successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# ---------------- GUI ---------------- #

root = tk.Tk()
root.title("PDF Table Extractor")
root.geometry("450x220")

pdf_path_var = tk.StringVar()
pages_var = tk.StringVar()

# PDF selection
tk.Label(root, text="PDF File:").pack(pady=5)
tk.Entry(root, textvariable=pdf_path_var, width=50).pack()
tk.Button(root, text="Browse", command=select_pdf).pack(pady=5)

# Pages input
tk.Label(root, text="Pages (example: 33-34 or 33,34):").pack(pady=5)
tk.Entry(root, textvariable=pages_var, width=20).pack()

# Extract button
tk.Button(root, text="Extract Tables", command=extract_tables).pack(pady=15)

root.mainloop()