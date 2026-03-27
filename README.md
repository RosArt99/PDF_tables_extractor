# PDF Table Extractor

## 📄 Description

This is a simple desktop application for extracting tables from PDF files and saving them into an Excel (.xlsx) file.

The tool uses:

* `pdfplumber` for reading PDF files
* `pandas` for data processing
* `tkinter` for a GUI

It also includes basic cleaning logic to remove unwanted rows such as headers, footers, and proprietary notes.

---

## 🚀 Features

* Select any PDF file via GUI
* Extract tables from specific pages (e.g. `33-34` or `33,34`)
<img width="464" height="259" alt="image" src="https://github.com/user-attachments/assets/74aaab08-bfcb-491b-89ef-1d9682f4dbb0" />

* Automatically merge all extracted tables
* Remove:
  * Empty rows
  * Service text (adjustable in "clean_table function")
* Export cleaned data to Excel
* Force Excel cells to be stored as text (avoids formatting issues)

---

## ▶️ Usage

Just use .exe or .py if additional settings required

### Steps:

1. Click **Browse** and select a PDF file
2. Enter page numbers:

   * Range: `33-34`
   * Multiple pages: `33,34`
3. Click **Extract Tables**
4. Choose where to save the Excel file

---

## 👤 Author

Rostyslav Artiukh

---
