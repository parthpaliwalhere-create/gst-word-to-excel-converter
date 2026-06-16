# GST Word to Excel Converter Demo

A Python desktop tool that extracts GST invoice data from Word/PDF files and exports accountant-style Excel GST registers.

> Public-safe demo version. Real business details have been removed.

## Features

- Batch reads Word/PDF invoice files
- Extracts invoice number, date, party, GSTIN, taxable value, CGST, SGST, IGST and total
- Sorts bills by bill number
- Generates Excel GST register
- Creates clean output for review/upload
- Supports validation and error/check sheets

## Tech Stack

- Python
- Tkinter
- python-docx
- OpenPyXL
- pdfplumber

## Run Locally

```powershell
pip install -r requirements.txt
python app.py
```

## Suggested GitHub Release

Upload the source code in the repository and put the runnable ZIP/EXE inside **GitHub Releases**, not inside the normal source-code area.

## Note

This project was developed as a learning project with AI assistance.
