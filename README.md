# pm_to_pdf_maker

This repository contains a minimal script for converting each row of the
`ENTREGAS` sheet from the provided Excel workbook into a separate PDF file.
The script uses only the Python standard library.

## Usage

Run the script from the project directory:

```bash
python3 src/excel_to_pdf.py
```

The Excel workbook must be located at `data/PROJECT MANAGEMENT.xlsx` and
PDF files will be created in `pdf_output/`.
