# PDF Email Processor and Excel Updater

This project automates the process of downloading PDF attachments from Outlook emails and updating corresponding Excel records. It's particularly designed for handling e.g. documents and purchase orders.

## Features

- Downloads PDF attachments from Outlook emails based on:
  - Subject filters
  - Date range (default: last 7 days)
- Extracts purchase order (PO) information from PDFs including:
  - PO Number
  - Order Date
  - Vendor details
  - Amount
- Updates Excel spreadsheets with extracted PO data
- Matches PDF content with Excel records using intelligent text comparison
- Preserves Excel formatting while updating cells

## Prerequisites

```python
pip install win32com.client
pip install pdfplumber
pip install pandas
pip install openpyxl
```

## Project Structure

- `outlook.py` - Handles email processing and PDF downloads
- `main.py` - Processes PDFs and updates Excel files
- `PO_in_excel/` - Directory for processed PDFs
- `PO_not_in_excel/` - Directory for downloaded PDFs

## Usage

1. Set up the required directories:
   ```
   PO_in_excel/
   PO_not_in_excel/
   ```

2. Run the Outlook processor to download PDFs:
   ```bash
   python outlook.py
   ```
   This will:
   - Connect to your Outlook
   - Filter emails with "Hotel booking" in subject
   - Download PDF attachments from the last 7 days

3. Process PDFs and update Excel:
   ```bash
   python main.py
   ```
   This will:
   - Read PDFs from the `PO_in_excel/` directory
   - Extract PO information
   - Update matching records in the Excel file

## Configuration

You can modify the following parameters in the code:

In `outlook.py`:
```python
subject_filter = "Hotel booking"  # Change email subject filter
days = 7    # Change number of days to look back
```

## Text Processing Features

- Text cleaning and normalization
- Intelligent matching using sequence comparison
- Handles multi-page PDFs
- Skips already processed records
- Preserves Excel formatting

## Error Handling

- Robust error handling for PDF processing
- Skips problematic files and continues processing
- Detailed logging of operations and errors

## Notes

- The system uses fuzzy matching to find corresponding Excel records
- Only updates Excel cells if they are empty
- Processes multiple PDFs in batch
- Maintains a log of processed files and errors

## Support

For any issues or questions, please submit an issue in the repository.