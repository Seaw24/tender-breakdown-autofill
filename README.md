Tender Breakdown Automation

1. Overview

This tool automates the transfer of financial data from daily Cash Sheet Reports into the master Tender Breakdown Excel file.

It reads specific rows (e.g., "TOTALS") from the daily reports and maps them to the correct Location and Date columns in the master file.

2. Project Structure

/Project_Folder
│
├── run_autofill.bat # <--- DOUBLE CLICK THIS TO RUN
├── Tender Breakdown.xlsx # The Master Excel File
│
├── cash sheets/ # FOLDER: Place downloaded daily reports here
│ ├── Beanyard 11-13.xlsx
│ └── ...
│
├── src/ # Source Code
│ ├── autofill.py # Main logic script
│ └── config.py # Configuration & Column Mappings
│
└── myenv/ # Python Virtual Environment

3. Configuration (src/config.py)

If a location name changes, a new location opens, or filenames change, you must update src/config.py.

How Mappings Work

The system uses the Filename to determine which Location column to fill.

Example in config.py:

FILENAME_TO_MASTER_NAME = { # "keyword in filename": [{"Master Column Name": "Row Name in Cash Sheet"}]

    "crimson corner": [
        {"Crimson Corner": "TOTALS"}, # Fills 'Crimson Corner' col using 'TOTALS' row
        {"Gardner": "Thirst"}         # Fills 'Gardner' col using 'Thirst' row (same file)
    ]

}

Key ("crimson corner"): The script looks for a file containing this text (case-insensitive).

Master Column Name ("Crimson Corner"): Must match the header in Tender Breakdown.xlsx exactly.

Row Name ("TOTALS"): The script searches the cash sheet for a row where Column C (or B) contains this text.

4. Technical Requirements

Python 3.x

Dependencies: openpyxl, python-dateutil

Virtual Environment: The run_autofill.bat expects a virtual environment named myenv.

5. Troubleshooting

"Permission Denied": Close Tender Breakdown.xlsx before running the script.

"Date not found": Ensure the cash sheet has a cell containing "Date:" in the first 5 rows.

"Location not found": Check config.py to ensure the filename keyword matches the file you downloaded.
