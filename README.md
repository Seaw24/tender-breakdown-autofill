# Tender Breakdown Automation Tool

## Overview

This tool automates the process of entering daily sales data into the Master Tender Breakdown Excel file.

It works by:
* Scanning the cash sheets folder for daily Excel reports
* Identifying the location based on the filename (e.g., `beanyard`, `hecs`)
* Opening each report and finding specific data rows (e.g., `TOTALS`, `Hive Tavlo`)
* Copying the sales data (Sales, Flex, UCash, etc.) into the matching column and date row in `Tender Breakdown.xlsx`

## Project Structure

Ensure your folder structure looks like this:

```
/Chartwells_Automation
│
├── run_autofill.bat        # [EXECUTABLE] Double-click this to run the tool
├── Tender Breakdown.xlsx   # [MASTER] Main Excel file
├── requirements.txt        # [SYSTEM] Python dependencies
│
├── cash sheets/            # [INPUT] Daily downloaded reports
│   ├── Beanyard 11-13.xlsx
│   ├── HECS 11-13.xlsx
│   └── ...
│
├── src/                    # [CODE] Source files
│   ├── autofill.py         # Main automation script
│   └── config.py           # Configuration & mappings
│
└── myenv/                  # [SYSTEM] Python virtual environment
```

## Installation & Setup

**Only required when setting up on a new computer**

### Install Python

Ensure Python 3.x is installed.

### Create Virtual Environment

Open a terminal in the project folder and run:

```bash
python -m venv myenv
```

### Install Dependencies

Activate the environment and install required libraries.

**Windows:**

```bash
myenv\Scripts\activate
pip install openpyxl python-dateutil
```

## Configuration (`src/config.py`)

This file controls how the script matches cash sheet files to master columns.

**You must edit this file if:**
* A new location opens
* A filename changes (e.g., `HECS` → `Hive`)
* A row label changes in the cash sheet (e.g., `TOTALS` → `Grand Total`)

### How to Edit Mappings

Open `src/config.py` in a text editor and locate:

```python
FILENAME_TO_MASTER_NAME
```

**Mapping Structure:**

```python
"filename_keyword": [
    {"Master Column Name": "Row Label in Cash Sheet"}
]
```

**Example:**

If you have a file named `HECS Daily.xlsx` containing:
* Hive data labeled "Hive Tavlo"
* Quartzdyne data labeled "Quartz Café"

```python
"hecs": [
    {"hive": "Hive Tavlo"},
    {"quartzdyne": "Quartz Café"}
]
```

## Troubleshooting

| Error | Cause | Solution |
|-------|-------|----------|
| PermissionError | Tender Breakdown.xlsx is open | Close the Excel file and try again |
| Date not found | `"Date:"` not found in rows 1–5 | Verify the cash sheet format |
| Keyword not found | Row label missing in Column C | Update row label in `config.py` |
| No mapping found | Filename doesn't match config keyword | Rename the file to match config (e.g., `Beanyard.xlsx`) |
