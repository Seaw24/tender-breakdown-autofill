from openpyxl import load_workbook
from pathlib import Path
from dateutil.parser import parse
import os
from config import FILENAME_TO_MASTER_NAME, BREAKDOWN_LOC_MAP, DIRECTORY_PATHS


def process_tender_breakdown(casheet_path, master_path):
    # Load Workbooks
    cs_wb = load_workbook(casheet_path, data_only=True)
    master_wb = load_workbook(master_path)
    master_ws = master_wb.active

    # 1. Identify Location from Filename
    # Extracts "Beanyard" from "cash sheets\Beanyard 11-13-25.xlsx"
    # FOR IMPROVMENT USING THE LOCATION ID IN THE SHEET
    fname = casheet_path.stem  # Gets "Beanyard 11-13-25"
    master_location = None

    for key, official_name in FILENAME_TO_MASTER_NAME.items():
        if key.lower() in fname.lower():
            master_location = official_name
            break

    if not master_location:
        print(f"❌ Could not map filename '{fname}' to a Master Location.")
        print("==================================================\n")
        print("==================================================\n")
        return

    print(f"📂 Processing: {fname} -> Master Location: {master_location}")

    # 2. Process each Sheet (Day) in the Casheet
    for ws in cs_wb.worksheets:
        # Find Date in Casheet
        report_date = None
        for c in range(1, 30):
            if str(ws.cell(1, c).value).upper().strip() == "DATE:":
                date_val = ws.cell(1, c+1).value
                if date_val:
                    # Standardize date to object for comparison
                    # CHECK IF FORMAT IS CORRECT
                    report_date = parse(str(date_val)).date()
                    break

        if not report_date:  # Need to return and inform user
            continue

        # 3. Find TOTALS row in Casheet
        totals_row = None
        for r in range(45, 60):
            # Better way to search bottom up/ how to not consider extra rows
            if str(ws.cell(r, 2).value).strip().upper() == "TOTALS":
                totals_row = r
                break

        if not totals_row:
            continue

        # 3.5 Print Found Info
        print(f"  🗓 Found Date: {report_date} | Location: {master_location}")
        print("Extracting data...")

        # 4. Extract Data from Casheet (Based on your offsets)
        # Order: Total Sales, Contract, Flex, Transfer, Coupons, Ucash, Dining, Amex, Disc, MC, Visa
        imp_indexes = [3, 4, 6, 7, 8, 11, 13, 16, 17, 18, 19]
        daily_values = [ws.cell(totals_row, i).value for i in imp_indexes]

        print("Filling data into Master Breakdown...")
        # 5. Find Target Row in Master Breakdown (Search Column A)
        # FOR IMPROVEMENT - USING OTHER SEARCHING ALGO
        target_row = None
        for r in range(3, master_ws.max_row + 1):
            master_date_val = master_ws.cell(r, 1).value
            if master_date_val:
                # CHECK IF FORMAT IS CORRECT
                master_date = parse(str(master_date_val)).date()
                if master_date == report_date:
                    target_row = r
                    break

        # 6. Fill Data into Master
        if target_row and master_location in BREAKDOWN_LOC_MAP:
            start_col = BREAKDOWN_LOC_MAP[master_location]
            for i, val in enumerate(daily_values):
                master_ws.cell(target_row, start_col + i).value = val
            print(f"   ✅ Filled {ws.title} ({report_date})")

        print(f"Finished day: {report_date}\n")
        print("--------------------------------------------------\n")

    # Save Master Workbook
    master_wb.save(master_path)
    print(f"💾 Master Breakdown Saved!\n")
    print("==================================================\n")
    print("==================================================\n")


# Main Execution
if __name__ == "__main__":
    # Set Paths
    casheet_dir = Path(DIRECTORY_PATHS["casheets_dir"])
    # NEED TO CHECK IF FILE EXISTS
    master_file = DIRECTORY_PATHS["master_file"]
    casheet_files_paths = [
        # IF NONE Handling
        casheet_dir / f for f in os.listdir(casheet_dir) if f.endswith('.xlsx')]

    # Fill Master Breakdown
    for casheet_path in casheet_files_paths:
        process_tender_breakdown(casheet_path, master_file)
