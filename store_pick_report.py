# ============================================================
# STORE WISE PICK / INTERNAL TRANSFER REPORT
# ============================================================

import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side
)


# ============================================================
# MAIN FUNCTION
# ============================================================

def generate_store_pick_report(file_path, output_file):

    # ========================================================
    # 1. READ ALL SHEETS
    #
    # HEADER IS ON EXCEL ROW 5
    # ========================================================

    all_sheets = pd.read_excel(
        file_path,
        sheet_name=None,
        header=5
    )


    # ========================================================
    # 2. SHEETS TO EXCLUDE
    # ========================================================

    EXCLUDED_SHEETS = [
        "SPM_Virtual_Store",
        "Export(SPM)",
        "Virtual_Store_AD"
    ]


    # ========================================================
    # 3. READ ALL STORE SHEETS
    # ========================================================

    all_data = []

    for sheet_name, df in all_sheets.items():

        # ----------------------------------------------------
        # SKIP EXCLUDED SHEETS
        # ----------------------------------------------------

        if sheet_name in EXCLUDED_SHEETS:

            print(f"SKIPPED: {sheet_name}")

            continue


        print(f"Reading: {sheet_name}")


        # ----------------------------------------------------
        # REMOVE COMPLETELY EMPTY ROWS
        # ----------------------------------------------------

        df = df.dropna(
            how="all"
        ).copy()


        # ----------------------------------------------------
        # REMOVE COMPLETELY EMPTY COLUMNS
        # ----------------------------------------------------

        df = df.dropna(
            axis=1,
            how="all"
        )


        # ----------------------------------------------------
        # CLEAN COLUMN NAMES
        # ----------------------------------------------------

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )
