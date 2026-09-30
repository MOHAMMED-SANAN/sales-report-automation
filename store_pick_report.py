# ============================================================
# STORE WISE PICK / INTERNAL TRANSFER REPORT
# FULL COMBINED + INDIVIDUAL STORE SHEETS
# ============================================================

import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# HELPER FUNCTION
# ============================================================

def clean_sheet_name(name):
    """
    Excel sheet names:
    - Maximum 31 characters
    - Cannot contain: / \ ? * [ ]
    """
    invalid_chars = ['/', '\\', '?', '*', '[', ']']

    name = str(name)

    for char in invalid_chars:
        name = name.replace(char, '_')

    name = name.strip()

    if not name:
        name = "Store"

    return name[:31]


def make_report_data(filtered_data):
    """
    Create Store + Person + Operation Type report.
    """

    if filtered_data.empty:
        return pd.DataFrame(
            columns=[
                "STORE",
                "Last Updated By",
                "Operation Type",
                "Distinct Count of Source Document",
                "Sum of Demand Qty",
                "ROW_TYPE"
            ]
        )

    # ========================================================
    # PIVOT
    # ========================================================

    pivot_data = (
        filtered_data
        .groupby(
            [
                "STORE",
                "Last Updated By",
                "Operation Type"
            ],
            dropna=False
        )
        .agg(
            **{
                "Distinct Count of Source Document":
                    ("Source Document", "nunique"),

                "Sum of Demand Qty":
                    ("Demand Qty", "sum")
            }
        )
        .reset_index()
    )

    # ========================================================
    # OPERATION ORDER
    # Internal Transfers first
    # Pick second
    # ========================================================

    operation_order = {
        "internal transfers": 0,
        "pick": 1
    }

    pivot_data["_operation_order"] = (
        pivot_data["Operation Type"]
        .astype(str)
        .str.lower()
        .map(operation_order)
        .fillna(99)
    )

    pivot_data = pivot_data.sort_values(
        by=[
            "STORE",
            "Last Updated By",
            "_operation_order",
            "Operation Type"
        ],
        na_position="last"
    )

    pivot_data = pivot_data.drop(
        columns=["_operation_order"]
    )

    # ========================================================
    # FINAL DISPLAY
    # ========================================================

    final_report = []

    for store, store_df in pivot_data.groupby(
        "STORE",
        sort=False
    ):

        first_store_row = True

        for person, person_df in store_df.groupby(
            "Last Updated By",
            sort=False
        ):

            first_person_row = True

            for _, row in person_df.iterrows():

                if first_store_row:
                    display_store = store
                    first_store_row = False
                else:
                    display_store = ""

                if first_person_row:
                    display_person = person
                    first_person_row = False
                else:
                    display_person = ""

                final_report.append({
                    "STORE": display_store,
                    "Last Updated By": display_person,
                    "Operation Type": row["Operation Type"],
                    "Distinct Count of Source Document":
                        row["Distinct Count of Source Document"],
                    "Sum of Demand Qty":
                        row["Sum of Demand Qty"],
                    "ROW_TYPE": "DATA"
                })

        # ====================================================
        # STORE TOTAL
        # ====================================================

        store_data = filtered_data[
            filtered_data["STORE"] == store
        ]

        final_report.append({
            "STORE": f"{store} Total",
            "Last Updated By": "",
            "Operation Type": "",
            "Distinct Count of Source Document":
                store_data["Source Document"].nunique(),
            "Sum of Demand Qty":
                store_data["Demand Qty"].sum(),
            "ROW_TYPE": "TOTAL"
        })

    return pd.DataFrame(final_report)


# ============================================================
# MAIN FUNCTION
# ============================================================

def generate_store_pick_report(file_path, output_file):

    # ========================================================
    # 1. SHEETS TO EXCLUDE
    # ========================================================

    EXCLUDED_SHEETS = [
        "SPM_Virtual_Store",
        "Export(SPM)",
        "Virtual_Store_AD"
    ]

    # ========================================================
    # 2. READ ALL EXCEL SHEETS
    # ========================================================

    all_sheets = pd.read_excel(
        file_path,
        sheet_name=None,
        skiprows=4
    )

    print("\n==============================================")
    print("ALL SHEETS")
    print("==============================================")

    for sheet_name in all_sheets.keys():
        print(" -", sheet_name)

    # ========================================================
    # 3. REQUIRED COLUMNS
    # ========================================================

    required_columns = [
        "Source Document",
        "Operation Type",
        "Demand Qty",
        "State",
        "Last Updated By"
    ]

    # ========================================================
    # 4. STORE-WISE FILTERED DATA
    #
    # Dictionary:
    #
    # {
    #   "Qusais": filtered data,
    #   "Dubai": filtered data,
    #   ...
    # }
    # ========================================================

    store_filtered_data = {}

    # ========================================================
    # 5. PROCESS EACH SHEET INDIVIDUALLY
    # ========================================================

    for sheet_name, df in all_sheets.items():

        # ----------------------------------------------------
        # EXCLUDE SHEETS
        # ----------------------------------------------------

        if sheet_name in EXCLUDED_SHEETS:

            print(f"\nSKIPPED: {sheet_name}")

            continue

        print(f"\nProcessing: {sheet_name}")

        # ----------------------------------------------------
        # REMOVE EMPTY ROWS / COLUMNS
        # ----------------------------------------------------

        df = df.dropna(
            how="all"
        ).copy()

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

        # ----------------------------------------------------
        # CHECK REQUIRED COLUMNS
        # ----------------------------------------------------

        missing_columns = [
            col
            for col in required_columns
            if col not in df.columns
        ]

        if missing_columns:

            print(
                f"SKIPPED: {sheet_name} "
                f"(missing columns: {missing_columns})"
            )

            continue

        # ----------------------------------------------------
        # CLEAN IMPORTANT COLUMNS
        # ----------------------------------------------------

        df["State"] = (
            df["State"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        df["Operation Type"] = (
            df["Operation Type"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        df["Last Updated By"] = (
            df["Last Updated By"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        # ----------------------------------------------------
        # DEMAND QTY
        # ----------------------------------------------------

        df["Demand Qty"] = pd.to_numeric(
            df["Demand Qty"],
            errors="coerce"
        ).fillna(0)

        # ====================================================
        # FILTER THIS SHEET
        #
        # State = DONE
        #
        # AND
        #
        # Operation Type =
        #     PICK
        #     OR
        #     INTERNAL TRANSFERS
        # ====================================================

        filtered_df = df[
            (
                df["State"]
                .str.lower()
                == "done"
            )
            &
            (
                df["Operation Type"]
                .str.lower()
                .isin([
                    "pick",
                    "internal transfers"
                ])
            )
        ].copy()

        # ----------------------------------------------------
        # SHEET NAME = STORE
        # ----------------------------------------------------

        filtered_df.insert(
            0,
            "STORE",
            sheet_name
        )

        store_filtered_data[
            sheet_name
        ] = filtered_df

        print(
            f"Total rows in sheet: {len(df)}"
        )

        print(
            f"Pick/Internal Transfer rows: "
            f"{len(filtered_df)}"
        )

    # ========================================================
    # 6. CHECK DATA
    # ========================================================

    valid_store_data = {
        store: data
        for store, data in store_filtered_data.items()
        if not data.empty
    }

    if not valid_store_data:

        raise ValueError(
            "No valid Pick/Internal Transfer data found."
        )

    # ========================================================
    # 7. COMBINE ALL FILTERED STORE DATA
    #
    # IMPORTANT:
    # We combine ONLY filtered rows.
    # ========================================================

    combined_filtered_data = pd.concat(
        valid_store_data.values(),
        ignore_index=True
    )

    # ========================================================
    # 8. CREATE FULL COMBINED REPORT
    # ========================================================

    combined_report = make_report_data(
        combined_filtered_data
    )

    combined_display = combined_report.drop(
        columns=["ROW_TYPE"]
    )

    # ========================================================
    # 9. CREATE DYNAMIC TITLE
    # ========================================================

    title_text = "STORE WISE PICK COUNT REPORT"

    if "Source Document Date" in combined_filtered_data.columns:

        dates = pd.to_datetime(
            combined_filtered_data["Source Document Date"],
            errors="coerce",
            dayfirst=True
        )

        if dates.notna().any():

            latest_date = dates.max()

            title_text = (
                f"STORE WISE PICK COUNT "
                f"{latest_date.strftime('%B').upper()} "
                f"- {latest_date.day}"
            )

    # ========================================================
    # 10. CREATE EXCEL
    #
    # SHEETS:
    #
    # 1. Store Wise Report
    # 2. Each individual store
    # ========================================================

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        # ====================================================
        # FULL COMBINED STORE REPORT
        # ====================================================

        combined_display.to_excel(
            writer,
            sheet_name="Store Wise Report",
            index=False,
            startrow=1
        )

        # ====================================================
        # INDIVIDUAL STORE REPORTS
        # ====================================================

        used_sheet_names = {
            "Store Wise Report"
        }

        for store, store_df in valid_store_data.items():

            # -----------------------------------------------
            # Create report for this store only
            # -----------------------------------------------

            individual_report = make_report_data(
                store_df
            )

            individual_display = individual_report.drop(
                columns=["ROW_TYPE"]
            )

            # -----------------------------------------------
            # Safe Excel sheet name
            # -----------------------------------------------

            base_sheet_name = clean_sheet_name(
                store
            )

            sheet_name = base_sheet_name

            counter = 1

            while sheet_name in used_sheet_names:

                suffix = f"_{counter}"

                sheet_name = (
                    base_sheet_name[:31 - len(suffix)]
                    + suffix
                )

                counter += 1

            used_sheet_names.add(
                sheet_name
            )

            # -----------------------------------------------
            # Write individual report
            # -----------------------------------------------

            individual_display.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                startrow=1
            )

    # ========================================================
    # 11. OPEN WORKBOOK FOR FORMATTING
    # ========================================================

    wb = load_workbook(
        output_file
    )

    # ========================================================
    # 12. STYLES
    # ========================================================

    yellow_fill = PatternFill(
        "solid",
        fgColor="FFFF00"
    )

    dark_blue_fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )

    white_font = Font(
        color="FFFFFF",
        bold=True
    )

    bold_font = Font(
        bold=True
    )

    title_font = Font(
        bold=True,
        size=14,
        color="1F4E78"
    )

    thin_side = Side(
        style="thin",
        color="000000"
    )

    border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side
    )

    # ========================================================
    # 13. FORMAT ALL REPORT SHEETS
    # ========================================================

    for ws in wb.worksheets:

        # ----------------------------------------------------
        # TITLE
        # ----------------------------------------------------

        ws.merge_cells(
            "A1:E1"
        )

        if ws.title == "Store Wise Report":

            ws["A1"] = title_text

        else:

            ws["A1"] = (
                f"{ws.title.upper()} "
                f"- PICK / INTERNAL TRANSFER REPORT"
            )

        ws["A1"].font = title_font

        ws["A1"].alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        ws["A1"].fill = yellow_fill

        ws.row_dimensions[1].height = 25

        # ----------------------------------------------------
        # HEADER
        # ----------------------------------------------------

        for cell in ws[2]:

            cell.fill = dark_blue_fill

            cell.font = white_font

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

            cell.border = border

        ws.row_dimensions[2].height = 30

        # ----------------------------------------------------
        # DATA ROWS
        # ----------------------------------------------------

        for row in ws.iter_rows(
            min_row=3,
            max_row=ws.max_row,
            min_col=1,
            max_col=5
        ):

            store_value = row[0].value

            # ----------------------------------------------
            # TOTAL ROWS
            # ----------------------------------------------

            if (
                store_value
                and
                (
                    str(store_value).endswith(" Total")
                    or
                    str(store_value) == "Grand Total"
                )
            ):

                for cell in row:

                    cell.fill = yellow_fill

                    cell.font = bold_font

                    cell.border = border

                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center"
                    )

            # ----------------------------------------------
            # NORMAL ROW
            # ----------------------------------------------

            else:

                for cell in row:

                    cell.border = border

                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center"
                    )

        # ----------------------------------------------------
        # NUMBER FORMATTING
        # ----------------------------------------------------

        for row in ws.iter_rows(
            min_row=3,
            max_row=ws.max_row
        ):

            row[3].number_format = "0"
            row[4].number_format = "0"

        # ----------------------------------------------------
        # COLUMN WIDTHS
        # ----------------------------------------------------

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 36
        ws.column_dimensions["C"].width = 23
        ws.column_dimensions["D"].width = 43
        ws.column_dimensions["E"].width = 30

        # ----------------------------------------------------
        # ROW HEIGHT
        # ----------------------------------------------------

        for row_number in range(
            3,
            ws.max_row + 1
        ):

            ws.row_dimensions[
                row_number
            ].height = 20

        # ----------------------------------------------------
        # FREEZE HEADER
        # ----------------------------------------------------

        ws.freeze_panes = "A3"

        # ----------------------------------------------------
        # AUTO FILTER
        # ----------------------------------------------------

        if ws.max_row >= 2:

            ws.auto_filter.ref = (
                f"A2:E{ws.max_row}"
            )

    # ========================================================
    # 14. SAVE
    # ========================================================

    wb.save(
        output_file
    )

    # ========================================================
    # 15. SUCCESS MESSAGE
    # ========================================================

    print("\n")
    print("==============================================")
    print("       REPORT CREATED SUCCESSFULLY")
    print("==============================================")

    print(
        "Output file:",
        output_file
    )

    print("\nExcluded sheets:")

    for sheet in EXCLUDED_SHEETS:
        print(" -", sheet)

    print("\nCreated sheets:")

    print(" - Store Wise Report (ALL STORES COMBINED)")

    for store in valid_store_data.keys():
        print(" -", store)

    print("\nReport logic:")

    print("1. Each Excel sheet is treated as a STORE.")
    print("2. Each sheet is filtered individually.")
    print("3. State must be DONE.")
    print("4. Operation Type must be PICK or INTERNAL TRANSFERS.")
    print("5. Only filtered rows are combined.")
    print("6. Full combined Store Wise Report is created.")
    print("7. Individual report sheet is created for every store.")
    print("8. Store totals include Pick + Internal Transfers.")
    print("9. Grand Total includes all stores.")

    print("==============================================")

    return output_file
