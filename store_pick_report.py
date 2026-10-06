# ============================================================
# STORE WISE PICK / INTERNAL TRANSFER REPORT
# ============================================================

import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def generate_store_pick_report(file_path, output_file):

    # ============================================================
    # 2. READ ALL SHEETS
    #
    # FIRST 5 ROWS ARE NOT REQUIRED
    # ROW 6 = HEADER
    # ============================================================

    all_sheets = pd.read_excel(
        file_path,
        sheet_name=None,
        skiprows=5
    )


    # ============================================================
    # 3. SHEETS TO EXCLUDE
    # ============================================================

    EXCLUDED_SHEETS = [
        "SPM_Virtual_Store",
        "Export(SPM)",
        "Virtual_Store_AD"
    ]


    # ============================================================
    # 4. SHOW ALL SHEETS
    # ============================================================

    print("\nAll sheets found:")

    for sheet_name in all_sheets.keys():
        print(" -", sheet_name)


    # ============================================================
    # 5. PROCESS ONLY REQUIRED SHEETS
    # ============================================================

    all_data = []

    for sheet_name, df in all_sheets.items():

        # --------------------------------------------------------
        # SKIP EXCLUDED SHEETS
        # --------------------------------------------------------

        if sheet_name in EXCLUDED_SHEETS:

            print(f"SKIPPED: {sheet_name}")

            continue

        print(f"Processing: {sheet_name}")

        # --------------------------------------------------------
        # Remove completely empty rows
        # --------------------------------------------------------

        df = df.dropna(
            how="all"
        ).copy()

        # --------------------------------------------------------
        # Remove completely empty columns
        # --------------------------------------------------------

        df = df.dropna(
            axis=1,
            how="all"
        )

        # --------------------------------------------------------
        # Clean column names
        # --------------------------------------------------------

        df.columns = (
            df.columns
            .astype(str)
            .str.strip()
        )

        # --------------------------------------------------------
        # ADD STORE COLUMN
        #
        # STORE = SHEET NAME
        # --------------------------------------------------------

        df.insert(
            0,
            "STORE",
            sheet_name
        )

        # --------------------------------------------------------
        # Add dataframe to list
        # --------------------------------------------------------

        all_data.append(df)


    # ============================================================
    # 6. CHECK WHETHER ANY SHEETS ARE AVAILABLE
    # ============================================================

    if not all_data:

        raise ValueError(
            "No valid sheets found after excluding the specified sheets."
        )


    # ============================================================
    # 7. APPEND ALL VALID SHEETS
    # ============================================================

    combined_data = pd.concat(
        all_data,
        ignore_index=True
    )


    # ============================================================
    # 8. CLEAN COLUMN NAMES
    # ============================================================

    combined_data.columns = (
        combined_data.columns
        .astype(str)
        .str.strip()
    )


    # ============================================================
    # 9. SHOW COMBINED DATA INFORMATION
    # ============================================================

    print("\n==============================================")
    print("COMBINED DATA")
    print("==============================================")

    print("Rows:", len(combined_data))
    print("Columns:", len(combined_data.columns))

    print("\nColumns:")

    for col in combined_data.columns:
        print(" -", col)


    # ============================================================
    # 10. CHECK REQUIRED COLUMNS
    #
    # SOURCE DOCUMENT CHANGED TO LAST UPDATED DATE
    # ============================================================

    required_columns = [
        "STORE",
        "Last Updated Date",
        "Operation Type",
        "Demand Qty",
        "State",
        "Last Updated By"
    ]


    missing_columns = [
        col
        for col in required_columns
        if col not in combined_data.columns
    ]


    if missing_columns:

        print("\nERROR - Missing columns:")

        for col in missing_columns:
            print(" -", col)

        raise ValueError(
            "Required columns are missing from the Excel file."
        )


    # ============================================================
    # 11. CLEAN IMPORTANT COLUMNS
    # ============================================================

    combined_data["State"] = (
        combined_data["State"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    combined_data["Operation Type"] = (
        combined_data["Operation Type"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    combined_data["Last Updated By"] = (
        combined_data["Last Updated By"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    # ============================================================
    # 12. CONVERT DEMAND QTY TO NUMBER
    # ============================================================

    combined_data["Demand Qty"] = pd.to_numeric(
        combined_data["Demand Qty"],
        errors="coerce"
    ).fillna(0)


    # ============================================================
    # 13. CLEAN LAST UPDATED DATE
    # ============================================================

    combined_data["Last Updated Date"] = pd.to_datetime(
        combined_data["Last Updated Date"],
        errors="coerce",
        dayfirst=True
    )


    # ============================================================
    # 14. FILTER DATA
    #
    # STATE = DONE
    #
    # OPERATION TYPE =
    # PICK
    # OR
    # INTERNAL TRANSFERS
    # ============================================================

    filtered_data = combined_data[
        (
            combined_data["State"]
            .str.lower()
            == "done"
        )
        &
        (
            combined_data["Operation Type"]
            .str.lower()
            .isin([
                "pick",
                "internal transfers"
            ])
        )
    ].copy()


    # ============================================================
    # 15. SHOW FILTER RESULT
    # ============================================================

    print("\n==============================================")
    print("FILTERED DATA")
    print("==============================================")

    print("Filtered rows:", len(filtered_data))


    # ============================================================
    # 16. CREATE PIVOT DATA
    #
    # LAST UPDATED DATE IS NOW USED INSTEAD OF
    # SOURCE DOCUMENT
    # ============================================================

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
                "Distinct Count of Last Updated Date":
                    ("Last Updated Date", "nunique"),

                "Sum of Demand Qty":
                    ("Demand Qty", "sum")
            }
        )
        .reset_index()
    )


    # ============================================================
    # 17. SORT PIVOT DATA
    # ============================================================

    # Keep Internal Transfers before Pick

    operation_order = {
        "internal transfers": 0,
        "pick": 1
    }

    pivot_data["_operation_order"] = (
        pivot_data["Operation Type"]
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


    # ============================================================
    # 18. CREATE FINAL REPORT
    #
    # STORE + PERSON WILL APPEAR ONLY ONCE
    #
    # Example:
    #
    # Qusais_Sales SPM | KHAYAZ | Internal Transfers | 35 | 89
    #                 |        | Pick              |  8 | 24
    #
    # ============================================================

    final_report = []


    for store, store_df in pivot_data.groupby(
        "STORE",
        sort=False
    ):

        first_store_row = True

        # --------------------------------------------------------
        # GROUP BY PERSON
        # --------------------------------------------------------

        for person, person_df in store_df.groupby(
            "Last Updated By",
            sort=False
        ):

            first_person_row = True

            # ----------------------------------------------------
            # OPERATION ROWS
            # ----------------------------------------------------

            for _, row in person_df.iterrows():

                # STORE appears only on first row

                if first_store_row:

                    display_store = store

                    first_store_row = False

                else:

                    display_store = ""


                # PERSON appears only on first row

                if first_person_row:

                    display_person = person

                    first_person_row = False

                else:

                    display_person = ""


                final_report.append({

                    "STORE":
                        display_store,

                    "Last Updated By":
                        display_person,

                    "Operation Type":
                        row["Operation Type"],

                    "Distinct Count of Last Updated Date":
                        row["Distinct Count of Last Updated Date"],

                    "Sum of Demand Qty":
                        row["Sum of Demand Qty"],

                    "ROW_TYPE":
                        "DATA"

                })


        # ========================================================
        # STORE TOTAL
        # ========================================================

        store_source_count = (
            filtered_data[
                filtered_data["STORE"] == store
            ]["Last Updated Date"]
            .nunique()
        )


        store_demand_qty = (
            filtered_data[
                filtered_data["STORE"] == store
            ]["Demand Qty"]
            .sum()
        )


        final_report.append({

            "STORE":
                f"{store} Total",

            "Last Updated By":
                "",

            "Operation Type":
                "",

            "Distinct Count of Last Updated Date":
                store_source_count,

            "Sum of Demand Qty":
                store_demand_qty,

            "ROW_TYPE":
                "TOTAL"

        })


    # ============================================================
    # 19. CONVERT FINAL REPORT TO DATAFRAME
    # ============================================================

    final_report = pd.DataFrame(
        final_report
    )


    # ============================================================
    # 20. GRAND TOTAL
    # ============================================================

    grand_source_count = (
        filtered_data["Last Updated Date"]
        .nunique()
    )


    grand_demand_qty = (
        filtered_data["Demand Qty"]
        .sum()
    )


    grand_total = pd.DataFrame([{

        "STORE":
            "Grand Total",

        "Last Updated By":
            "",

        "Operation Type":
            "",

        "Distinct Count of Last Updated Date":
            grand_source_count,

        "Sum of Demand Qty":
            grand_demand_qty,

        "ROW_TYPE":
            "GRAND_TOTAL"

    }])


    final_report = pd.concat(
        [
            final_report,
            grand_total
        ],
        ignore_index=True
    )


    # ============================================================
    # 21. REMOVE ROW_TYPE FROM DISPLAY
    # ============================================================

    report_display = final_report.drop(
        columns=["ROW_TYPE"]
    )


    # ============================================================
    # 22. CREATE DYNAMIC TITLE
    #
    # USE LAST UPDATED DATE
    # ============================================================

    title_text = "STORE WISE PICK COUNT REPORT"


    if "Last Updated Date" in combined_data.columns:

        dates = pd.to_datetime(
            combined_data["Last Updated Date"],
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


    print("\n==============================================")
    print("REPORT TITLE")
    print("==============================================")

    print(title_text)


    # ============================================================
    # 23. WRITE EXCEL
    #
    # ROW 1 = TITLE
    # ROW 2 = HEADER
    # ROW 3 = DATA
    # ============================================================

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        # --------------------------------------------------------
        # COMBINED DATA SHEET
        # --------------------------------------------------------

        combined_data.to_excel(
            writer,
            sheet_name="Combined Data",
            index=False
        )

        # --------------------------------------------------------
        # FILTERED DATA SHEET
        # --------------------------------------------------------

        filtered_data.to_excel(
            writer,
            sheet_name="Filtered Data",
            index=False
        )

        # --------------------------------------------------------
        # STORE WISE REPORT
        # --------------------------------------------------------

        report_display.to_excel(
            writer,
            sheet_name="Store Wise Report",
            index=False,
            startrow=1
        )


    # ============================================================
    # 24. OPEN WORKBOOK
    # ============================================================

    wb = load_workbook(
        output_file
    )


    # ============================================================
    # 25. COLORS
    # ============================================================

    yellow_fill = PatternFill(
        "solid",
        fgColor="FFFF00"
    )


    dark_blue_fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )


    blue_fill = PatternFill(
        "solid",
        fgColor="4472C4"
    )


    # ============================================================
    # 26. FONTS
    # ============================================================

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


    # ============================================================
    # 27. BORDER
    # ============================================================

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


    # ============================================================
    # 28. FORMAT COMBINED DATA
    # ============================================================

    ws = wb["Combined Data"]

    ws.freeze_panes = "A2"


    for cell in ws[1]:

        cell.fill = blue_fill

        cell.font = white_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = border


    # Auto column width

    for column_cells in ws.columns:

        max_length = 0

        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        ws.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            35
        )


    # ============================================================
    # 29. FORMAT FILTERED DATA
    # ============================================================

    ws = wb["Filtered Data"]

    ws.freeze_panes = "A2"


    for cell in ws[1]:

        cell.fill = blue_fill

        cell.font = white_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = border


    for column_cells in ws.columns:

        max_length = 0

        column_letter = get_column_letter(
            column_cells[0].column
        )

        for cell in column_cells:

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        ws.column_dimensions[
            column_letter
        ].width = min(
            max_length + 2,
            35
        )


    # ============================================================
    # 30. FORMAT STORE WISE REPORT
    # ============================================================

    ws = wb["Store Wise Report"]


    # ============================================================
    # 31. TITLE
    # ============================================================

    ws.merge_cells(
        "A1:E1"
    )


    ws["A1"] = title_text


    ws["A1"].font = title_font


    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )


    ws["A1"].fill = yellow_fill


    ws.row_dimensions[1].height = 25


    # ============================================================
    # 32. HEADER
    # ============================================================

    for cell in ws[2]:

        cell.fill = dark_blue_fill

        cell.font = white_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        cell.border = border


    ws.row_dimensions[2].height = 22


    # ============================================================
    # 33. DATA ROWS
    # ============================================================

    for row in ws.iter_rows(
        min_row=3,
        max_row=ws.max_row,
        min_col=1,
        max_col=5
    ):

        store_value = row[0].value

        # --------------------------------------------------------
        # STORE TOTAL / GRAND TOTAL
        # --------------------------------------------------------

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

        # --------------------------------------------------------
        # NORMAL DATA ROW
        # --------------------------------------------------------

        else:

            for cell in row:

                cell.border = border

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )


    # ============================================================
    # 34. NUMBER FORMATTING
    # ============================================================

    for row in ws.iter_rows(
        min_row=3,
        max_row=ws.max_row
    ):

        # Distinct Count of Last Updated Date
        row[3].number_format = "0"

        # Demand Qty
        row[4].number_format = "0"


    # ============================================================
    # 35. COLUMN WIDTHS
    # ============================================================

    ws.column_dimensions["A"].width = 28

    ws.column_dimensions["B"].width = 36

    ws.column_dimensions["C"].width = 23

    ws.column_dimensions["D"].width = 43

    ws.column_dimensions["E"].width = 30


    # ============================================================
    # 36. ROW HEIGHT
    # ============================================================

    for row_number in range(
        3,
        ws.max_row + 1
    ):

        ws.row_dimensions[
            row_number
        ].height = 20


    # ============================================================
    # 37. FREEZE HEADER
    # ============================================================

    ws.freeze_panes = "A3"


    # ============================================================
    # 38. AUTO FILTER
    # ============================================================

    ws.auto_filter.ref = (
        f"A2:E{ws.max_row}"
    )


    # ============================================================
    # 39. SAVE FILE
    # ============================================================

    wb.save(
        output_file
    )


    # ============================================================
    # 40. SUCCESS MESSAGE
    # ============================================================

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


    print("\nReport contains:")

    print("1. Combined Data")
    print("2. Filtered Data")
    print("3. Store Wise Report")


    print(
        "\nStore and person names appear only on their first row."
    )

    print(
        "Excluded sheets are NOT included in the report."
    )

    print(
        "Store totals and Grand Total are included."
    )

    print("==============================================")


    return output_file
