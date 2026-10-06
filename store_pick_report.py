# ============================================================
# STORE WISE PICK / INTERNAL TRANSFER REPORT
# ============================================================

import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def generate_store_pick_report(file_path, output_file):

    # ============================================================
    # 1. READ ALL SHEETS
    #
    # FIRST 4 ROWS ARE NOT REQUIRED
    # ROW 5 = HEADER
    # ROW 6 = DATA
    # ============================================================

    all_sheets = pd.read_excel(
        file_path,
        sheet_name=None,
        header=4
    )


    # ============================================================
    # 2. SHEETS TO EXCLUDE
    # ============================================================

    EXCLUDED_SHEETS = [
        "SPM_Virtual_Store",
        "Export(SPM)",
        "Virtual_Store_AD"
    ]


    # ============================================================
    # 3. SHOW ALL SHEETS
    # ============================================================

    print("\nAll sheets found:")

    for sheet_name in all_sheets.keys():

        print(" -", sheet_name)


    # ============================================================
    # 4. PROCESS ONLY REQUIRED SHEETS
    # ============================================================

    all_data = []


    for sheet_name, df in all_sheets.items():

        # --------------------------------------------------------
        # SKIP EXCLUDED SHEETS
        # --------------------------------------------------------

        if sheet_name in EXCLUDED_SHEETS:

            print(
                f"SKIPPED: {sheet_name}"
            )

            continue


        print(
            f"Processing: {sheet_name}"
        )


        # --------------------------------------------------------
        # REMOVE COMPLETELY EMPTY ROWS
        # --------------------------------------------------------

        df = df.dropna(
            how="all"
        ).copy()


        # --------------------------------------------------------
        # REMOVE COMPLETELY EMPTY COLUMNS
        # --------------------------------------------------------

        df = df.dropna(
            axis=1,
            how="all"
        )


        # --------------------------------------------------------
        # CLEAN COLUMN NAMES
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
        # ADD DATAFRAME TO LIST
        # --------------------------------------------------------

        all_data.append(
            df
        )


    # ============================================================
    # 5. CHECK WHETHER ANY SHEETS ARE AVAILABLE
    # ============================================================

    if not all_data:

        raise ValueError(
            "No valid sheets found after excluding the specified sheets."
        )


    # ============================================================
    # 6. APPEND ALL VALID SHEETS
    # ============================================================

    combined_data = pd.concat(
        all_data,
        ignore_index=True
    )


    # ============================================================
    # 7. CLEAN COLUMN NAMES
    # ============================================================

    combined_data.columns = (
        combined_data.columns
        .astype(str)
        .str.strip()
    )


    # ============================================================
    # 8. SHOW COMBINED DATA INFORMATION
    # ============================================================

    print("\n==============================================")
    print("COMBINED DATA")
    print("==============================================")

    print(
        "Rows:",
        len(combined_data)
    )

    print(
        "Columns:",
        len(combined_data.columns)
    )


    print("\nColumns:")


    for col in combined_data.columns:

        print(
            " -",
            col
        )


    # ============================================================
    # 9. CHECK REQUIRED COLUMNS
    #
    # SOURCE DOCUMENT IS REQUIRED
    #
    # SOURCE DOCUMENT = DISTINCT COUNT
    # LAST UPDATED BY = PERSON
    # LAST UPDATED DATE = TITLE DATE ONLY
    # ============================================================

    required_columns = [
        "STORE",
        "Source Document",
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

            print(
                " -",
                col
            )


        print("\nActual columns found:")


        for col in combined_data.columns:

            print(
                " -",
                col
            )


        raise ValueError(
            "Required columns are missing from the Excel file."
        )


    # ============================================================
    # 10. CLEAN STATE
    # ============================================================

    combined_data["State"] = (
        combined_data["State"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    # ============================================================
    # 11. CLEAN OPERATION TYPE
    # ============================================================

    combined_data["Operation Type"] = (
        combined_data["Operation Type"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    # ============================================================
    # 12. CLEAN LAST UPDATED BY
    #
    # THIS COLUMN STAYS IN THE REPORT
    # ============================================================

    combined_data["Last Updated By"] = (
        combined_data["Last Updated By"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    # ============================================================
    # 13. CLEAN SOURCE DOCUMENT
    #
    # SOURCE DOCUMENT WILL BE USED FOR DISTINCT COUNT
    # ============================================================

    combined_data["Source Document"] = (
        combined_data["Source Document"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    # ============================================================
    # 14. CONVERT DEMAND QTY TO NUMBER
    # ============================================================

    combined_data["Demand Qty"] = pd.to_numeric(
        combined_data["Demand Qty"],
        errors="coerce"
    ).fillna(0)


    # ============================================================
    # 15. CLEAN LAST UPDATED DATE
    #
    # USED ONLY FOR REPORT TITLE
    #
    # TIME IS IGNORED
    # ============================================================

    combined_data["Last Updated Date"] = pd.to_datetime(
        combined_data["Last Updated Date"],
        errors="coerce"
    )


    # ============================================================
    # 16. FILTER DATA
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
    # 17. SHOW FILTER RESULT
    # ============================================================

    print("\n==============================================")
    print("FILTERED DATA")
    print("==============================================")

    print(
        "Filtered rows:",
        len(filtered_data)
    )


    # ============================================================
    # 18. CREATE PIVOT DATA
    #
    # STORE
    # +
    # LAST UPDATED BY
    # +
    # OPERATION TYPE
    #
    # DISTINCT COUNT = SOURCE DOCUMENT
    #
    # DEMAND = SUM OF DEMAND QTY
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

                # -----------------------------------------------
                # DISTINCT COUNT OF SOURCE DOCUMENT
                # -----------------------------------------------

                "Distinct Count of Source Document":
                    (
                        "Source Document",
                        "nunique"
                    ),

                # -----------------------------------------------
                # SUM OF DEMAND QTY
                # -----------------------------------------------

                "Sum of Demand Qty":
                    (
                        "Demand Qty",
                        "sum"
                    )

            }
        )
        .reset_index()
    )


    # ============================================================
    # 19. SORT PIVOT DATA
    #
    # INTERNAL TRANSFERS FIRST
    # PICK SECOND
    # ============================================================

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
        columns=[
            "_operation_order"
        ]
    )


    # ============================================================
    # 20. CREATE FINAL REPORT
    #
    # STORE APPEARS ONLY ON FIRST ROW
    #
    # LAST UPDATED BY APPEARS ONLY ON FIRST ROW
    #
    # EXAMPLE:
    #
    # Sharjah_UAE | JIBIN | Internal Transfers | 27   | 57
    #             |       | Pick               | 157  | 322
    #
    #             | KANJAN B.K | Internal Transfers | 1413 | 6102
    #             |            | Pick               | 1463 | 10685
    #
    # ============================================================

    final_report = []


    for store, store_df in pivot_data.groupby(
        "STORE",
        sort=False
    ):

        first_store_row = True


        # --------------------------------------------------------
        # GROUP BY LAST UPDATED BY
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


                # ------------------------------------------------
                # STORE APPEARS ONLY ON FIRST ROW
                # ------------------------------------------------

                if first_store_row:

                    display_store = store

                    first_store_row = False

                else:

                    display_store = ""


                # ------------------------------------------------
                # PERSON APPEARS ONLY ON FIRST ROW
                # ------------------------------------------------

                if first_person_row:

                    display_person = person

                    first_person_row = False

                else:

                    display_person = ""


                # ------------------------------------------------
                # ADD DATA ROW
                # ------------------------------------------------

                final_report.append({

                    "STORE":
                        display_store,

                    "Last Updated By":
                        display_person,

                    "Operation Type":
                        row[
                            "Operation Type"
                        ],

                    "Distinct Count of Source Document":
                        row[
                            "Distinct Count of Source Document"
                        ],

                    "Sum of Demand Qty":
                        row[
                            "Sum of Demand Qty"
                        ],

                    "ROW_TYPE":
                        "DATA"

                })


        # ========================================================
        # STORE TOTAL
        # ========================================================

        store_source_count = (
            filtered_data[
                filtered_data["STORE"] == store
            ]["Source Document"]
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

            "Distinct Count of Source Document":
                store_source_count,

            "Sum of Demand Qty":
                store_demand_qty,

            "ROW_TYPE":
                "TOTAL"

        })


    # ============================================================
    # 21. CONVERT FINAL REPORT TO DATAFRAME
    # ============================================================

    final_report = pd.DataFrame(
        final_report
    )


    # ============================================================
    # 22. GRAND TOTAL
    #
    # DISTINCT COUNT = SOURCE DOCUMENT
    # ============================================================

    grand_source_count = (
        filtered_data[
            "Source Document"
        ].nunique()
    )


    grand_demand_qty = (
        filtered_data[
            "Demand Qty"
        ].sum()
    )


    grand_total = pd.DataFrame([{

        "STORE":
            "Grand Total",

        "Last Updated By":
            "",

        "Operation Type":
            "",

        "Distinct Count of Source Document":
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
    # 23. REMOVE ROW_TYPE FROM DISPLAY
    # ============================================================

    report_display = final_report.drop(
        columns=[
            "ROW_TYPE"
        ]
    )


    # ============================================================
    # 24. CREATE DYNAMIC TITLE
    #
    # LAST UPDATED DATE IS USED ONLY FOR TITLE
    # ============================================================

    title_text = (
        "STORE WISE PICK COUNT REPORT"
    )


    if "Last Updated Date" in combined_data.columns:

        dates = pd.to_datetime(
            combined_data[
                "Last Updated Date"
            ],
            errors="coerce"
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


    print(
        title_text
    )


    # ============================================================
    # 25. WRITE EXCEL
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
        # COMBINED DATA
        # --------------------------------------------------------

        combined_data.to_excel(
            writer,
            sheet_name="Combined Data",
            index=False
        )


        # --------------------------------------------------------
        # FILTERED DATA
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
    # 26. OPEN WORKBOOK
    # ============================================================

    wb = load_workbook(
        output_file
    )


    # ============================================================
    # 27. COLORS
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
    # 28. FONTS
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
    # 29. BORDER
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
    # 30. FORMAT COMBINED DATA
    # ============================================================

    ws = wb[
        "Combined Data"
    ]


    ws.freeze_panes = "A2"


    for cell in ws[1]:

        cell.fill = blue_fill

        cell.font = white_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        cell.border = border


    # ------------------------------------------------------------
    # AUTO COLUMN WIDTH
    # ------------------------------------------------------------

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
    # 31. FORMAT FILTERED DATA
    # ============================================================

    ws = wb[
        "Filtered Data"
    ]


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
    # 32. FORMAT STORE WISE REPORT
    # ============================================================

    ws = wb[
        "Store Wise Report"
    ]


    # ============================================================
    # 33. TITLE
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


    ws.row_dimensions[
        1
    ].height = 25


    # ============================================================
    # 34. HEADER
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


    ws.row_dimensions[
        2
    ].height = 22


    # ============================================================
    # 35. DATA ROWS
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
                str(store_value).endswith(
                    " Total"
                )
                or
                str(store_value)
                == "Grand Total"
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
    # 36. NUMBER FORMATTING
    # ============================================================

    for row in ws.iter_rows(
        min_row=3,
        max_row=ws.max_row
    ):

        # --------------------------------------------------------
        # DISTINCT COUNT OF SOURCE DOCUMENT
        # --------------------------------------------------------

        row[3].number_format = "0"


        # --------------------------------------------------------
        # DEMAND QTY
        # --------------------------------------------------------

        row[4].number_format = "0"


    # ============================================================
    # 37. COLUMN WIDTHS
    # ============================================================

    ws.column_dimensions[
        "A"
    ].width = 28


    ws.column_dimensions[
        "B"
    ].width = 36


    ws.column_dimensions[
        "C"
    ].width = 23


    ws.column_dimensions[
        "D"
    ].width = 43


    ws.column_dimensions[
        "E"
    ].width = 30


    # ============================================================
    # 38. ROW HEIGHT
    # ============================================================

    for row_number in range(
        3,
        ws.max_row + 1
    ):

        ws.row_dimensions[
            row_number
        ].height = 20


    # ============================================================
    # 39. FREEZE HEADER
    # ============================================================

    ws.freeze_panes = "A3"


    # ============================================================
    # 40. AUTO FILTER
    # ============================================================

    ws.auto_filter.ref = (
        f"A2:E{ws.max_row}"
    )


    # ============================================================
    # 41. SAVE FILE
    # ============================================================

    wb.save(
        output_file
    )


    # ============================================================
    # 42. SUCCESS MESSAGE
    # ============================================================

    print("\n")

    print(
        "=============================================="
    )

    print(
        "       REPORT CREATED SUCCESSFULLY"
    )

    print(
        "=============================================="
    )


    print(
        "Output file:",
        output_file
    )


    print("\nExcluded sheets:")


    for sheet in EXCLUDED_SHEETS:

        print(
            " -",
            sheet
        )


    print("\nReport contains:")


    print(
        "1. Combined Data"
    )


    print(
        "2. Filtered Data"
    )


    print(
        "3. Store Wise Report"
    )


    print(
        "\nStore and Last Updated By names appear only on their first row."
    )


    print(
        "Distinct Count is calculated from Source Document."
    )


    print(
        "Last Updated Date is used only for the report title."
    )


    print(
        "Store totals and Grand Total are included."
    )


    print(
        "=============================================="
    )


    return output_file
