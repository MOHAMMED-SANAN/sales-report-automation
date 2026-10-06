# ============================================================
# STORE WISE PICK / INTERNAL TRANSFER REPORT
#
# OUTPUT:
#
# 1. MASTER DATA
#       -> ALL STORES IN ONE SUMMARY REPORT
#
# 2. INDIVIDUAL STORE SHEETS
#       -> ONE SHEET PER STORE
#
# DISTINCT COUNT = SOURCE DOCUMENT
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
    # ROW 5 = HEADER
    # ROW 6 = DATA
    # ========================================================

    all_sheets = pd.read_excel(
        file_path,
        sheet_name=None,
        header=4
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

            print(
                f"SKIPPED: {sheet_name}"
            )

            continue


        print(
            f"Reading: {sheet_name}"
        )


        # ----------------------------------------------------
        # REMOVE EMPTY ROWS
        # ----------------------------------------------------

        df = df.dropna(
            how="all"
        ).copy()


        # ----------------------------------------------------
        # REMOVE EMPTY COLUMNS
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


        # ----------------------------------------------------
        # ADD STORE COLUMN
        # ----------------------------------------------------

        df.insert(
            0,
            "STORE",
            sheet_name
        )


        # ----------------------------------------------------
        # ADD TO MASTER LIST
        # ----------------------------------------------------

        all_data.append(
            df
        )


    # ========================================================
    # 4. CHECK DATA
    # ========================================================

    if not all_data:

        raise ValueError(
            "No valid store sheets found."
        )


    # ========================================================
    # 5. COMBINE ALL DATA
    # ========================================================

    master_raw_data = pd.concat(
        all_data,
        ignore_index=True
    )


    # ========================================================
    # 6. CLEAN COLUMN NAMES
    # ========================================================

    master_raw_data.columns = (
        master_raw_data.columns
        .astype(str)
        .str.strip()
    )


    # ========================================================
    # 7. REQUIRED COLUMNS
    # ========================================================

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

        if col not in master_raw_data.columns

    ]


    if missing_columns:

        print(
            "\nMissing columns:"
        )


        for col in missing_columns:

            print(
                " -",
                col
            )


        print(
            "\nAvailable columns:"
        )


        for col in master_raw_data.columns:

            print(
                " -",
                col
            )


        raise ValueError(
            "Required columns are missing from the Excel file."
        )


    # ========================================================
    # 8. CLEAN STATE
    # ========================================================

    master_raw_data["State"] = (

        master_raw_data["State"]

        .fillna("")

        .astype(str)

        .str.strip()

    )


    # ========================================================
    # 9. CLEAN OPERATION TYPE
    # ========================================================

    master_raw_data["Operation Type"] = (

        master_raw_data["Operation Type"]

        .fillna("")

        .astype(str)

        .str.strip()

    )


    # ========================================================
    # 10. CLEAN LAST UPDATED BY
    # ========================================================

    master_raw_data["Last Updated By"] = (

        master_raw_data["Last Updated By"]

        .fillna("")

        .astype(str)

        .str.strip()

    )


    # ========================================================
    # 11. CLEAN SOURCE DOCUMENT
    #
    # THIS IS THE DISTINCT COUNT FIELD
    # ========================================================

    master_raw_data["Source Document"] = (

        master_raw_data["Source Document"]

        .fillna("")

        .astype(str)

        .str.strip()

    )


    # ========================================================
    # 12. CONVERT DEMAND QTY
    # ========================================================

    master_raw_data["Demand Qty"] = pd.to_numeric(

        master_raw_data["Demand Qty"],

        errors="coerce"

    ).fillna(0)


    # ========================================================
    # 13. CONVERT LAST UPDATED DATE
    #
    # ONLY USED FOR REPORT TITLE
    # ========================================================

    master_raw_data["Last Updated Date"] = pd.to_datetime(

        master_raw_data["Last Updated Date"],

        errors="coerce"

    )


    # ========================================================
    # 14. FILTER DATA
    #
    # STATE = DONE
    #
    # OPERATION TYPE =
    # PICK
    # INTERNAL TRANSFERS
    # ========================================================

    filtered_data = master_raw_data[

        (

            master_raw_data["State"]

            .str.lower()

            == "done"

        )

        &

        (

            master_raw_data["Operation Type"]

            .str.lower()

            .isin([

                "pick",

                "internal transfers"

            ])

        )

    ].copy()


    # ========================================================
    # 15. OPERATION ORDER
    #
    # INTERNAL TRANSFERS FIRST
    # PICK SECOND
    # ========================================================

    operation_order = {

        "internal transfers": 0,

        "pick": 1

    }


    filtered_data["_operation_order"] = (

        filtered_data["Operation Type"]

        .str.lower()

        .map(operation_order)

        .fillna(99)

    )


    # ========================================================
    # 16. STORE ORDER
    #
    # SAME ORDER AS ORIGINAL EXCEL SHEETS
    # ========================================================

    store_list = [

        sheet_name

        for sheet_name in all_sheets.keys()

        if sheet_name not in EXCLUDED_SHEETS

    ]


    store_order_map = {

        store: i

        for i, store in enumerate(store_list)

    }


    filtered_data["_store_order"] = (

        filtered_data["STORE"]

        .map(store_order_map)

        .fillna(9999)

    )


    # ========================================================
    # 17. SORT DATA
    # ========================================================

    filtered_data = filtered_data.sort_values(

        by=[

            "_store_order",

            "Last Updated By",

            "_operation_order"

        ],

        kind="stable"

    )


    # ========================================================
    # 18. CREATE STORE REPORT FUNCTION
    # ========================================================

    def create_store_report(store_name):

        # ----------------------------------------------------
        # GET ONLY CURRENT STORE
        # ----------------------------------------------------

        store_data = filtered_data[

            filtered_data["STORE"]

            == store_name

        ].copy()


        # ----------------------------------------------------
        # IF EMPTY
        # ----------------------------------------------------

        if store_data.empty:

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


        # ====================================================
        # GROUP DATA
        #
        # DISTINCT COUNT = SOURCE DOCUMENT
        # ====================================================

        grouped = (

            store_data

            .groupby(

                [

                    "STORE",

                    "Last Updated By",

                    "Operation Type"

                ],

                sort=False,

                dropna=False

            )

            .agg(

                **{

                    "Distinct Count of Source Document":

                        (

                            "Source Document",

                            "nunique"

                        ),

                    "Sum of Demand Qty":

                        (

                            "Demand Qty",

                            "sum"

                        )

                }

            )

            .reset_index()

        )


        # ====================================================
        # OPERATION SORT
        # ====================================================

        grouped["_operation_order"] = (

            grouped["Operation Type"]

            .str.lower()

            .map(operation_order)

            .fillna(99)

        )


        grouped = grouped.sort_values(

            by=[

                "Last Updated By",

                "_operation_order"

            ],

            kind="stable"

        )


        grouped = grouped.drop(

            columns=[

                "_operation_order"

            ]

        )


        # ====================================================
        # CREATE DISPLAY ROWS
        # ====================================================

        final_rows = []


        first_store_row = True


        # ----------------------------------------------------
        # GROUP BY PERSON
        # ----------------------------------------------------

        for person, person_df in grouped.groupby(

            "Last Updated By",

            sort=False

        ):

            first_person_row = True


            for _, row in person_df.iterrows():


                # ------------------------------------------------
                # STORE ONLY ON FIRST ROW
                # ------------------------------------------------

                if first_store_row:

                    display_store = store_name

                    first_store_row = False

                else:

                    display_store = ""


                # ------------------------------------------------
                # PERSON ONLY ON FIRST ROW
                # ------------------------------------------------

                if first_person_row:

                    display_person = person

                    first_person_row = False

                else:

                    display_person = ""


                # ------------------------------------------------
                # ADD DATA ROW
                # ------------------------------------------------

                final_rows.append({

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


        # ====================================================
        # STORE TOTAL
        #
        # DISTINCT COUNT = SOURCE DOCUMENT
        # ====================================================

        store_distinct_count = (

            store_data[

                "Source Document"

            ].nunique()

        )


        store_demand_total = (

            store_data[

                "Demand Qty"

            ].sum()

        )


        final_rows.append({

            "STORE":

                f"{store_name} Total",

            "Last Updated By":

                "",

            "Operation Type":

                "",

            "Distinct Count of Source Document":

                store_distinct_count,

            "Sum of Demand Qty":

                store_demand_total,

            "ROW_TYPE":

                "TOTAL"

        })


        # ====================================================
        # RETURN
        # ====================================================

        return pd.DataFrame(
            final_rows
        )


    # ========================================================
    # 19. CREATE MASTER REPORT
    #
    # ALL STORES IN ONE TABLE
    # ========================================================

    master_report_list = []


    for store_name in store_list:

        store_report = create_store_report(
            store_name
        )


        if store_report.empty:

            continue


        master_report_list.append(
            store_report
        )


    # ========================================================
    # 20. COMBINE ALL STORE REPORTS
    # ========================================================

    if master_report_list:

        master_report = pd.concat(

            master_report_list,

            ignore_index=True

        )

    else:

        master_report = pd.DataFrame(

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
    # 21. REMOVE HELPER COLUMN
    # ========================================================

    master_report_display = master_report.drop(

        columns=[

            "ROW_TYPE"

        ],

        errors="ignore"

    )


    # ========================================================
    # 22. GRAND TOTAL
    #
    # DISTINCT COUNT = SOURCE DOCUMENT
    # ========================================================

    grand_distinct_count = (

        filtered_data[

            "Source Document"

        ].nunique()

    )


    grand_demand_total = (

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
            grand_distinct_count,

        "Sum of Demand Qty":
            grand_demand_total

    }])


    # ========================================================
    # 23. ADD GRAND TOTAL
    # ========================================================

    master_report_display = pd.concat(

        [

            master_report_display,

            grand_total

        ],

        ignore_index=True

    )


    # ========================================================
    # 24. REPORT TITLE
    # ========================================================

    title_text = (

        "STORE WISE PICK COUNT"

    )


    valid_dates = (

        master_raw_data[

            "Last Updated Date"

        ]

        .dropna()

    )


    if not valid_dates.empty:

        latest_date = valid_dates.max()


        title_text = (

            "STORE WISE PICK COUNT "

            + latest_date.strftime("%B").upper()

            + " - "

            + str(latest_date.day)

        )


    # ========================================================
    # 25. WRITE EXCEL
    # ========================================================

    with pd.ExcelWriter(

        output_file,

        engine="openpyxl"

    ) as writer:


        # ====================================================
        # MASTER DATA
        #
        # COMPLETE SUMMARY
        # ====================================================

        master_report_display.to_excel(

            writer,

            sheet_name="Master Data",

            index=False,

            startrow=1

        )


        # ====================================================
        # INDIVIDUAL STORE SHEETS
        # ====================================================

        for store_name in store_list:


            store_report = create_store_report(

                store_name

            )


            # ------------------------------------------------
            # REMOVE HELPER COLUMN
            # ------------------------------------------------

            store_report = store_report.drop(

                columns=[

                    "ROW_TYPE"

                ],

                errors="ignore"

            )


            if store_report.empty:

                continue


            # ------------------------------------------------
            # EXCEL SHEET NAME
            # ------------------------------------------------

            sheet_name = str(
                store_name
            )[:31]


            store_report.to_excel(

                writer,

                sheet_name=sheet_name,

                index=False,

                startrow=1

            )


    # ========================================================
    # 26. OPEN WORKBOOK
    # ========================================================

    wb = load_workbook(

        output_file

    )


    # ========================================================
    # 27. PROFESSIONAL COLORS
    # ========================================================

    dark_blue_fill = PatternFill(

        "solid",

        fgColor="1F4E78"

    )


    yellow_fill = PatternFill(

        "solid",

        fgColor="FFF200"

    )


    # ========================================================
    # 28. PROFESSIONAL FONTS
    # ========================================================

    title_font = Font(

        bold=True,

        size=15,

        color="1F4E78"

    )


    header_font = Font(

        bold=True,

        size=11,

        color="FFFFFF"

    )


    normal_font = Font(

        size=10

    )


    total_font = Font(

        bold=True,

        size=10

    )


    # ========================================================
    # 29. PROFESSIONAL BORDERS
    # ========================================================

    thin_side = Side(

        style="thin",

        color="B7B7B7"

    )


    medium_side = Side(

        style="medium",

        color="808080"

    )


    normal_border = Border(

        left=thin_side,

        right=thin_side,

        top=thin_side,

        bottom=thin_side

    )


    total_border = Border(

        left=medium_side,

        right=medium_side,

        top=medium_side,

        bottom=medium_side

    )


    # ========================================================
    # 30. PROFESSIONAL COLUMN WIDTHS
    # ========================================================

    def set_professional_widths(ws):

        ws.column_dimensions["A"].width = 30

        ws.column_dimensions["B"].width = 34

        ws.column_dimensions["C"].width = 26

        ws.column_dimensions["D"].width = 45

        ws.column_dimensions["E"].width = 28


    # ========================================================
    # 31. PROFESSIONAL SHEET FORMAT FUNCTION
    # ========================================================

    def format_report_sheet(
        ws,
        report_title
    ):

        # ====================================================
        # TITLE
        # ====================================================

        ws.merge_cells(
            "A1:E1"
        )


        ws["A1"] = report_title


        ws["A1"].font = title_font


        ws["A1"].alignment = Alignment(

            horizontal="center",

            vertical="center"

        )


        ws["A1"].fill = yellow_fill


        ws["A1"].border = total_border


        ws.row_dimensions[1].height = 32


        # ====================================================
        # HEADER
        # ====================================================

        for cell in ws[2]:

            cell.fill = dark_blue_fill


            cell.font = header_font


            cell.alignment = Alignment(

                horizontal="center",

                vertical="center",

                wrap_text=True

            )


            cell.border = normal_border


        ws.row_dimensions[2].height = 34


        # ====================================================
        # DATA ROWS
        # ====================================================

        for row in ws.iter_rows(

            min_row=3,

            max_row=ws.max_row,

            min_col=1,

            max_col=5

        ):

            store_value = row[0].value


            # =================================================
            # TOTAL ROW
            # =================================================

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

                    cell.font = total_font

                    cell.border = total_border

                    cell.alignment = Alignment(

                        horizontal="center",

                        vertical="center",

                        wrap_text=True

                    )


                ws.row_dimensions[
                    row[0].row
                ].height = 26


            # =================================================
            # NORMAL DATA ROW
            # =================================================

            else:

                for cell in row:

                    cell.font = normal_font

                    cell.border = normal_border

                    cell.alignment = Alignment(

                        horizontal="center",

                        vertical="center",

                        wrap_text=True

                    )


                ws.row_dimensions[
                    row[0].row
                ].height = 24


        # ====================================================
        # NUMBER FORMATTING
        # ====================================================

        for row in ws.iter_rows(

            min_row=3,

            max_row=ws.max_row

        ):

            # -----------------------------------------------
            # DISTINCT COUNT
            # -----------------------------------------------

            row[3].number_format = "#,##0"


            # -----------------------------------------------
            # DEMAND QTY
            # -----------------------------------------------

            row[4].number_format = "#,##0"


        # ====================================================
        # COLUMN WIDTHS
        # ====================================================

        set_professional_widths(
            ws
        )


        # ====================================================
        # FREEZE HEADER
        # ====================================================

        ws.freeze_panes = "A3"


        # ====================================================
        # FILTER
        # ====================================================

        ws.auto_filter.ref = (

            f"A2:E{ws.max_row}"

        )


        # ====================================================
        # REMOVE GRIDLINES
        # ====================================================

        ws.sheet_view.showGridLines = False


        # ====================================================
        # PAGE SETTINGS
        # ====================================================

        ws.sheet_properties.pageSetUpPr.fitToPage = True


        ws.page_setup.fitToWidth = 1


        ws.page_setup.fitToHeight = 0


        ws.page_setup.orientation = "landscape"


        ws.page_margins.left = 0.25

        ws.page_margins.right = 0.25

        ws.page_margins.top = 0.5

        ws.page_margins.bottom = 0.5


    # ========================================================
    # 32. FORMAT MASTER DATA
    # ========================================================

    ws = wb["Master Data"]


    format_report_sheet(

        ws,

        title_text

    )


    # ========================================================
    # 33. FORMAT INDIVIDUAL STORE SHEETS
    # ========================================================

    for store_name in store_list:

        sheet_name = str(
            store_name
        )[:31]


        if sheet_name not in wb.sheetnames:

            continue


        ws = wb[sheet_name]


        store_title = (

            f"{store_name} - "

            "PICK / INTERNAL TRANSFER REPORT"

        )


        format_report_sheet(

            ws,

            store_title

        )


    # ========================================================
    # 34. SAVE FINAL FILE
    # ========================================================

    wb.save(

        output_file

    )


    # ========================================================
    # 35. SUCCESS MESSAGE
    # ========================================================

    print(
        "\n=============================================="
    )

    print(
        "REPORT CREATED SUCCESSFULLY"
    )

    print(
        "=============================================="
    )


    print(
        f"Output file: {output_file}"
    )


    print(
        "\nCreated sheets:"
    )


    print(
        " - Master Data"
    )


    for store_name in store_list:

        print(
            f" - {store_name}"
        )


    print(
        "\nMaster Data = ALL STORES SUMMARY"
    )


    print(
        "Individual sheets = STORE-WISE SUMMARY"
    )


    print(
        "Distinct Count = Source Document"
    )


    print(
        "Demand = Sum of Demand Qty"
    )


    print(
        "Professional width / height applied"
    )


    print(
        "=============================================="
    )


    return output_file
