# ============================================================
# STORE WISE PICK / INTERNAL TRANSFER REPORT
#
# OUTPUT:
#   1. MASTER DATA
#   2. ONE SEPARATE SHEET FOR EACH STORE
#
# DISTINCT COUNT = SOURCE DOCUMENT
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
    # 3. STORE DATA
    # ============================================================

    all_data = []


    # ============================================================
    # 4. READ EACH STORE SHEET
    # ============================================================

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
            f"Reading store: {sheet_name}"
        )


        # --------------------------------------------------------
        # REMOVE EMPTY ROWS
        # --------------------------------------------------------

        df = df.dropna(
            how="all"
        ).copy()


        # --------------------------------------------------------
        # REMOVE EMPTY COLUMNS
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
        # --------------------------------------------------------

        df.insert(
            0,
            "STORE",
            sheet_name
        )


        # --------------------------------------------------------
        # ADD TO MASTER LIST
        # --------------------------------------------------------

        all_data.append(
            df
        )


    # ============================================================
    # 5. CHECK DATA
    # ============================================================

    if not all_data:

        raise ValueError(
            "No valid store sheets found."
        )


    # ============================================================
    # 6. COMBINE ALL STORE DATA
    # ============================================================

    master_data = pd.concat(
        all_data,
        ignore_index=True
    )


    # ============================================================
    # 7. CLEAN COLUMN NAMES
    # ============================================================

    master_data.columns = (
        master_data.columns
        .astype(str)
        .str.strip()
    )


    # ============================================================
    # 8. REQUIRED COLUMNS
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
        if col not in master_data.columns

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


        for col in master_data.columns:

            print(
                " -",
                col
            )


        raise ValueError(
            "Required columns are missing from the Excel file."
        )


    # ============================================================
    # 9. CLEAN STATE
    # ============================================================

    master_data["State"] = (

        master_data["State"]

        .fillna("")

        .astype(str)

        .str.strip()

    )


    # ============================================================
    # 10. CLEAN OPERATION TYPE
    # ============================================================

    master_data["Operation Type"] = (

        master_data["Operation Type"]

        .fillna("")

        .astype(str)

        .str.strip()

    )


    # ============================================================
    # 11. CLEAN LAST UPDATED BY
    # ============================================================

    master_data["Last Updated By"] = (

        master_data["Last Updated By"]

        .fillna("")

        .astype(str)

        .str.strip()

    )


    # ============================================================
    # 12. CLEAN SOURCE DOCUMENT
    #
    # THIS IS THE DISTINCT COUNT FIELD
    # ============================================================

    master_data["Source Document"] = (

        master_data["Source Document"]

        .fillna("")

        .astype(str)

        .str.strip()

    )


    # ============================================================
    # 13. CONVERT DEMAND QTY
    # ============================================================

    master_data["Demand Qty"] = pd.to_numeric(

        master_data["Demand Qty"],

        errors="coerce"

    ).fillna(0)


    # ============================================================
    # 14. CONVERT LAST UPDATED DATE
    #
    # USED ONLY FOR DATE / TITLE
    # ============================================================

    master_data["Last Updated Date"] = pd.to_datetime(

        master_data["Last Updated Date"],

        errors="coerce"

    )


    # ============================================================
    # 15. FILTER
    #
    # STATE = DONE
    #
    # OPERATION TYPE =
    # PICK
    # INTERNAL TRANSFERS
    # ============================================================

    filtered_data = master_data[

        (

            master_data["State"]

            .str.lower()

            == "done"

        )

        &

        (

            master_data["Operation Type"]

            .str.lower()

            .isin([

                "pick",

                "internal transfers"

            ])

        )

    ].copy()


    # ============================================================
    # 16. OPERATION ORDER
    #
    # INTERNAL TRANSFERS FIRST
    # PICK SECOND
    # ============================================================

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


    # ============================================================
    # 17. SORT FILTERED DATA
    # ============================================================

    filtered_data = filtered_data.sort_values(

        by=[

            "STORE",

            "Last Updated By",

            "_operation_order",

            "Operation Type"

        ],

        kind="stable"

    )


    # ============================================================
    # 18. REMOVE HELPER COLUMN
    # ============================================================

    filtered_data = filtered_data.drop(

        columns=[

            "_operation_order"

        ]

    )


    # ============================================================
    # 19. CREATE STORE REPORT FUNCTION
    # ============================================================

    def create_store_report(store_name):

        # --------------------------------------------------------
        # GET ONLY CURRENT STORE
        # --------------------------------------------------------

        store_data = filtered_data[

            filtered_data["STORE"]

            == store_name

        ].copy()


        # --------------------------------------------------------
        # IF NO DATA
        # --------------------------------------------------------

        if store_data.empty:

            return pd.DataFrame(

                columns=[

                    "STORE",

                    "Last Updated By",

                    "Operation Type",

                    "Distinct Count of Source Document",

                    "Sum of Demand Qty"

                ]

            )


        # ========================================================
        # GROUP DATA
        #
        # STORE
        # LAST UPDATED BY
        # OPERATION TYPE
        #
        # DISTINCT COUNT = SOURCE DOCUMENT
        # ========================================================

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


        # ========================================================
        # SORT OPERATION
        # ========================================================

        grouped["_operation_order"] = (

            grouped["Operation Type"]

            .str.lower()

            .map(operation_order)

            .fillna(99)

        )


        grouped = grouped.sort_values(

            by=[

                "Last Updated By",

                "_operation_order",

                "Operation Type"

            ],

            kind="stable"

        )


        grouped = grouped.drop(

            columns=[

                "_operation_order"

            ]

        )


        # ========================================================
        # DISPLAY REPORT
        #
        # STORE NAME ONLY FIRST ROW
        # LAST UPDATED BY ONLY FIRST ROW
        # ========================================================

        final_rows = []


        first_store_row = True


        for person, person_df in grouped.groupby(

            "Last Updated By",

            sort=False

        ):

            first_person_row = True


            for _, row in person_df.iterrows():


                # ------------------------------------------------
                # STORE
                # ------------------------------------------------

                if first_store_row:

                    display_store = store_name

                    first_store_row = False

                else:

                    display_store = ""


                # ------------------------------------------------
                # LAST UPDATED BY
                # ------------------------------------------------

                if first_person_row:

                    display_person = person

                    first_person_row = False

                else:

                    display_person = ""


                # ------------------------------------------------
                # ADD ROW
                # ------------------------------------------------

                final_rows.append({

                    "STORE":
                        display_store,

                    "Last Updated By":
                        display_person,

                    "Operation Type":
                        row["Operation Type"],

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

        total_distinct_documents = (

            store_data["Source Document"]

            .nunique()

        )


        total_demand_qty = (

            store_data["Demand Qty"]

            .sum()

        )


        final_rows.append({

            "STORE":

                f"{store_name} Total",

            "Last Updated By":

                "",

            "Operation Type":

                "",

            "Distinct Count of Source Document":

                total_distinct_documents,

            "Sum of Demand Qty":

                total_demand_qty,

            "ROW_TYPE":

                "TOTAL"

        })


        # ========================================================
        # CREATE DATAFRAME
        # ========================================================

        report_df = pd.DataFrame(
            final_rows
        )


        # ========================================================
        # REMOVE HELPER COLUMN
        # ========================================================

        report_df = report_df.drop(

            columns=[

                "ROW_TYPE"

            ]

        )


        return report_df


    # ============================================================
    # 20. GET STORE LIST
    #
    # KEEP SAME ORDER AS ORIGINAL EXCEL SHEETS
    # ============================================================

    store_list = [

        sheet_name

        for sheet_name in all_sheets.keys()

        if sheet_name not in EXCLUDED_SHEETS

    ]


    # ============================================================
    # 21. GET REPORT DATE
    # ============================================================

    title_text = (

        "STORE WISE PICK / INTERNAL TRANSFER REPORT"

    )


    valid_dates = master_data[

        "Last Updated Date"

    ].dropna()


    if not valid_dates.empty:

        latest_date = valid_dates.max()


        title_text = (

            "STORE WISE PICK / INTERNAL TRANSFER REPORT - "

            + latest_date.strftime("%d-%m-%Y")

        )


    # ============================================================
    # 22. CREATE EXCEL FILE
    #
    # MASTER DATA
    # +
    # EACH STORE AS SEPARATE SHEET
    # ============================================================

    with pd.ExcelWriter(

        output_file,

        engine="openpyxl"

    ) as writer:


        # ========================================================
        # MASTER DATA SHEET
        # ========================================================

        master_data.to_excel(

            writer,

            sheet_name="Master Data",

            index=False

        )


        # ========================================================
        # ONE SHEET FOR EACH STORE
        # ========================================================

        for store_name in store_list:


            store_report = create_store_report(

                store_name

            )


            # ----------------------------------------------------
            # EXCEL SHEET NAME
            #
            # Excel maximum sheet name = 31 characters
            # ----------------------------------------------------

            sheet_name = str(store_name)[:31]


            store_report.to_excel(

                writer,

                sheet_name=sheet_name,

                index=False,

                startrow=1

            )


    # ============================================================
    # 23. OPEN WORKBOOK FOR FORMATTING
    # ============================================================

    wb = load_workbook(

        output_file

    )


    # ============================================================
    # 24. COLORS
    # ============================================================

    dark_blue_fill = PatternFill(

        "solid",

        fgColor="1F4E78"

    )


    blue_fill = PatternFill(

        "solid",

        fgColor="4472C4"

    )


    yellow_fill = PatternFill(

        "solid",

        fgColor="FFFF00"

    )


    light_blue_fill = PatternFill(

        "solid",

        fgColor="D9EAF7"

    )


    # ============================================================
    # 25. FONTS
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
    # 26. BORDER
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
    # 27. FORMAT MASTER DATA
    # ============================================================

    ws = wb["Master Data"]


    ws.freeze_panes = "A2"


    ws.auto_filter.ref = (

        f"A1:{get_column_letter(ws.max_column)}"

        f"{ws.max_row}"

    )


    for cell in ws[1]:

        cell.fill = blue_fill

        cell.font = white_font

        cell.alignment = Alignment(

            horizontal="center",

            vertical="center",

            wrap_text=True

        )

        cell.border = border


    # ------------------------------------------------------------
    # MASTER COLUMN WIDTH
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
    # 28. FORMAT EACH STORE SHEET
    # ============================================================

    for store_name in store_list:


        sheet_name = str(store_name)[:31]


        ws = wb[sheet_name]


        # ========================================================
        # TITLE
        # ========================================================

        ws.merge_cells(

            "A1:E1"

        )


        ws["A1"] = (

            f"{store_name} - "

            "PICK / INTERNAL TRANSFER REPORT"

        )


        ws["A1"].font = title_font


        ws["A1"].alignment = Alignment(

            horizontal="center",

            vertical="center"

        )


        ws["A1"].fill = light_blue_fill


        ws.row_dimensions[1].height = 25


        # ========================================================
        # HEADER
        # ========================================================

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


        # ========================================================
        # DATA ROWS
        # ========================================================

        for row in ws.iter_rows(

            min_row=3,

            max_row=ws.max_row,

            min_col=1,

            max_col=5

        ):

            store_value = row[0].value


            # ----------------------------------------------------
            # TOTAL ROW
            # ----------------------------------------------------

            if (

                store_value

                and

                str(store_value).endswith(

                    " Total"

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


            # ----------------------------------------------------
            # NORMAL ROW
            # ----------------------------------------------------

            else:

                for cell in row:

                    cell.border = border

                    cell.alignment = Alignment(

                        horizontal="center",

                        vertical="center",

                        wrap_text=True

                    )


        # ========================================================
        # NUMBER FORMATS
        # ========================================================

        for row in ws.iter_rows(

            min_row=3,

            max_row=ws.max_row

        ):

            # DISTINCT SOURCE DOCUMENT COUNT

            row[3].number_format = "0"


            # DEMAND QTY

            row[4].number_format = "0"


        # ========================================================
        # COLUMN WIDTHS
        # ========================================================

        ws.column_dimensions["A"].width = 28

        ws.column_dimensions["B"].width = 32

        ws.column_dimensions["C"].width = 23

        ws.column_dimensions["D"].width = 42

        ws.column_dimensions["E"].width = 25


        # ========================================================
        # ROW HEIGHT
        # ========================================================

        for row_number in range(

            3,

            ws.max_row + 1

        ):

            ws.row_dimensions[

                row_number

            ].height = 20


        # ========================================================
        # FREEZE
        # ========================================================

        ws.freeze_panes = "A3"


        # ========================================================
        # FILTER
        # ========================================================

        ws.auto_filter.ref = (

            f"A2:E{ws.max_row}"

        )


    # ============================================================
    # 29. SAVE FINAL FILE
    # ============================================================

    wb.save(

        output_file

    )


    # ============================================================
    # 30. SUCCESS MESSAGE
    # ============================================================

    print("\n")

    print(
        "=================================================="
    )

    print(
        "       REPORT CREATED SUCCESSFULLY"
    )

    print(
        "=================================================="
    )

    print(
        f"Output: {output_file}"
    )

    print(
        "\nCreated sheets:"
    )

    print(
        " - Master Data"
    )


    for store in store_list:

        print(
            f" - {store}"
        )


    print(
        "\nDistinct Count = Source Document"
    )

    print(
        "Last Updated By = Person"
    )

    print(
        "Operation Type = Pick / Internal Transfers"
    )

    print(
        "Each store has its own separate sheet."
    )

    print(
        "=================================================="
    )


    return output_file
