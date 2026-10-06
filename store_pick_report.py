# ============================================================
# STORE WISE PICK / INTERNAL TRANSFER REPORT
#
# OUTPUT:
# 1. Store Wise Report -> ALL STORES COMBINED
# 2. One sheet for EACH STORE
#
# FILTER:
# - State = Done
# - Operation Type = Pick OR Internal Transfers
#
# HEADING:
# STORE WISE PICK COUNT <TODAY'S MONTH> - <TODAY'S DAY>
#
# IMPORTANT:
# - Source Document Date is NOT considered
# - Report heading uses today's date
# - NO DATE FILTER IS USED
# ============================================================


import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ============================================================
# 1. EXCLUDED SHEETS
# ============================================================

EXCLUDED_SHEETS = [
    "SPM_Virtual_Store",
    "Export(SPM)",
    "Virtual_Store_AD"
]


# ============================================================
# 2. EXCEL SHEET NAME CLEANER
# ============================================================

def clean_sheet_name(name):

    invalid_chars = [
        "/", "\\", "?", "*", "[", "]", ":"
    ]

    name = str(name)

    for char in invalid_chars:
        name = name.replace(char, "_")

    name = name.strip()

    if not name:
        name = "Store"

    return name[:31]


# ============================================================
# 3. CREATE REPORT TITLE
#
# IMPORTANT:
# Source Document Date is NOT used.
# Today's date is used instead.
#
# Example:
# STORE WISE PICK COUNT OCTOBER - 6
# ============================================================

def create_report_title(data=None):

    report_date = pd.Timestamp.today()

    return (
        "STORE WISE PICK COUNT "
        f"{report_date.strftime('%B').upper()} "
        f"- {report_date.day}"
    )


# ============================================================
# 4. FILTER ONE SHEET
# ============================================================

def filter_store_data(df, store_name):

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
    # Required columns
    #
    # Source Document Date is intentionally NOT included.
    # --------------------------------------------------------

    required_columns = [
        "Source Document",
        "Operation Type",
        "Demand Qty",
        "State",
        "Last Updated By"
    ]

    missing_columns = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:

        print(
            f"SKIPPED: {store_name}"
        )

        print(
            "Missing columns:",
            missing_columns
        )

        return None

    # --------------------------------------------------------
    # Clean State
    # --------------------------------------------------------

    df["State"] = (
        df["State"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # Clean Operation Type
    # --------------------------------------------------------

    df["Operation Type"] = (
        df["Operation Type"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # Clean Last Updated By
    # --------------------------------------------------------

    df["Last Updated By"] = (
        df["Last Updated By"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # --------------------------------------------------------
    # Demand Qty -> number
    # --------------------------------------------------------

    df["Demand Qty"] = pd.to_numeric(
        df["Demand Qty"],
        errors="coerce"
    ).fillna(0)

    # ========================================================
    # IMPORTANT FILTER
    #
    # State = DONE
    #
    # AND
    #
    # Operation Type =
    #     PICK
    #     INTERNAL TRANSFERS
    #
    # NO DATE FILTER
    #
    # Source Document Date is completely ignored.
    # ========================================================

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

    # --------------------------------------------------------
    # Add STORE column
    # Sheet name = Store
    # --------------------------------------------------------

    filtered_df.insert(
        0,
        "STORE",
        store_name
    )

    return filtered_df


# ============================================================
# 5. CREATE REPORT
#
# Example:
#
# STORE              Last Updated By     Operation Type
# WH_Rashidiya(SPM)  CHHATRA BHANDARA    Internal Transfers
#                                        Pick
#                    NIRANJ.K.R          Pick
#
# WH_Rashidiya(SPM) Total
# ============================================================

def create_store_report(
    filtered_data,
    include_grand_total=False
):

    if (
        filtered_data is None
        or filtered_data.empty
    ):

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
    # GROUP
    # ========================================================

    grouped = (
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
    #
    # Internal Transfers first
    # Pick second
    # ========================================================

    operation_order = {
        "internal transfers": 0,
        "pick": 1
    }

    grouped["_operation_order"] = (
        grouped["Operation Type"]
        .astype(str)
        .str.lower()
        .map(operation_order)
        .fillna(99)
    )

    grouped = grouped.sort_values(
        by=[
            "STORE",
            "Last Updated By",
            "_operation_order",
            "Operation Type"
        ],
        na_position="last"
    )

    grouped = grouped.drop(
        columns="_operation_order"
    )

    # ========================================================
    # BUILD DISPLAY DATA
    # ========================================================

    result = []

    for store, store_df in grouped.groupby(
        "STORE",
        sort=False
    ):

        first_store_row = True

        # ----------------------------------------------------
        # GROUP BY PERSON
        # ----------------------------------------------------

        for person, person_df in store_df.groupby(
            "Last Updated By",
            sort=False
        ):

            first_person_row = True

            # ------------------------------------------------
            # PICK / INTERNAL TRANSFER
            # ------------------------------------------------

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

                result.append({

                    "STORE":
                        display_store,

                    "Last Updated By":
                        display_person,

                    "Operation Type":
                        row["Operation Type"],

                    "Distinct Count of Source Document":
                        int(
                            row[
                                "Distinct Count of Source Document"
                            ]
                        ),

                    "Sum of Demand Qty":
                        row["Sum of Demand Qty"],

                    "ROW_TYPE":
                        "DATA"
                })

        # ====================================================
        # STORE TOTAL
        # ====================================================

        store_data = filtered_data[
            filtered_data["STORE"] == store
        ]

        result.append({

            "STORE":
                f"{store} Total",

            "Last Updated By":
                "",

            "Operation Type":
                "",

            "Distinct Count of Source Document":
                int(
                    store_data[
                        "Source Document"
                    ].nunique()
                ),

            "Sum of Demand Qty":
                store_data[
                    "Demand Qty"
                ].sum(),

            "ROW_TYPE":
                "TOTAL"
        })

    # ========================================================
    # GRAND TOTAL
    # ONLY FOR FULL COMBINED REPORT
    # ========================================================

    if include_grand_total:

        result.append({

            "STORE":
                "Grand Total",

            "Last Updated By":
                "",

            "Operation Type":
                "",

            "Distinct Count of Source Document":
                int(
                    filtered_data[
                        "Source Document"
                    ].nunique()
                ),

            "Sum of Demand Qty":
                filtered_data[
                    "Demand Qty"
                ].sum(),

            "ROW_TYPE":
                "GRAND_TOTAL"
        })

    return pd.DataFrame(result)


# ============================================================
# 6. FORMAT REPORT SHEET
#
# Row 1 = TITLE
# Row 2 = HEADER
# Row 3 onward = DATA
# ============================================================

def format_report_sheet(
    ws,
    title_text
):

    # ========================================================
    # COLORS
    # ========================================================

    dark_blue_fill = PatternFill(
        "solid",
        fgColor="1F4E78"
    )

    yellow_fill = PatternFill(
        "solid",
        fgColor="FFFF00"
    )

    white_font = Font(
        color="FFFFFF",
        bold=True
    )

    title_font = Font(
        bold=True,
        size=14,
        color="1F4E78"
    )

    bold_font = Font(
        bold=True
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
    # TITLE
    #
    # Example:
    # STORE WISE PICK COUNT OCTOBER - 6
    # ========================================================

    ws.merge_cells(
        "A1:E1"
    )

    ws["A1"] = title_text

    ws["A1"].fill = yellow_fill

    ws["A1"].font = title_font

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws["A1"].border = border

    ws.row_dimensions[1].height = 28

    # ========================================================
    # HEADER
    # ========================================================

    headers = [
        "STORE",
        "Last Updated By",
        "Operation Type",
        "Distinct Count of Source Document",
        "Sum of Demand Qty"
    ]

    for col_num, header in enumerate(
        headers,
        start=1
    ):

        cell = ws.cell(
            row=2,
            column=col_num
        )

        cell.value = header

        cell.fill = dark_blue_fill

        cell.font = white_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        cell.border = border

    ws.row_dimensions[2].height = 35

    # ========================================================
    # DATA ROWS
    # ========================================================

    for row_num in range(
        3,
        ws.max_row + 1
    ):

        # ROW_TYPE is in column F
        row_type = ws.cell(
            row=row_num,
            column=6
        ).value

        # ----------------------------------------------------
        # Normal borders / alignment
        # ----------------------------------------------------

        for col_num in range(1, 6):

            cell = ws.cell(
                row=row_num,
                column=col_num
            )

            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # ----------------------------------------------------
        # TOTAL ROW
        # ----------------------------------------------------

        if row_type == "TOTAL":

            for col_num in range(1, 6):

                cell = ws.cell(
                    row=row_num,
                    column=col_num
                )

                cell.fill = yellow_fill

                cell.font = bold_font

        # ----------------------------------------------------
        # GRAND TOTAL ROW
        # ----------------------------------------------------

        elif row_type == "GRAND_TOTAL":

            for col_num in range(1, 6):

                cell = ws.cell(
                    row=row_num,
                    column=col_num
                )

                cell.fill = yellow_fill

                cell.font = bold_font

        # ----------------------------------------------------
        # Normal data
        # ----------------------------------------------------

        else:

            for col_num in range(1, 6):

                ws.cell(
                    row=row_num,
                    column=col_num
                ).font = Font(
                    size=10
                )

    # ========================================================
    # COLUMN WIDTH
    # ========================================================

    ws.column_dimensions["A"].width = 24

    ws.column_dimensions["B"].width = 25

    ws.column_dimensions["C"].width = 23

    ws.column_dimensions["D"].width = 35

    ws.column_dimensions["E"].width = 22

    # ========================================================
    # HIDE ROW_TYPE
    # ========================================================

    ws.column_dimensions["F"].hidden = True

    # ========================================================
    # FREEZE
    # ========================================================

    ws.freeze_panes = "A3"

    # ========================================================
    # FILTER
    # ========================================================

    if ws.max_row >= 2:

        ws.auto_filter.ref = (
            f"A2:E{ws.max_row}"
        )

    # ========================================================
    # GRIDLINES
    # ========================================================

    ws.sheet_view.showGridLines = False


# ============================================================
# 7. PROCESS ONE INPUT SHEET
# ============================================================

def process_input_sheet(
    input_file,
    sheet_name
):

    print(
        f"Processing: {sheet_name}"
    )

    # --------------------------------------------------------
    # Read Excel sheet
    #
    # Change header=6 if your original file has header on row 7.
    # --------------------------------------------------------

    df = pd.read_excel(
        input_file,
        sheet_name=sheet_name,
        header=0
    )

    # --------------------------------------------------------
    # Filter
    # --------------------------------------------------------

    filtered_df = filter_store_data(
        df,
        sheet_name
    )

    return filtered_df


# ============================================================
# 8. MAIN
# ============================================================

def main():

    # ========================================================
    # INPUT FILE
    # ========================================================

    input_file = "Inventory Pending (62).xlsx"

    # ========================================================
    # OUTPUT FILE
    # ========================================================

    report_date = pd.Timestamp.today()

    output_file = (
        "Store_Wise_Pick_Report_"
        f"{report_date.strftime('%Y-%m-%d')}.xlsx"
    )

    # ========================================================
    # REPORT TITLE
    # ========================================================

    title_text = create_report_title()

    print()
    print(
        "=============================================="
    )

    print(
        "STORE WISE PICK / INTERNAL TRANSFER REPORT"
    )

    print(
        "=============================================="
    )

    print(
        f"Report Title: {title_text}"
    )

    print(
        "Source Document Date: NOT CONSIDERED"
    )

    print(
        "=============================================="
    )

    # ========================================================
    # READ ALL SHEETS
    # ========================================================

    excel_file = pd.ExcelFile(
        input_file
    )

    sheet_names = excel_file.sheet_names

    # ========================================================
    # ALL STORE DATA
    # ========================================================

    all_filtered_data = []

    store_filtered_data = {}

    # ========================================================
    # PROCESS EACH SHEET
    # ========================================================

    for sheet_name in sheet_names:

        # ----------------------------------------------------
        # Skip excluded sheets
        # ----------------------------------------------------

        if sheet_name in EXCLUDED_SHEETS:

            print(
                f"SKIPPED SHEET: {sheet_name}"
            )

            continue

        # ----------------------------------------------------
        # Process
        # ----------------------------------------------------

        filtered_df = process_input_sheet(
            input_file,
            sheet_name
        )

        # ----------------------------------------------------
        # Store only if data exists
        # ----------------------------------------------------

        if (
            filtered_df is not None
            and not filtered_df.empty
        ):

            all_filtered_data.append(
                filtered_df
            )

            store_filtered_data[
                sheet_name
            ] = filtered_df

    # ========================================================
    # COMBINE ALL STORES
    # ========================================================

    if all_filtered_data:

        combined_data = pd.concat(
            all_filtered_data,
            ignore_index=True
        )

    else:

        combined_data = pd.DataFrame()

    # ========================================================
    # CREATE COMBINED REPORT
    # ========================================================

    combined_report = create_store_report(
        combined_data,
        include_grand_total=True
    )

    # ========================================================
    # CREATE EXCEL
    # ========================================================

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        # ====================================================
        # STORE WISE REPORT
        # ====================================================

        combined_report.to_excel(
            writer,
            sheet_name="Store Wise Report",
            index=False,
            startrow=1
        )

        # ====================================================
        # INDIVIDUAL STORE SHEETS
        # ====================================================

        for store_name, store_df in store_filtered_data.items():

            store_report = create_store_report(
                store_df,
                include_grand_total=False
            )

            sheet_name = clean_sheet_name(
                store_name
            )

            store_report.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                startrow=1
            )

    # ========================================================
    # LOAD WORKBOOK FOR FORMATTING
    # ========================================================

    wb = load_workbook(
        output_file
    )

    # ========================================================
    # FORMAT ALL SHEETS
    # ========================================================

    for ws in wb.worksheets:

        format_report_sheet(
            ws,
            title_text
        )

    # ========================================================
    # SAVE
    # ========================================================

    wb.save(
        output_file
    )

    # ========================================================
    # COMPLETED
    # ========================================================

    print()
    print(
        "=============================================="
    )

    print(
        "REPORT CREATED SUCCESSFULLY"
    )

    print(
        "=============================================="
    )

    print(
        f"Output File: {output_file}"
    )

    print(
        f"Total Stores: {len(store_filtered_data)}"
    )

    print(
        f"Total Records: {len(combined_data)}"
    )

    print(
        "Source Document Date was NOT considered."
    )

    print(
        "No date filter was applied."
    )

    print(
        "=============================================="
    )


# ============================================================
# 9. RUN PROGRAM
# ============================================================

if __name__ == "__main__":

    main()
