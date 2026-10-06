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
# IMPORTANT:
# - Source Document Date is NOT considered
# - Last Updated Date is NOT considered
# - NO DATE FILTER
# - Today's date is used ONLY for the report heading
#
# ============================================================


import os
import re
import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# 1. INPUT / OUTPUT FILE
# ============================================================

INPUT_FILE = "Inventory Pending (62).xlsx"

REPORT_DATE = pd.Timestamp.today().normalize()

OUTPUT_FILE = (
    f"Store_Wise_Pick_Report_"
    f"{REPORT_DATE.strftime('%Y-%m-%d')}.xlsx"
)


# ============================================================
# 2. EXCLUDED SHEETS
# ============================================================

EXCLUDED_SHEETS = [
    "SPM_Virtual_Store",
    "Export(SPM)",
    "Virtual_Store_AD"
]


# ============================================================
# 3. REQUIRED COLUMNS
# ============================================================

REQUIRED_COLUMNS = [
    "Source Document",
    "Operation Type",
    "Demand Qty",
    "State",
    "Last Updated By"
]


# ============================================================
# 4. FIND HEADER ROW
#
# Your current Excel has the actual headers on ROW 5.
# This function automatically finds the row containing
# "Source Document", so it does not depend on a fixed row.
# ============================================================

def find_header_row(file_path, sheet_name):

    preview = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        header=None,
        nrows=30
    )

    for row_number in range(len(preview)):

        row_values = (
            preview.iloc[row_number]
            .fillna("")
            .astype(str)
            .str.strip()
            .tolist()
        )

        if "Source Document" in row_values:

            return row_number

    return None


# ============================================================
# 5. CLEAN SHEET NAME
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

    # Excel sheet names maximum = 31 characters
    return name[:31]


# ============================================================
# 6. CREATE REPORT TITLE
#
# SOURCE DOCUMENT DATE IS NOT USED HERE.
#
# Example:
#
# STORE WISE PICK COUNT OCTOBER - 6
# ============================================================

def create_report_title():

    report_date = pd.Timestamp.today()

    return (
        "STORE WISE PICK COUNT "
        f"{report_date.strftime('%B').upper()} "
        f"- {report_date.day}"
    )


# ============================================================
# 7. CLEAN COLUMN NAMES
# ============================================================

def clean_column_names(df):

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    return df


# ============================================================
# 8. NORMALIZE TEXT
# ============================================================

def clean_text_column(df, column):

    df[column] = (
        df[column]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    return df


# ============================================================
# 9. FILTER ONE STORE
#
# IMPORTANT:
# NO DATE FILTER HERE.
#
# Source Document Date is completely ignored.
# Last Updated Date is also completely ignored.
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

    df = clean_column_names(df)

    # --------------------------------------------------------
    # Check required columns
    # --------------------------------------------------------

    missing_columns = [
        col
        for col in REQUIRED_COLUMNS
        if col not in df.columns
    ]

    if missing_columns:

        print()
        print("=" * 70)
        print(f"SKIPPED STORE: {store_name}")
        print("Missing columns:")
        print(missing_columns)
        print("=" * 70)

        return None

    # ========================================================
    # CLEAN REQUIRED COLUMNS
    # ========================================================

    # --------------------------------------------------------
    # State
    # --------------------------------------------------------

    df = clean_text_column(
        df,
        "State"
    )

    # --------------------------------------------------------
    # Operation Type
    # --------------------------------------------------------

    df = clean_text_column(
        df,
        "Operation Type"
    )

    # --------------------------------------------------------
    # Last Updated By
    # --------------------------------------------------------

    df = clean_text_column(
        df,
        "Last Updated By"
    )

    # --------------------------------------------------------
    # Source Document
    # --------------------------------------------------------

    df = clean_text_column(
        df,
        "Source Document"
    )

    # --------------------------------------------------------
    # Convert blank Source Documents to NA
    #
    # This prevents blank Source Documents from being counted
    # as one distinct document.
    # --------------------------------------------------------

    df["Source Document"] = (
        df["Source Document"]
        .replace("", pd.NA)
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
    # Source Document Date is NOT used.
    # Last Updated Date is NOT used.
    # ========================================================

    filtered_df = df[
        (
            df["State"]
            .str.lower()
            .eq("done")
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

    # ========================================================
    # ADD STORE COLUMN
    # ========================================================

    filtered_df.insert(
        0,
        "STORE",
        store_name
    )

    return filtered_df


# ============================================================
# 10. CREATE REPORT DATA
# ============================================================

def create_store_report(
    filtered_data,
    include_grand_total=False
):

    # --------------------------------------------------------
    # Empty report
    # --------------------------------------------------------

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
    # GROUP DATA
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
                    (
                        "Source Document",
                        lambda x: x.dropna().nunique()
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

    # ========================================================
    # SORT
    # ========================================================

    grouped = grouped.sort_values(
        by=[
            "STORE",
            "Last Updated By",
            "_operation_order",
            "Operation Type"
        ],
        ascending=[
            True,
            True,
            True,
            True
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

    # --------------------------------------------------------
    # Each store
    # --------------------------------------------------------

    for store, store_df in grouped.groupby(
        "STORE",
        sort=False
    ):

        first_store_row = True

        # ----------------------------------------------------
        # Each employee
        # ----------------------------------------------------

        for person, person_df in store_df.groupby(
            "Last Updated By",
            sort=False
        ):

            first_person_row = True

            # ------------------------------------------------
            # Each operation
            # ------------------------------------------------

            for _, row in person_df.iterrows():

                # --------------------------------------------
                # STORE NAME
                # --------------------------------------------

                if first_store_row:

                    display_store = store

                    first_store_row = False

                else:

                    display_store = ""

                # --------------------------------------------
                # EMPLOYEE NAME
                # --------------------------------------------

                if first_person_row:

                    display_person = person

                    first_person_row = False

                else:

                    display_person = ""

                # --------------------------------------------
                # ADD DATA ROW
                # --------------------------------------------

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
                        float(
                            row["Sum of Demand Qty"]
                        ),

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
                    ]
                    .dropna()
                    .nunique()
                ),

            "Sum of Demand Qty":
                float(
                    store_data[
                        "Demand Qty"
                    ].sum()
                ),

            "ROW_TYPE":
                "TOTAL"
        })

    # ========================================================
    # GRAND TOTAL
    #
    # ONLY FOR COMBINED REPORT
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
                    ]
                    .dropna()
                    .nunique()
                ),

            "Sum of Demand Qty":
                float(
                    filtered_data[
                        "Demand Qty"
                    ].sum()
                ),

            "ROW_TYPE":
                "GRAND_TOTAL"
        })

    return pd.DataFrame(result)


# ============================================================
# 11. FORMAT REPORT SHEET
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

        row_type = ws.cell(
            row=row_num,
            column=6
        ).value

        # ----------------------------------------------------
        # DATA / TOTAL formatting
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
        # GRAND TOTAL
        # ----------------------------------------------------

        elif row_type == "GRAND_TOTAL":

            for col_num in range(1, 6):

                cell = ws.cell(
                    row=row_num,
                    column=col_num
                )

                cell.fill = yellow_fill

                cell.font = Font(
                    bold=True,
                    size=11
                )

        # ----------------------------------------------------
        # Normal data row
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
    # HIDE ROW_TYPE COLUMN
    # ========================================================

    ws.column_dimensions["F"].hidden = True

    # ========================================================
    # COLUMN WIDTHS
    # ========================================================

    ws.column_dimensions["A"].width = 25

    ws.column_dimensions["B"].width = 25

    ws.column_dimensions["C"].width = 23

    ws.column_dimensions["D"].width = 35

    ws.column_dimensions["E"].width = 22

    # ========================================================
    # FREEZE HEADER
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
    # PAGE / VIEW SETTINGS
    # ========================================================

    ws.sheet_view.showGridLines = False

    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws.page_setup.fitToWidth = 1

    ws.page_setup.fitToHeight = 0

    ws.page_setup.orientation = "landscape"


# ============================================================
# 12. WRITE REPORT DATA INTO EXCEL
# ============================================================

def write_report_to_sheet(
    ws,
    report_df,
    title_text
):

    # --------------------------------------------------------
    # Write report rows
    # --------------------------------------------------------

    output_columns = [
        "STORE",
        "Last Updated By",
        "Operation Type",
        "Distinct Count of Source Document",
        "Sum of Demand Qty",
        "ROW_TYPE"
    ]

    for row_index, row in enumerate(
        report_df[output_columns].itertuples(
            index=False,
            name=None
        ),
        start=3
    ):

        for col_index, value in enumerate(
            row,
            start=1
        ):

            ws.cell(
                row=row_index,
                column=col_index
            ).value = value

    # --------------------------------------------------------
    # Format
    # --------------------------------------------------------

    format_report_sheet(
        ws,
        title_text
    )


# ============================================================
# 13. READ ONE STORE SHEET
# ============================================================

def process_input_sheet(
    input_file,
    sheet_name
):

    print()
    print("-" * 70)
    print(f"Processing: {sheet_name}")

    # --------------------------------------------------------
    # Find actual header row
    # --------------------------------------------------------

    header_row = find_header_row(
        input_file,
        sheet_name
    )

    if header_row is None:

        print(
            f"SKIPPED: Header row not found in {sheet_name}"
        )

        return None

    print(
        f"Header found at Excel row: {header_row + 1}"
    )

    # --------------------------------------------------------
    # Read actual data
    #
    # pandas header is zero-based.
    # --------------------------------------------------------

    df = pd.read_excel(
        input_file,
        sheet_name=sheet_name,
        header=header_row
    )

    # --------------------------------------------------------
    # Filter
    # --------------------------------------------------------

    filtered_df = filter_store_data(
        df,
        sheet_name
    )

    if filtered_df is None:

        return None

    print(
        f"Rows after filter: {len(filtered_df)}"
    )

    return filtered_df


# ============================================================
# 14. MAIN PROGRAM
# ============================================================

def main():

    print()
    print("=" * 70)
    print("STORE WISE PICK / INTERNAL TRANSFER REPORT")
    print("=" * 70)

    print(
        f"Input file : {INPUT_FILE}"
    )

    print(
        f"Output file: {OUTPUT_FILE}"
    )

    print(
        f"Report date: {REPORT_DATE.strftime('%d-%m-%Y')}"
    )

    print()
    print(
        "IMPORTANT: Source Document Date is NOT considered."
    )

    print(
        "IMPORTANT: Last Updated Date is NOT considered."
    )

    print(
        "IMPORTANT: NO DATE FILTER is applied."
    )

    # ========================================================
    # CHECK INPUT FILE
    # ========================================================

    if not os.path.exists(INPUT_FILE):

        print()
        print(
            f"ERROR: Input file not found:"
        )

        print(
            INPUT_FILE
        )

        return

    # ========================================================
    # GET SHEET NAMES
    # ========================================================

    excel_file = pd.ExcelFile(
        INPUT_FILE
    )

    sheet_names = excel_file.sheet_names

    # ========================================================
    # FILTER EXCLUDED SHEETS
    # ========================================================

    process_sheets = [
        sheet
        for sheet in sheet_names
        if sheet not in EXCLUDED_SHEETS
    ]

    print()
    print(
        "Sheets to process:"
    )

    for sheet in process_sheets:

        print(
            f"  - {sheet}"
        )

    # ========================================================
    # PROCESS ALL STORES
    # ========================================================

    all_store_data = []

    store_data_dict = {}

    for sheet_name in process_sheets:

        filtered_df = process_input_sheet(
            INPUT_FILE,
            sheet_name
        )

        if (
            filtered_df is not None
            and not filtered_df.empty
        ):

            all_store_data.append(
                filtered_df
            )

            store_data_dict[
                sheet_name
            ] = filtered_df

        else:

            print(
                f"No matching records found for: {sheet_name}"
            )

    # ========================================================
    # CREATE COMBINED DATA
    # ========================================================

    if all_store_data:

        combined_data = pd.concat(
            all_store_data,
            ignore_index=True
        )

    else:

        combined_data = pd.DataFrame(
            columns=[
                "STORE",
                "Source Document",
                "Operation Type",
                "Demand Qty",
                "State",
                "Last Updated By"
            ]
        )

    # ========================================================
    # CREATE COMBINED REPORT
    # ========================================================

    combined_report = create_store_report(
        combined_data,
        include_grand_total=True
    )

    # ========================================================
    # CREATE EXCEL FILE
    # ========================================================

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        # ====================================================
        # 1. COMBINED REPORT
        # ====================================================

        combined_report.to_excel(
            writer,
            sheet_name="Store Wise Report",
            index=False,
            startrow=1
        )

        # ====================================================
        # 2. INDIVIDUAL STORE SHEETS
        # ====================================================

        for store_name, store_df in store_data_dict.items():

            store_report = create_store_report(
                store_df,
                include_grand_total=False
            )

            sheet_name = clean_sheet_name(
                store_name
            )

            # Excel cannot have duplicate sheet names
            # after cleaning.
            original_sheet_name = sheet_name

            counter = 1

            while sheet_name in writer.book.sheetnames:

                counter += 1

                suffix = f"_{counter}"

                sheet_name = (
                    original_sheet_name[:31 - len(suffix)]
                    + suffix
                )

            store_report.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                startrow=1
            )

    # ========================================================
    # FORMAT ALL SHEETS
    # ========================================================

    wb = load_workbook(
        OUTPUT_FILE
    )

    title_text = create_report_title()

    # ========================================================
    # FORMAT COMBINED REPORT
    # ========================================================

    if "Store Wise Report" in wb.sheetnames:

        ws = wb[
            "Store Wise Report"
        ]

        format_report_sheet(
            ws,
            title_text
        )

    # ========================================================
    # FORMAT STORE SHEETS
    # ========================================================

    for sheet_name in wb.sheetnames:

        if sheet_name == "Store Wise Report":
            continue

        ws = wb[
            sheet_name
        ]

        format_report_sheet(
            ws,
            title_text
        )

    # ========================================================
    # SAVE
    # ========================================================

    wb.save(
        OUTPUT_FILE
    )

    # ========================================================
    # SUMMARY
    # ========================================================

    print()
    print("=" * 70)
    print("REPORT CREATED SUCCESSFULLY")
    print("=" * 70)

    print(
        f"Output: {OUTPUT_FILE}"
    )

    print(
        f"Stores processed: {len(store_data_dict)}"
    )

    print(
        f"Total filtered records: {len(combined_data)}"
    )

    print()
    print(
        "Report heading:"
    )

    print(
        create_report_title()
    )

    print()
    print(
        "Source Document Date was NOT considered."
    )

    print(
        "No date filter was applied."
    )

    print("=" * 70)


# ============================================================
# 15. RUN
# ============================================================

if __name__ == "__main__":

    main()
