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
# STORE WISE PICK COUNT <MONTH> - <DAY>
# Example:
# STORE WISE PICK COUNT SEPTEMBER - 29
#
# Date comes from:
# Source Document Date
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
# Source Document Date is used.
#
# Example:
# STORE WISE PICK COUNT SEPTEMBER - 29
# ============================================================

def create_report_title(data):

    default_title = "STORE WISE PICK COUNT REPORT"

    if (
        data is None
        or data.empty
        or "Source Document Date" not in data.columns
    ):
        return default_title

    dates = pd.to_datetime(
        data["Source Document Date"],
        errors="coerce",
        dayfirst=True
    )

    dates = dates.dropna()

    if dates.empty:
        return default_title

    latest_date = dates.max()

    return (
        "STORE WISE PICK COUNT "
        f"{latest_date.strftime('%B').upper()} "
        f"- {latest_date.day}"
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
    # --------------------------------------------------------

    required_columns = [
        "Source Document",
        "Operation Type",
        "Demand Qty",
        "State",
        "Last Updated By",
        "Source Document Date"
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

    # --------------------------------------------------------
    # Source Document Date
    # --------------------------------------------------------

    df["Source Document Date"] = pd.to_datetime(
        df["Source Document Date"],
        errors="coerce",
        dayfirst=True
    )

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
    # STORE WISE PICK COUNT SEPTEMBER - 29
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

    ws.row_dimensions[1].height = 30

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

    ws.row_dimensions[2].height = 35

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

        # ----------------------------------------------------
        # NORMAL ROW
        # ----------------------------------------------------

        else:

            for cell in row:

                cell.border = border

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

    # ========================================================
    # NUMBER FORMAT
    # ========================================================

    for row in ws.iter_rows(
        min_row=3,
        max_row=ws.max_row
    ):

        row[3].number_format = "0"
        row[4].number_format = "0"

    # ========================================================
    # COLUMN WIDTHS
    # ========================================================

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["C"].width = 23
    ws.column_dimensions["D"].width = 43
    ws.column_dimensions["E"].width = 30

    # ========================================================
    # ROW HEIGHT
    # ========================================================

    for row_number in range(
        3,
        ws.max_row + 1
    ):

        ws.row_dimensions[
            row_number
        ].height = 30

    # ========================================================
    # FREEZE HEADER
    # ========================================================

    ws.freeze_panes = "A3"

    # ========================================================
    # AUTO FILTER
    # ========================================================

    if ws.max_row >= 2:

        ws.auto_filter.ref = (
            f"A2:E{ws.max_row}"
        )


# ============================================================
# 7. MAIN FUNCTION
# ============================================================

def generate_store_pick_report(
    file_path,
    output_file
):

    # ========================================================
    # READ ALL SHEETS
    #
    # sheet_name=None returns all worksheets as a dictionary.
    # ========================================================

    all_sheets = pd.read_excel(
        file_path,
        sheet_name=None,
        skiprows=4
    )

    print("\n==============================================")
    print("ALL SHEETS FOUND")
    print("==============================================")

    for sheet_name in all_sheets.keys():

        print(
            " -",
            sheet_name
        )

    # ========================================================
    # STORE DATA DICTIONARY
    # ========================================================

    store_filtered_data = {}

    # ========================================================
    # PROCESS EVERY SHEET
    # ========================================================

    for sheet_name, df in all_sheets.items():

        # ----------------------------------------------------
        # EXCLUDED SHEET
        # ----------------------------------------------------

        if sheet_name in EXCLUDED_SHEETS:

            print(
                f"\nSKIPPED: {sheet_name}"
            )

            continue

        print(
            f"\nProcessing: {sheet_name}"
        )

        # ----------------------------------------------------
        # FILTER THIS STORE
        # ----------------------------------------------------

        filtered_df = filter_store_data(
            df,
            sheet_name
        )

        # ----------------------------------------------------
        # INVALID SHEET
        # ----------------------------------------------------

        if filtered_df is None:

            continue

        # ----------------------------------------------------
        # SAVE STORE DATA
        # ----------------------------------------------------

        store_filtered_data[
            sheet_name
        ] = filtered_df

        print(
            "Filtered rows:",
            len(filtered_df)
        )

        # ----------------------------------------------------
        # OPERATION COUNTS
        # ----------------------------------------------------

        if not filtered_df.empty:

            pick_count = (
                filtered_df[
                    filtered_df[
                        "Operation Type"
                    ].str.lower()
                    == "pick"
                ]
                .shape[0]
            )

            internal_count = (
                filtered_df[
                    filtered_df[
                        "Operation Type"
                    ].str.lower()
                    == "internal transfers"
                ]
                .shape[0]
            )

            print(
                f"Pick rows: {pick_count}"
            )

            print(
                f"Internal Transfer rows: "
                f"{internal_count}"
            )

        else:

            print("Pick rows: 0")
            print("Internal Transfer rows: 0")

    # ========================================================
    # VALID STORES WITH DATA
    # ========================================================

    valid_store_data = {

        store: data

        for store, data
        in store_filtered_data.items()

        if not data.empty
    }

    if not valid_store_data:

        raise ValueError(
            "No State = Done Pick/Internal Transfer "
            "data found in any valid sheet."
        )

    # ========================================================
    # COMBINE ALL FILTERED STORE DATA
    # ========================================================

    combined_filtered_data = pd.concat(
        valid_store_data.values(),
        ignore_index=True
    )

    # ========================================================
    # CREATE FULL COMBINED REPORT
    # ========================================================

    combined_report = create_store_report(
        combined_filtered_data,
        include_grand_total=True
    )

    combined_display = combined_report.drop(
        columns=["ROW_TYPE"]
    )

    # ========================================================
    # COMBINED REPORT TITLE
    #
    # Uses latest Source Document Date from ALL
    # filtered stores.
    # ========================================================

    combined_title = create_report_title(
        combined_filtered_data
    )

    # ========================================================
    # WRITE EXCEL
    # ========================================================

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        # ====================================================
        # FULL COMBINED REPORT
        #
        # Row 1 = title
        # Row 2 = headers
        # ====================================================

        combined_display.to_excel(
            writer,
            sheet_name="Store Wise Report",
            index=False,
            startrow=1
        )

        # ====================================================
        # INDIVIDUAL STORE SHEETS
        # ====================================================

        used_sheet_names = {
            "Store Wise Report"
        }

        for store, store_df in valid_store_data.items():

            # -----------------------------------------------
            # CREATE STORE REPORT
            # -----------------------------------------------

            individual_report = create_store_report(
                store_df,
                include_grand_total=False
            )

            individual_display = (
                individual_report
                .drop(columns=["ROW_TYPE"])
            )

            # -----------------------------------------------
            # STORE TITLE
            #
            # Uses Source Document Date from that store.
            # -----------------------------------------------

            store_title = create_report_title(
                store_df
            )

            # -----------------------------------------------
            # CLEAN EXCEL SHEET NAME
            # -----------------------------------------------

            base_name = clean_sheet_name(
                store
            )

            sheet_name = base_name

            counter = 1

            while sheet_name in used_sheet_names:

                suffix = f"_{counter}"

                sheet_name = (
                    base_name[
                        :31 - len(suffix)
                    ]
                    + suffix
                )

                counter += 1

            used_sheet_names.add(
                sheet_name
            )

            # -----------------------------------------------
            # WRITE STORE REPORT
            # -----------------------------------------------

            individual_display.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                startrow=1
            )

    # ========================================================
    # OPEN WORKBOOK
    # ========================================================

    wb = load_workbook(
        output_file
    )

    # ========================================================
    # FORMAT COMBINED SHEET
    # ========================================================

    ws = wb["Store Wise Report"]

    format_report_sheet(
        ws,
        combined_title
    )

    # ========================================================
    # FORMAT EACH STORE SHEET
    # ========================================================

    for store, store_df in valid_store_data.items():

        # Find matching worksheet safely
        possible_name = clean_sheet_name(
            store
        )

        matching_ws = None

        for ws in wb.worksheets:

            if (
                ws.title != "Store Wise Report"
                and
                ws.title == possible_name
            ):
                matching_ws = ws
                break

        # If duplicate/suffixed sheet name was used,
        # find by position/content instead.
        if matching_ws is None:

            for ws in wb.worksheets:

                if ws.title == "Store Wise Report":
                    continue

                if (
                    ws.max_row >= 2
                    and
                    ws["A2"].value == "STORE"
                ):

                    # Check whether this sheet contains
                    # the current store name.
                    found_store = False

                    for row in range(
                        3,
                        min(ws.max_row, 20) + 1
                    ):

                        value = ws.cell(
                            row=row,
                            column=1
                        ).value

                        if (
                            value
                            and
                            (
                                str(value) == store
                                or
                                str(value) == f"{store} Total"
                            )
                        ):

                            found_store = True
                            break

                    if found_store:

                        matching_ws = ws
                        break

        if matching_ws is not None:

            store_title = create_report_title(
                store_df
            )

            format_report_sheet(
                matching_ws,
                store_title
            )

    # ========================================================
    # SAVE FINAL FILE
    # ========================================================

    wb.save(
        output_file
    )

    # ========================================================
    # SUCCESS MESSAGE
    # ========================================================

    print("\n==============================================")
    print("REPORT CREATED SUCCESSFULLY")
    print("==============================================")

    print(
        "\nOutput:",
        output_file
    )

    print(
        "\nCombined sheet:"
    )

    print(
        " - Store Wise Report"
    )

    print(
        "\nIndividual store sheets:"
    )

    for store in valid_store_data.keys():

        print(
            " -",
            store
        )

    print(
        "\nFiltering:"
    )

    print(
        " - State = Done"
    )

    print(
        " - Operation = Pick"
    )

    print(
        " - Operation = Internal Transfers"
    )

    print(
        "\nHeading date:"
    )

    print(
        " - Source Document Date"
    )

    print(
        "\nExample heading:"
    )

    print(
        " - STORE WISE PICK COUNT SEPTEMBER - 29"
    )

    print(
        "\nEach store shows:"
    )

    print(
        " - Person-wise Pick"
    )

    print(
        " - Person-wise Internal Transfers"
    )

    print(
        " - Store Total"
    )

    print(
        "\nCombined report also includes:"
    )

    print(
        " - Grand Total"
    )

    print("==============================================")

    return output_file
