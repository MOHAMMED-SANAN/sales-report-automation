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
# 3. FILTER ONE SHEET
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
# 4. CREATE REPORT
#
# This creates EXACTLY this type of layout:
#
# STORE | Last Updated By | Operation Type | Count | Qty
#
# WH_Rashidiya(SPM) | CHHATRA | Internal Transfers | 1 | 8
#                    |         | Pick              | 3 | 9
#                    | NIRANJ  | Pick              | 62| 2031
# WH_Rashidiya(SPM) Total                         | 66| 2048
# ============================================================

def create_store_report(filtered_data, include_grand_total=False):

    if filtered_data is None or filtered_data.empty:

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
        #
        # Pick + Internal Transfers together
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
    #
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
# 5. FORMAT ONE REPORT SHEET
# ============================================================

def format_report_sheet(
    ws,
    is_combined=False
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
    # HEADER
    #
    # Header starts at ROW 1
    # Just like your screenshot
    # ========================================================

    for cell in ws[1]:

        cell.fill = dark_blue_fill

        cell.font = white_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        cell.border = border

    ws.row_dimensions[1].height = 35

    # ========================================================
    # DATA ROWS
    # ========================================================

    for row in ws.iter_rows(
        min_row=2,
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
        min_row=2,
        max_row=ws.max_row
    ):

        # Count
        row[3].number_format = "0"

        # Demand Qty
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
        2,
        ws.max_row + 1
    ):

        ws.row_dimensions[
            row_number
        ].height = 30

    # ========================================================
    # FREEZE HEADER
    # ========================================================

    ws.freeze_panes = "A2"

    # ========================================================
    # AUTO FILTER
    # ========================================================

    if ws.max_row >= 1:

        ws.auto_filter.ref = (
            f"A1:E{ws.max_row}"
        )


# ============================================================
# 6. MAIN FUNCTION
# ============================================================

def generate_store_pick_report(
    file_path,
    output_file
):

    # ========================================================
    # READ ALL SHEETS
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
        # EVEN IF EMPTY
        #
        # This allows us to know which sheets were processed.
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
                    ].str.lower() == "pick"
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
    # WRITE EXCEL
    #
    # FIRST:
    # Store Wise Report
    #
    # THEN:
    # Every individual store
    # ========================================================

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        # ====================================================
        # FULL COMBINED REPORT
        # ====================================================

        combined_display.to_excel(
            writer,
            sheet_name="Store Wise Report",
            index=False
        )

        # ====================================================
        # INDIVIDUAL STORE SHEETS
        # ====================================================

        used_sheet_names = {
            "Store Wise Report"
        }

        for store, store_df in valid_store_data.items():

            # -----------------------------------------------
            # Create this store's report
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
            # Clean Excel sheet name
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
            # Write individual store
            # -----------------------------------------------

            individual_display.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

    # ========================================================
    # OPEN WORKBOOK
    # ========================================================

    wb = load_workbook(
        output_file
    )

    # ========================================================
    # FORMAT EVERY SHEET
    # ========================================================

    for ws in wb.worksheets:

        format_report_sheet(
            ws,
            is_combined=(
                ws.title
                == "Store Wise Report"
            )
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
