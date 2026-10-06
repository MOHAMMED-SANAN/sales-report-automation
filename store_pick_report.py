```python
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
    # Source Document Date REMOVED
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

   
