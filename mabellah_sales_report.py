import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ==============================================================
# 1. READ AND CLEAN INPUT FILE
# ==============================================================

def salesman_newfile(file_path):

    df = pd.read_excel(
        file_path,
        header=6
    )

    # Clean column names
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    # ----------------------------------------------------------
    # STORE
    # ----------------------------------------------------------
    # File already contains MABELLAH.
    # No mapping is required.

    df["STORE"] = (
        df["STORE"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ----------------------------------------------------------
    # SALES PERSON
    # ----------------------------------------------------------

    df["SALES MAN"] = (
        df["SALESPERSON"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ----------------------------------------------------------
    # DATE
    # ----------------------------------------------------------

    df["DATE"] = pd.to_datetime(
        df["DATE"],
        dayfirst=True,
        errors="coerce"
    )

    # ----------------------------------------------------------
    # ENQUIRY STATUS
    # ----------------------------------------------------------

    df["ENQUIRY STATUS"] = (
        df["ENQUIRY STATUS"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # ----------------------------------------------------------
    # SUB TOTAL
    # ----------------------------------------------------------

    df["SUB TOTAL"] = pd.to_numeric(
        df["SUB TOTAL"],
        errors="coerce"
    ).fillna(0)

    # ----------------------------------------------------------
    # ONLY MABELLAH
    # ----------------------------------------------------------
    # No salesperson filter here.
    # Every salesperson in the SALES data will be included.

    df = df[
        df["STORE"] == "MABELLAH"
    ].copy()

    return df


# ==============================================================
# 2. CREATE SALES PERSON REPORT
# ==============================================================

def sales_person_report(df):

    # ----------------------------------------------------------
    # TARGET TABLE
    # ----------------------------------------------------------
    # Only SHIBILU has a target currently.
    # New salespeople will automatically appear with "-".

    sales_person_targets = {
        ("MABELLAH", "SHIBILU RAHMAN"): 25000
    }

    # ----------------------------------------------------------
    # FIRST SHEET USES ONLY ENQUIRY STATUS = SALES
    # ----------------------------------------------------------

    sales_df = df[
        df["ENQUIRY STATUS"] == "SALES"
    ].copy()

    branch_order = [
        "MABELLAH"
    ]

    # ----------------------------------------------------------
    # CHECK DATA
    # ----------------------------------------------------------

    if sales_df.empty:

        raise ValueError(
            "No SALES records found for MABELLAH."
        )

    # ----------------------------------------------------------
    # REPORT DATE
    # ----------------------------------------------------------

    report_date = (
        pd.to_datetime(
            sales_df["DATE"],
            errors="coerce"
        )
        .max()
    )

    if pd.isna(report_date):

        raise ValueError(
            "No valid DATE found in the filtered data."
        )

    report_date = report_date.normalize()

    # ----------------------------------------------------------
    # MONTH
    # ----------------------------------------------------------

    month_start = report_date.replace(day=1)

    month_end = (
        report_date
        + pd.offsets.MonthEnd(0)
    )

    # ----------------------------------------------------------
    # DATE RANGE
    # ----------------------------------------------------------

    all_days = pd.date_range(
        start=month_start,
        end=month_end
    )

    days_till_now = pd.date_range(
        start=month_start,
        end=report_date
    )

    # ----------------------------------------------------------
    # WORKING DAYS
    # ----------------------------------------------------------
    # MABELLAH:
    # Friday = holiday
    # Sunday = working day

    branch_working_days = {}
    branch_working_days_till_now = {}

    for branch in branch_order:

        if branch == "MABELLAH":

            working_days = (
                all_days.dayofweek != 4
            ).sum()

            working_days_till_now = (
                days_till_now.dayofweek != 4
            ).sum()

        else:

            working_days = len(all_days)
            working_days_till_now = len(days_till_now)

        branch_working_days[branch] = working_days

        branch_working_days_till_now[branch] = (
            working_days_till_now
        )

    # ----------------------------------------------------------
    # SALES BY SALESPERSON
    # ----------------------------------------------------------

    sales = (
        sales_df.groupby(
            [
                "STORE",
                "SALES MAN"
            ],
            as_index=False
        )["SUB TOTAL"]
        .sum()
        .rename(
            columns={
                "SUB TOTAL": "ACHIEVED SALES"
            }
        )
    )

    # ----------------------------------------------------------
    # REPORT ROWS
    # ----------------------------------------------------------

    report_rows = []

    for branch in branch_order:

        branch_sales = sales[
            sales["STORE"] == branch
        ].copy()

        if branch_sales.empty:
            continue

        # People who have sales come first.
        # If sales are equal, sort alphabetically.
        branch_sales = (
            branch_sales
            .assign(
                HAS_SALES=(
                    branch_sales["ACHIEVED SALES"] > 0
                )
            )
            .sort_values(
                ["HAS_SALES", "ACHIEVED SALES", "SALES MAN"],
                ascending=[False, False, True]
            )
            .drop(columns=["HAS_SALES"])
        )

        position = 1

        for _, row in branch_sales.iterrows():

            sales_person = row["SALES MAN"]
            achieved_sales = row["ACHIEVED SALES"]

            # --------------------------------------------------
            # TARGET LOOKUP
            # --------------------------------------------------

            monthly_target = sales_person_targets.get(
                (
                    branch,
                    sales_person
                ),
                None
            )

            working_days = branch_working_days[branch]

            working_days_till_now = (
                branch_working_days_till_now[branch]
            )

            remaining_days = (
                working_days
                - working_days_till_now
            )

            # --------------------------------------------------
            # CALCULATIONS
            # --------------------------------------------------

            if monthly_target is not None:

                remaining_target = (
                    monthly_target
                    - achieved_sales
                )

                if monthly_target != 0:

                    sales_percentage = (
                        achieved_sales
                        / monthly_target
                    )

                    required_sales_percentage = (
                        remaining_target
                        / monthly_target
                    )

                else:

                    sales_percentage = 0
                    required_sales_percentage = 0

                if working_days != 0:

                    daily_target = (
                        monthly_target
                        / working_days
                    )

                else:

                    daily_target = 0

                if remaining_days != 0:

                    remaining_daily_target = (
                        remaining_target
                        / remaining_days
                    )

                else:

                    remaining_daily_target = 0

                total_target_till_now = (
                    working_days_till_now
                    * daily_target
                )

            else:

                remaining_target = None
                sales_percentage = None
                required_sales_percentage = None
                daily_target = None
                remaining_daily_target = None
                total_target_till_now = None

            # --------------------------------------------------
            # ADD ROW
            # --------------------------------------------------

            report_rows.append({

                "BRANCH": branch,

                "SL.NO": position,

                "SALES MAN": sales_person,

                "MONTHLY TARGET": monthly_target,

                "ACHIEVED SALES": achieved_sales,

                "REMAINING TARGET": remaining_target,

                "REMAINING DAYS": remaining_days,

                "REMAINING DAILY TARGET":
                    remaining_daily_target,

                "SALES %": sales_percentage,

                "WORKING DAYS IN THIS MONTH":
                    working_days,

                "DAILY TARGET": daily_target,

                "WORKING DAYS TILL NOW":
                    working_days_till_now,

                "TOTAL TARGET HAVE TO HAPPEN":
                    total_target_till_now,

                "REQUIRED SALES %":
                    required_sales_percentage
            })

            position += 1

    report = pd.DataFrame(report_rows)

    return report, report_date


# ==============================================================
# 3. BORDER
# ==============================================================

thin_border = Border(
    left=Side(
        style="thin",
        color="000000"
    ),
    right=Side(
        style="thin",
        color="000000"
    ),
    top=Side(
        style="thin",
        color="000000"
    ),
    bottom=Side(
        style="thin",
        color="000000"
    )
)


# ==============================================================
# 4. COLUMN WIDTHS
# ==============================================================

def set_column_widths(ws):

    widths = {

        "A": 8,
        "B": 28,
        "C": 20,
        "D": 20,
        "E": 20,
        "F": 17,
        "G": 28,
        "H": 14,
        "I": 24,
        "J": 16,
        "K": 24,
        "L": 25,
        "M": 18

    }

    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width


# ==============================================================
# 5. HEADERS
# ==============================================================

def write_headers(ws):

    headers = [

        "SL.NO",
        "SALES MAN",
        "MONTHLY TARGET",
        "ACHIEVED SALES",
        "REMAINING TARGET",
        "REMAINING DAYS",
        "REMAINING DAILY TARGET",
        "SALES %",
        "WORKING DAYS IN THIS MONTH",
        "DAILY TARGET",
        "WORKING DAYS TILL NOW",
        "TOTAL TARGET HAVE TO HAPPEN",
        "REQUIRED SALES %"

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

        cell.font = Font(
            name="Calibri",
            size=10,
            bold=True
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="E6B8B7"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        cell.border = thin_border

    ws.row_dimensions[2].height = 35


# ==============================================================
# 6. SALESPERSON ROW
# ==============================================================

def write_sales_row(ws, row_number, row):

    values = [

        row["SL.NO"],
        row["SALES MAN"],
        row["MONTHLY TARGET"],
        row["ACHIEVED SALES"],
        row["REMAINING TARGET"],
        row["REMAINING DAYS"],
        row["REMAINING DAILY TARGET"],
        row["SALES %"],
        row["WORKING DAYS IN THIS MONTH"],
        row["DAILY TARGET"],
        row["WORKING DAYS TILL NOW"],
        row["TOTAL TARGET HAVE TO HAPPEN"],
        row["REQUIRED SALES %"]

    ]

    for col_num, value in enumerate(
        values,
        start=1
    ):

        cell = ws.cell(
            row=row_number,
            column=col_num
        )

        cell.value = value

        cell.border = thin_border

        cell.font = Font(
            name="Calibri",
            size=10,
            bold=True
        )

        if col_num == 2:

            cell.alignment = Alignment(
                horizontal="left",
                vertical="center"
            )

        else:

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        if col_num in [
            3, 4, 5, 7, 10, 12
        ]:

            cell.number_format = "#,##0.00"

        if col_num in [
            8, 13
        ]:

            cell.number_format = "0.00%"


# ==============================================================
# 7. TOTAL ROW
# ==============================================================

def write_total_row(
    ws,
    row_number,
    branch_report
):

    label_cell = ws.cell(
        row=row_number,
        column=2
    )

    label_cell.value = "TOTAL"

    label_cell.font = Font(
        name="Calibri",
        size=10,
        bold=True,
        italic=True,
        color="C65911"
    )

    label_cell.alignment = Alignment(
        horizontal="left",
        vertical="center"
    )

    # Total target
    total_target = (
        branch_report["MONTHLY TARGET"]
        .dropna()
        .sum()
    )

    target_cell = ws.cell(
        row=row_number,
        column=3
    )

    target_cell.value = total_target
    target_cell.number_format = "#,##0.00"

    target_cell.font = Font(
        name="Calibri",
        size=10,
        bold=True,
        color="C65911"
    )

    # Total achieved
    total_achieved = (
        branch_report["ACHIEVED SALES"]
        .sum()
    )

    achieved_cell = ws.cell(
        row=row_number,
        column=4
    )

    achieved_cell.value = total_achieved
    achieved_cell.number_format = "#,##0.00"

    achieved_cell.font = Font(
        name="Calibri",
        size=10,
        bold=True,
        color="4472C4"
    )

    # Format entire total row
    for col in range(1, 14):

        cell = ws.cell(
            row=row_number,
            column=col
        )

        cell.border = thin_border

        cell.fill = PatternFill(
            "solid",
            fgColor="FCE4D6"
        )

        if col not in [2, 3, 4]:

            cell.font = Font(
                name="Calibri",
                size=10,
                bold=True
            )


# ==============================================================
# 8. NEW TARGET SHEET
# ==============================================================

def create_new_target_sheet(
    ws,
    report,
    report_date
):

    # Title
    title = (
        "SALES MAN DAILY TARGET REVIEW - "
        + report_date.strftime("%Y")
        + " TILL "
        + report_date.strftime("%B %d")
    )

    ws.merge_cells("A1:M1")

    ws["A1"] = title

    ws["A1"].font = Font(
        name="Calibri",
        size=16,
        bold=True
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws["A1"].fill = PatternFill(
        "solid",
        fgColor="FFFF00"
    )

    ws.row_dimensions[1].height = 30

    # Headers
    write_headers(ws)

    current_row = 3

    # MABELLAH
    branch = "MABELLAH"

    branch_report = report[
        report["BRANCH"] == branch
    ].copy()

    if not branch_report.empty:

        # MABELLAH header
        ws.merge_cells(
            start_row=current_row,
            start_column=1,
            end_row=current_row,
            end_column=13
        )

        branch_cell = ws.cell(
            row=current_row,
            column=1
        )

        branch_cell.value = "MABELLAH"

        branch_cell.font = Font(
            name="Calibri",
            size=11,
            bold=True
        )

        branch_cell.fill = PatternFill(
            "solid",
            fgColor="B7DEE8"
        )

        branch_cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        current_row += 1

        # All salespeople are written here
        for _, row in branch_report.iterrows():

            write_sales_row(
                ws,
                current_row,
                row
            )

            current_row += 1

        # Total
        write_total_row(
            ws,
            current_row,
            branch_report
        )

    set_column_widths(ws)

    for row in range(
        3,
        ws.max_row + 1
    ):

        ws.row_dimensions[row].height = 20

    ws.freeze_panes = "A3"


# ==============================================================
# 9. MABELLAH SHEET
# ==============================================================

def create_mabellah_sheet(
    ws,
    branch_report,
    report_date
):

    title = (
        "MABELLAH  "
        + report_date.strftime("%Y %B")
    )

    ws.merge_cells("A1:M1")

    ws["A1"] = title

    ws["A1"].font = Font(
        name="Calibri",
        size=16,
        bold=True
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    ws["A1"].fill = PatternFill(
        "solid",
        fgColor="FFFF00"
    )

    ws.row_dimensions[1].height = 30

    write_headers(ws)

    current_row = 3

    for _, row in branch_report.iterrows():

        write_sales_row(
            ws,
            current_row,
            row
        )

        current_row += 1

    write_total_row(
        ws,
        current_row,
        branch_report
    )

    set_column_widths(ws)

    for row in range(
        3,
        ws.max_row + 1
    ):

        ws.row_dimensions[row].height = 20

    ws.freeze_panes = "A3"



# ==============================================================
# 10. DAILY SALES + QUOTATION REPORT
# ==============================================================

def create_daily_sales_quotation_sheet(
    ws,
    df,
    report_date
):
    """Detailed day-by-day Sales and Quotation tables.

    Days are written in DESCENDING order: max date, max-1, ... 1.
    Every calendar day is included, including Fridays.
    """

    yellow_fill = PatternFill("solid", fgColor="FFFF00")
    dark_blue_fill = PatternFill("solid", fgColor="17365D")
    light_blue_fill = PatternFill("solid", fgColor="D9EAF7")

    white_font = Font(
        name="Calibri", size=11, bold=True, color="FFFFFF"
    )
    normal_font = Font(name="Calibri", size=11)
    red_font = Font(
        name="Calibri", size=11, bold=True, color="FF0000"
    )

    # ----------------------------------------------------------
    # FIND QUOTATION AMOUNT COLUMN
    # ----------------------------------------------------------
    normalized_columns = {
        str(col).strip().upper(): col
        for col in df.columns
    }

    quotation_amount_column = None

    for candidate in [
        "ORDER LINES/TOTAL",
        "ORDER LINES / TOTAL",
        "ORDER LINES",
        "TOTAL"
    ]:
        if candidate in normalized_columns:
            quotation_amount_column = normalized_columns[candidate]
            break

    if quotation_amount_column is None:
        quotation_amount_column = "SUB TOTAL"

    working_df = df.copy()

    working_df["DATE"] = pd.to_datetime(
        working_df["DATE"], errors="coerce"
    )

    working_df["SUB TOTAL"] = pd.to_numeric(
        working_df["SUB TOTAL"], errors="coerce"
    ).fillna(0)

    working_df[quotation_amount_column] = pd.to_numeric(
        working_df[quotation_amount_column], errors="coerce"
    ).fillna(0)

    # ----------------------------------------------------------
    # MONTH DATES - DESCENDING
    # ----------------------------------------------------------
    month_start = report_date.replace(day=1)

    dates = pd.date_range(
        start=month_start,
        end=report_date,
        freq="D"
    )[::-1]

    # ----------------------------------------------------------
    # COLUMN WIDTHS
    # ----------------------------------------------------------
    # ----------------------------------------------------------
    # WIDER DAILY TABLE COLUMNS
    # ----------------------------------------------------------

    widths = {
        "A": 21,
        "B": 28,
        "C": 21,
        "D": 28,
        "E": 23,
        "F": 4,
        "G": 21,
        "H": 28,
        "I": 21,
        "J": 31,
        "K": 28
    }

    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    # ----------------------------------------------------------
    # DAY BLOCK
    # ----------------------------------------------------------
    def write_day_block(date_value, start_row):

        date_text = date_value.strftime("%B-%d")

        # SALES TITLE
        ws.merge_cells(
            start_row=start_row,
            start_column=1,
            end_row=start_row,
            end_column=5
        )

        sales_title = ws.cell(
            row=start_row, column=1
        )
        sales_title.value = "SALES " + date_text
        sales_title.fill = yellow_fill
        sales_title.font = Font(
            name="Calibri", size=13, bold=True
        )
        sales_title.alignment = Alignment(
            horizontal="center", vertical="center"
        )
        sales_title.border = thin_border

        # QUOTATION TITLE
        ws.merge_cells(
            start_row=start_row,
            start_column=7,
            end_row=start_row,
            end_column=11
        )

        quotation_title = ws.cell(
            row=start_row, column=7
        )
        quotation_title.value = "QUOTATION " + date_text
        quotation_title.fill = yellow_fill
        quotation_title.font = Font(
            name="Calibri", size=13, bold=True
        )
        quotation_title.alignment = Alignment(
            horizontal="center", vertical="center"
        )
        quotation_title.border = thin_border

        # HEADERS
        sales_headers = [
            "STORE",
            "MAKE",
            "Count of MAKE",
            "Sum of SUB TOTAL",
            "% of SUB TOTAL"
        ]

        quotation_headers = [
            "Store",
            "Make",
            "Count of Make",
            "Sum of Order Lines/Total",
            "% of Order Lines/Total"
        ]

        for col_num, header in enumerate(sales_headers, start=1):
            cell = ws.cell(
                row=start_row + 1,
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
            cell.border = thin_border

        for col_num, header in enumerate(quotation_headers, start=7):
            cell = ws.cell(
                row=start_row + 1,
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
            cell.border = thin_border

        # DAY DATA
        day_start = pd.Timestamp(date_value).normalize()
        day_end = day_start + pd.Timedelta(days=1)

        day_df = working_df[
            (working_df["DATE"] >= day_start) &
            (working_df["DATE"] < day_end)
        ].copy()

        # SALES
        sales_day = day_df[
            day_df["ENQUIRY STATUS"] == "SALES"
        ].copy()

        if not sales_day.empty:
            sales_group = (
                sales_day.groupby("MAKE", dropna=False)
                .agg(
                    COUNT=("MAKE", "size"),
                    AMOUNT=("SUB TOTAL", "sum")
                )
                .reset_index()
                .sort_values("AMOUNT", ascending=False)
            )
        else:
            sales_group = pd.DataFrame(
                columns=["MAKE", "COUNT", "AMOUNT"]
            )

        # QUOTATION
        quotation_day = day_df[
            day_df["ENQUIRY STATUS"] == "QUOTATION"
        ].copy()

        if not quotation_day.empty:
            quotation_group = (
                quotation_day.groupby("MAKE", dropna=False)
                .agg(
                    COUNT=("MAKE", "size"),
                    AMOUNT=(quotation_amount_column, "sum")
                )
                .reset_index()
                .sort_values("AMOUNT", ascending=False)
            )
        else:
            quotation_group = pd.DataFrame(
                columns=["MAKE", "COUNT", "AMOUNT"]
            )

        detail_count = max(
            len(sales_group),
            len(quotation_group),
            1
        )

        first_detail_row = start_row + 2

        sales_total = (
            float(sales_group["AMOUNT"].sum())
            if not sales_group.empty else 0
        )

        quotation_total = (
            float(quotation_group["AMOUNT"].sum())
            if not quotation_group.empty else 0
        )

        # SALES ROWS
        for index in range(detail_count):
            row_num = first_detail_row + index

            if index < len(sales_group):
                make_value = sales_group.iloc[index]["MAKE"]
                if pd.isna(make_value):
                    make_value = "(BLANK)"

                count_value = int(
                    sales_group.iloc[index]["COUNT"]
                )
                amount_value = float(
                    sales_group.iloc[index]["AMOUNT"]
                )
                percentage = (
                    amount_value / sales_total
                    if sales_total != 0 else 0
                )

                values = [
                    "Mabellah" if index == 0 else "",
                    make_value,
                    count_value,
                    amount_value,
                    percentage
                ]
            else:
                values = [None] * 5

            for col_num, value in enumerate(values, start=1):
                cell = ws.cell(
                    row=row_num,
                    column=col_num
                )
                cell.value = value
                cell.border = thin_border
                cell.font = normal_font
                cell.fill = light_blue_fill
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )
                if col_num == 1:
                    cell.alignment = Alignment(
                        horizontal="left",
                        vertical="center"
                    )
                if col_num == 4:
                    cell.number_format = "#,##0.00"
                if col_num == 5:
                    cell.number_format = "0.00%"

        # QUOTATION ROWS
        for index in range(detail_count):
            row_num = first_detail_row + index

            if index < len(quotation_group):
                make_value = quotation_group.iloc[index]["MAKE"]
                if pd.isna(make_value):
                    make_value = "(BLANK)"

                count_value = int(
                    quotation_group.iloc[index]["COUNT"]
                )
                amount_value = float(
                    quotation_group.iloc[index]["AMOUNT"]
                )
                percentage = (
                    amount_value / quotation_total
                    if quotation_total != 0 else 0
                )

                values = [
                    "Mabellah" if index == 0 else "",
                    make_value,
                    count_value,
                    amount_value,
                    percentage
                ]
            else:
                values = [None] * 5

            for col_num, value in enumerate(values, start=7):
                cell = ws.cell(
                    row=row_num,
                    column=col_num
                )
                cell.value = value
                cell.border = thin_border
                cell.font = normal_font
                cell.fill = light_blue_fill
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )
                if col_num == 7:
                    cell.alignment = Alignment(
                        horizontal="left",
                        vertical="center"
                    )
                if col_num == 10:
                    cell.number_format = "#,##0.00"
                if col_num == 11:
                    cell.number_format = "0.00%"

        # TOTAL ROW
        total_row = first_detail_row + detail_count

        sales_total_values = [
            "Mabellah Total",
            None,
            int(sales_group["COUNT"].sum()) if not sales_group.empty else 0,
            sales_total,
            1 if sales_total != 0 else 0
        ]

        for col_num, value in enumerate(sales_total_values, start=1):
            cell = ws.cell(
                row=total_row,
                column=col_num
            )
            cell.value = value
            cell.border = thin_border
            cell.fill = light_blue_fill
            cell.font = red_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )
            if col_num == 1:
                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center"
                )
            if col_num == 4:
                cell.number_format = "#,##0.00"
            if col_num == 5:
                cell.number_format = "0.00%"

        quotation_total_values = [
            "Mabellah Total",
            None,
            int(quotation_group["COUNT"].sum()) if not quotation_group.empty else 0,
            quotation_total,
            1 if quotation_total != 0 else 0
        ]

        for col_num, value in enumerate(quotation_total_values, start=7):
            cell = ws.cell(
                row=total_row,
                column=col_num
            )
            cell.value = value
            cell.border = thin_border
            cell.fill = light_blue_fill
            cell.font = red_font
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )
            if col_num == 7:
                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center"
                )
            if col_num == 10:
                cell.number_format = "#,##0.00"
            if col_num == 11:
                cell.number_format = "0.00%"

        # Shorter heading + header rows, as in the reference.
        ws.row_dimensions[start_row].height = 19
        ws.row_dimensions[start_row + 1].height = 24

        return total_row + 3

    # ----------------------------------------------------------
    # WRITE MAX DATE DOWN TO DAY 1
    # ----------------------------------------------------------
    current_row = 1

    for date_value in dates:

        # ------------------------------------------------------
        # MABELLAH HOLIDAY
        # FRIDAY = HOLIDAY
        # Do not create a Sales/Quotation table for Friday.
        # Other days are created normally.
        # ------------------------------------------------------

        if date_value.dayofweek == 4:
            continue

        current_row = write_day_block(
            date_value,
            current_row
        )

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ==============================================================
# 11. DAY-WISE SUMMARY SHEET
# ==============================================================

def create_daywise_summary_sheet(
    ws,
    df,
    report_date
):
    """Summary table: day 1 up to the maximum date in the file."""

    dark_blue_fill = PatternFill("solid", fgColor="17365D")
    header_fill = PatternFill("solid", fgColor="FCE4D6")
    alt_fill = PatternFill("solid", fgColor="C6D9F1")
    total_fill = PatternFill("solid", fgColor="17365D")
    first_row_fill = PatternFill("solid", fgColor="EAF2D9")

    title = (
        "MABELLAH REPORT "
        + report_date.strftime("%B - %Y")
        + " DAY-WISE SALES"
    )

    ws.merge_cells("A1:D1")

    ws["A1"] = title
    ws["A1"].font = Font(
        name="Calibri",
        size=16,
        bold=True,
        color="000000"
    )
    ws["A1"].fill = header_fill
    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    headers = [
        "SL/NO",
        "DATE",
        "DAY-WISE SALES",
        "DAY-WISE QUOTATION"
    ]

    for col_num, header in enumerate(headers, start=1):
        cell = ws.cell(
            row=2,
            column=col_num
        )
        cell.value = header
        cell.fill = dark_blue_fill
        cell.font = Font(
            name="Calibri",
            size=10,
            bold=True,
            color="FFFFFF"
        )
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )
        cell.border = thin_border

    # ----------------------------------------------------------
    # AMOUNT COLUMN
    # ----------------------------------------------------------
    normalized_columns = {
        str(col).strip().upper(): col
        for col in df.columns
    }

    quotation_amount_column = None

    for candidate in [
        "ORDER LINES/TOTAL",
        "ORDER LINES / TOTAL",
        "ORDER LINES",
        "TOTAL"
    ]:
        if candidate in normalized_columns:
            quotation_amount_column = normalized_columns[candidate]
            break

    if quotation_amount_column is None:
        quotation_amount_column = "SUB TOTAL"

    work = df.copy()
    work["DATE"] = pd.to_datetime(
        work["DATE"], errors="coerce"
    )
    work["SUB TOTAL"] = pd.to_numeric(
        work["SUB TOTAL"], errors="coerce"
    ).fillna(0)
    work[quotation_amount_column] = pd.to_numeric(
        work[quotation_amount_column], errors="coerce"
    ).fillna(0)

    # ----------------------------------------------------------
    # ASCENDING DATES: 1ST DATE FIRST
    # ----------------------------------------------------------
    month_start = report_date.replace(day=1)
    dates = pd.date_range(
        start=month_start,
        end=report_date,
        freq="D"
    )

    row = 3

    for sl_no, date_value in enumerate(dates, start=1):

        day_start = pd.Timestamp(date_value).normalize()
        day_end = day_start + pd.Timedelta(days=1)

        day_df = work[
            (work["DATE"] >= day_start) &
            (work["DATE"] < day_end)
        ].copy()

        sales_total = day_df.loc[
            day_df["ENQUIRY STATUS"] == "SALES",
            "SUB TOTAL"
        ].sum()

        quotation_total = day_df.loc[
            day_df["ENQUIRY STATUS"] == "QUOTATION",
            quotation_amount_column
        ].sum()

        values = [
            sl_no,
            date_value.strftime("%b-%d"),
            float(sales_total),
            float(quotation_total)
        ]

        for col_num, value in enumerate(values, start=1):
            cell = ws.cell(
                row=row,
                column=col_num
            )
            cell.value = value
            cell.border = thin_border
            cell.font = Font(
                name="Calibri",
                size=10,
                bold=True
            )
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

            if sl_no % 2 == 1:
                cell.fill = first_row_fill
            else:
                cell.fill = alt_fill

            if col_num in [3, 4]:
                cell.number_format = "#,##0.000"

            ws.row_dimensions[row].height = 21

        row += 1

    # ----------------------------------------------------------
    # GRAND TOTAL
    # ----------------------------------------------------------
    sales_grand_total = work.loc[
        work["ENQUIRY STATUS"] == "SALES",
        "SUB TOTAL"
    ].sum()

    quotation_grand_total = work.loc[
        work["ENQUIRY STATUS"] == "QUOTATION",
        quotation_amount_column
    ].sum()

    total_values = [
        "",
        "TOTAL",
        float(sales_grand_total),
        float(quotation_grand_total)
    ]

    for col_num, value in enumerate(total_values, start=1):
        cell = ws.cell(
            row=row,
            column=col_num
        )
        cell.value = value
        cell.fill = total_fill
        cell.font = Font(
            name="Calibri",
            size=10,
            bold=True,
            color="FFFFFF"
        )
        cell.border = thin_border
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

        if col_num in [3, 4]:
            cell.number_format = "#,##0.00"

    
    ws.row_dimensions[row].height = 22
# ----------------------------------------------------------
    # EXACT TABLE SIZE FROM REFERENCE
    # ----------------------------------------------------------

    ws.column_dimensions["A"].width = 11
    ws.column_dimensions["B"].width = 29
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 29

    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 25
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:D{row - 1}"
    ws.sheet_view.showGridLines = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# ==============================================================
# 10. CREATE EXCEL FILE
# ==============================================================

def create_excel_report(
    report,
    report_date,
    df,
    output_file="Mabellah_Sales_Report.xlsx"
):

    wb = Workbook()

    # ----------------------------------------------------------
    # SHEET 1 - SALESMAN TARGET
    # ----------------------------------------------------------
    ws1 = wb.active
    ws1.title = "DAILY SALES MAN TARGET"

    create_new_target_sheet(
        ws1,
        report,
        report_date
    )

    # ----------------------------------------------------------
    # SHEET 2 - DETAILED DAILY SALES + QUOTATION
    # ----------------------------------------------------------
    ws2 = wb.create_sheet(
        title="DAILY SALES & QUOTATION"
    )

    create_daily_sales_quotation_sheet(
        ws2,
        df,
        report_date
    )

    # ----------------------------------------------------------
    # SHEET 3 - DAY-WISE SUMMARY
    # ----------------------------------------------------------
    ws3 = wb.create_sheet(
        title="MABELLAH DAY-WISE"
    )

    create_daywise_summary_sheet(
        ws3,
        df,
        report_date
    )

    # ----------------------------------------------------------
    # SHEET ORDER
    # ----------------------------------------------------------
    wb._sheets = [
        wb["DAILY SALES MAN TARGET"],
        wb["DAILY SALES & QUOTATION"],
        wb["MABELLAH DAY-WISE"]
    ]

    wb.save(output_file)

    print()
    print("==============================================")
    print("MABELLAH REPORT CREATED SUCCESSFULLY")
    print("==============================================")
    print("File:", output_file)
    print()
    print("Sheets:")
    print("  1. DAILY SALES MAN TARGET")
    print("  2. DAILY SALES & QUOTATION")
    print("  3. MABELLAH DAY-WISE")


# ==============================================================
# STREAMLIT WRAPPER
# ==============================================================

def generate_mabellah_sales_report(
    input_file,
    output_file
):
    """
    Generate the MABELLAH Sales Report workbook.

    The workbook contains:
    1. DAILY SALES MAN TARGET
    2. DAILY SALES & QUOTATION
    """

    df = salesman_newfile(input_file)

    report, report_date = sales_person_report(df)

    create_excel_report(
        report,
        report_date,
        df,
        output_file
    )

    return output_file
