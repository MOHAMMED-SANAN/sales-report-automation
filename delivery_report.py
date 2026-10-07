import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment


def generate_delivery_report(
    input_file,
    output_file,
    start_date,
    end_date
):
    """
    Generate Store Wise Delivery Report.

    Filters:
        Order Updated At = selected date range
        Status = delivered

    Output columns:
        Showroom | Driver | Status | Total
    """

    # ==========================================================
    # READ EXCEL
    # ==========================================================

    df = pd.read_excel(input_file)

    # ==========================================================
    # REQUIRED COLUMNS
    # ==========================================================

    required_columns = [
        "Showroom",
        "Driver",
        "Status",
        "Order Updated At"
    ]

    missing_columns = [
        col for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing columns: " + ", ".join(missing_columns)
        )

    # ==========================================================
    # CLEAN DATA
    # ==========================================================

    df["Order Updated At"] = pd.to_datetime(
        df["Order Updated At"],
        errors="coerce"
    )

    df["Status"] = (
        df["Status"]
        .astype(str)
        .str.strip()
        .str.lower()
    )

    df["Showroom"] = (
        df["Showroom"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    df["Driver"] = (
        df["Driver"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    # ==========================================================
    # DATE RANGE
    # ==========================================================

    start_date = pd.Timestamp(start_date)

    end_date = (
        pd.Timestamp(end_date)
        + pd.Timedelta(days=1)
        - pd.Timedelta(seconds=1)
    )

    # ==========================================================
    # FILTER
    # ONLY DELIVERED ORDERS
    # ==========================================================

    df = df[
        (df["Order Updated At"] >= start_date)
        & (df["Order Updated At"] <= end_date)
        & (df["Status"] == "delivered")
    ].copy()

    # ==========================================================
    # DRIVER-WISE DELIVERY COUNT
    # ==========================================================

    report_df = (
        df.groupby(
            ["Showroom", "Driver"],
            dropna=False
        )
        .size()
        .reset_index(name="Total")
    )

    # ==========================================================
    # SORT
    # ==========================================================

    report_df = (
        report_df
        .sort_values(
            by=["Showroom", "Total"],
            ascending=[True, False]
        )
        .reset_index(drop=True)
    )

    # ==========================================================
    # CREATE WORKBOOK
    # ==========================================================

    wb = Workbook()
    ws = wb.active
    ws.title = "Delivery Report"

    # ==========================================================
    # TITLE
    # ==========================================================

    if start_date.month == end_date.month:
        title = (
            "STORE WISE DELIVERIES - "
            f"{start_date.strftime('%B').upper()}-"
            f"{end_date.day}"
        )
    else:
        title = (
            "STORE WISE DELIVERIES - "
            f"{start_date.strftime('%d-%m-%Y')} "
            f"TO {end_date.strftime('%d-%m-%Y')}"
        )

    ws.merge_cells("A1:D1")
    ws["A1"] = title

    # ==========================================================
    # TITLE STYLE
    # ==========================================================

    ws["A1"].fill = PatternFill(
        fill_type="solid",
        fgColor="FFFF00"
    )

    ws["A1"].font = Font(
        bold=True,
        color="000000",
        size=12
    )

    ws["A1"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

    # ==========================================================
    # HEADERS
    # ==========================================================

    headers = [
        "Showroom",
        "Driver",
        "Status",
        "Total"
    ]

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="5B7DBB"
    )

    header_font = Font(
        color="FFFFFF"
    )

    black_side = Side(
        style="thin",
        color="000000"
    )

    border = Border(
        left=black_side,
        right=black_side,
        top=black_side,
        bottom=black_side
    )

    for col_num, header in enumerate(headers, start=1):
        cell = ws.cell(
            row=2,
            column=col_num,
            value=header
        )

        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

    # ==========================================================
    # WRITE REPORT
    # ==========================================================

    excel_row = 3
    grand_total = 0

    for showroom, group in report_df.groupby(
        "Showroom",
        sort=False
    ):
        showroom_total = int(group["Total"].sum())
        grand_total += showroom_total

        first_driver = True

        # ------------------------------------------------------
        # DRIVER ROWS
        # ------------------------------------------------------

        for _, row in group.iterrows():

            showroom_value = showroom if first_driver else ""
            first_driver = False

            ws.cell(
                row=excel_row,
                column=1,
                value=showroom_value
            )

            ws.cell(
                row=excel_row,
                column=2,
                value=row["Driver"]
            )

            ws.cell(
                row=excel_row,
                column=3,
                value="delivered"
            )

            ws.cell(
                row=excel_row,
                column=4,
                value=int(row["Total"])
            )

            for col_num in range(1, 5):
                cell = ws.cell(
                    row=excel_row,
                    column=col_num
                )

                cell.border = border
                cell.alignment = Alignment(
                    vertical="center"
                )

            ws.cell(
                row=excel_row,
                column=4
            ).alignment = Alignment(
                horizontal="right",
                vertical="center"
            )

            excel_row += 1

        # ------------------------------------------------------
        # SHOWROOM TOTAL
        # ------------------------------------------------------

        ws.cell(
            row=excel_row,
            column=1,
            value=f"{showroom} Total"
        )

        ws.cell(
            row=excel_row,
            column=4,
            value=showroom_total
        )

        for col_num in range(1, 5):
            cell = ws.cell(
                row=excel_row,
                column=col_num
            )

            cell.border = border
            cell.font = Font(bold=True)
            cell.alignment = Alignment(
                vertical="center"
            )

        ws.cell(
            row=excel_row,
            column=4
        ).alignment = Alignment(
            horizontal="right",
            vertical="center"
        )

        excel_row += 1

    # ==========================================================
    # GRAND TOTAL
    # ==========================================================

    ws.cell(
        row=excel_row,
        column=1,
        value="GRAND TOTAL"
    )

    ws.cell(
        row=excel_row,
        column=4,
        value=grand_total
    )

    for col_num in range(1, 5):
        cell = ws.cell(
            row=excel_row,
            column=col_num
        )

        cell.border = border
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            vertical="center"
        )

    ws.cell(
        row=excel_row,
        column=4
    ).alignment = Alignment(
        horizontal="right",
        vertical="center"
    )

    # ==========================================================
    # COLUMN WIDTHS
    # ==========================================================

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 12

    # ==========================================================
    # ROW HEIGHTS
    # ==========================================================

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[2].height = 22

    for row_num in range(3, excel_row + 1):
        ws.row_dimensions[row_num].height = 20

    # ==========================================================
    # FREEZE HEADER
    # ==========================================================

    ws.freeze_panes = "A3"

    # ==========================================================
    # SAVE FILE
    # ==========================================================

    wb.save(output_file)
