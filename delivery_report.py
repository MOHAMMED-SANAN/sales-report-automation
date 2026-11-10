
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Border,
    Side,
    Alignment
)


def generate_delivery_report(
    input_file,
    output_file,
    start_date,
    end_date
):

    # ==========================================================
    # READ EXCEL
    # ==========================================================

    df = pd.read_excel(input_file)

    # ==========================================================
    # COLUMN SETTINGS
    # ==========================================================

    ORDER_TYPE_COLUMN = "Order Type"

    required_columns = [
        "Showroom",
        "Driver",
        "Status",
        "Order Updated At",
        ORDER_TYPE_COLUMN
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
        errors="coerce",
        dayfirst=True
    )

    for col in ["Status", "Showroom", "Driver", ORDER_TYPE_COLUMN]:
        df[col] = (
            df[col]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    df["Status"] = df["Status"].str.lower()

    # ==========================================================
    # NORMALIZE ORDER TYPE
    # ==========================================================

    order_type_mapping = {
        "purchase order": "PURCHASE ORDER",
        "purchase": "PURCHASE ORDER",
        "po": "PURCHASE ORDER",

        "sales order": "SALES ORDER",
        "sale order": "SALES ORDER",
        "sales": "SALES ORDER",
        "sale": "SALES ORDER",
        "so": "SALES ORDER"
    }

    df["REPORT ORDER TYPE"] = (
        df[ORDER_TYPE_COLUMN]
        .str.lower()
        .map(order_type_mapping)
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
    # FILTER DELIVERED ORDERS
    # ==========================================================

    df = df[
        (df["Order Updated At"] >= start_date)
        &
        (df["Order Updated At"] <= end_date)
        &
        (df["Status"] == "delivered")
        &
        (df["REPORT ORDER TYPE"].notna())
    ].copy()

    # ==========================================================
    # DRIVER-WISE ORDER TYPE COUNT
    # ==========================================================

    report_df = (
        df.groupby(
            [
                "Showroom",
                "Driver",
                "REPORT ORDER TYPE"
            ]
        )
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )

    # Ensure both columns always exist
    for col in ["PURCHASE ORDER", "SALES ORDER"]:
        if col not in report_df.columns:
            report_df[col] = 0

    report_df["COUNT"] = (
        report_df["PURCHASE ORDER"]
        + report_df["SALES ORDER"]
    )

    # ==========================================================
    # SORT
    # ==========================================================

    report_df = (
        report_df
        .sort_values(
            by=["Showroom", "COUNT"],
            ascending=[True, False]
        )
        .reset_index(drop=True)
    )

    # ==========================================================
    # CREATE WORKBOOK
    # ==========================================================

    wb = Workbook()
    wb.remove(wb.active)

    # ==========================================================
    # COMMON STYLES
    # ==========================================================

    blue_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    white_font = Font(
        color="FFFFFF",
        bold=True
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

    headers = [
        "SL NO",
        "DRIVER NAME",
        "VEHICLE",
        "PURCHASE ORDER",
        "SALES ORDER",
        "COUNT"
    ]

    # ==========================================================
    # SHOWROOM SHEETS
    # ==========================================================

    if report_df.empty:

        showroom_groups = [( "Delivery Report", report_df )]

    else:

        showroom_groups = list(
            report_df.groupby("Showroom", sort=False)
        )

    for showroom, group in showroom_groups:

        # Clean Excel sheet name
        sheet_name = str(showroom)

        for character in ["\\", "/", "*", "?", ":", "[", "]"]:
            sheet_name = sheet_name.replace(character, "-")

        sheet_name = sheet_name[:31] or "Delivery Report"

        # Avoid duplicate sheet names
        original_name = sheet_name
        counter = 1

        while sheet_name in wb.sheetnames:
            suffix = f"_{counter}"
            sheet_name = original_name[:31-len(suffix)] + suffix
            counter += 1

        ws = wb.create_sheet(title=sheet_name)

        # ======================================================
        # HEADERS
        # ======================================================

        for column_number, header in enumerate(headers, start=1):

            cell = ws.cell(
                row=1,
                column=column_number,
                value=header
            )

            cell.fill = blue_fill
            cell.font = white_font
            cell.border = border

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # ======================================================
        # DRIVER DATA
        # ======================================================

        excel_row = 2
        sl_no = 1

        purchase_total = 0
        sales_total = 0

        for _, row in group.iterrows():

            driver_name = row["Driver"]

            purchase_count = int(row["PURCHASE ORDER"])
            sales_count = int(row["SALES ORDER"])

            total_count = purchase_count + sales_count

            purchase_total += purchase_count
            sales_total += sales_count

            values = [
                sl_no,
                driver_name,
                "",
                purchase_count,
                sales_count,
                total_count
            ]

            for column_number, value in enumerate(values, start=1):

                cell = ws.cell(
                    row=excel_row,
                    column=column_number,
                    value=value
                )

                cell.border = border

                cell.font = Font(
                    size=12,
                    bold=True
                )

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            excel_row += 1
            sl_no += 1

        # ======================================================
        # TOTAL ROW
        # ======================================================

        total_count = purchase_total + sales_total

        total_values = [
            "TOTAL",
            "",
            "",
            purchase_total,
            sales_total,
            total_count
        ]

        for column_number, value in enumerate(total_values, start=1):

            cell = ws.cell(
                row=excel_row,
                column=column_number,
                value=value
            )

            cell.border = border

            cell.font = Font(
                bold=True,
                size=12
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # ======================================================
        # COLUMN WIDTHS
        # ======================================================

        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 40
        ws.column_dimensions["C"].width = 25
        ws.column_dimensions["D"].width = 22
        ws.column_dimensions["E"].width = 22
        ws.column_dimensions["F"].width = 14

        # ======================================================
        # ROW HEIGHTS
        # ======================================================

        ws.row_dimensions[1].height = 28

        for row_number in range(2, excel_row + 1):
            ws.row_dimensions[row_number].height = 26

        # ======================================================
        # FREEZE HEADER
        # ======================================================

        ws.freeze_panes = "A2"

    # ==========================================================
    # SAVE WORKBOOK
    # ==========================================================

    wb.save(output_file)

    print("Delivery report generated successfully!")
    print("Output:", output_file)
