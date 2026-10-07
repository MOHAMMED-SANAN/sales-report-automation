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

    df = pd.read_excel(
        input_file
    )


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
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:

        raise ValueError(
            "Missing columns: "
            + ", ".join(missing_columns)
        )


    # ==========================================================
    # CLEAN DATA
    # ==========================================================

    df["Order Updated At"] = pd.to_datetime(
        df["Order Updated At"],
        errors="coerce",
        dayfirst=True
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

    start_date = pd.Timestamp(
        start_date
    )


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
        &
        (df["Order Updated At"] <= end_date)
        &
        (df["Status"] == "delivered")
    ].copy()


    # ==========================================================
    # DRIVER-WISE DELIVERY COUNT
    # ==========================================================

    report_df = (
        df
        .groupby(
            [
                "Showroom",
                "Driver"
            ],
            dropna=False
        )
        .size()
        .reset_index(
            name="Total"
        )
    )


    # ==========================================================
    # SORT
    # ==========================================================

    report_df = (
        report_df
        .sort_values(
            by=[
                "Showroom",
                "Total"
            ],
            ascending=[
                True,
                False
            ]
        )
        .reset_index(
            drop=True
        )
    )


    grand_total = 0


    # ==========================================================
    # CREATE WORKBOOK
    # ==========================================================

    wb = Workbook()

    # Remove the default sheet.
    default_ws = wb.active
    wb.remove(default_ws)


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


    # ==========================================================
    # CREATE SEPARATE SHEET FOR EACH SHOWROOM
    # ==========================================================

    if report_df.empty:

        ws = wb.create_sheet(
            title="Delivery Report"
        )

        headers = [
            "SL NO",
            "DRIVER NAME",
            "VEHICLE",
            "COUNT"
        ]

        for column_number, header in enumerate(
            headers,
            start=1
        ):

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

        ws.cell(
            row=2,
            column=1,
            value="TOTAL"
        )

        ws.cell(
            row=2,
            column=4,
            value=0
        )

        for column_number in range(1, 5):

            cell = ws.cell(
                row=2,
                column=column_number
            )

            cell.border = border
            cell.font = Font(
                bold=True
            )

    else:

        for showroom, group in report_df.groupby(
            "Showroom",
            sort=False
        ):

            # Excel sheet names cannot contain these characters.
            sheet_name = str(showroom)

            for character in [
                "\\",
                "/",
                "*",
                "?",
                ":",
                "[",
                "]"
            ]:

                sheet_name = sheet_name.replace(
                    character,
                    "-"
                )

            sheet_name = sheet_name[:31]

            if not sheet_name:
                sheet_name = "Delivery Report"


            ws = wb.create_sheet(
                title=sheet_name
            )


            # ==================================================
            # HEADERS
            # ==================================================

            headers = [
                "SL NO",
                "DRIVER NAME",
                "VEHICLE",
                "COUNT"
            ]


            for column_number, header in enumerate(
                headers,
                start=1
            ):

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


            # ==================================================
            # DRIVER DATA
            # ==================================================

            excel_row = 2

            sl_no = 1

            showroom_total = 0


            for _, row in group.iterrows():

                driver_name = row["Driver"]

                driver_count = int(
                    row["Total"]
                )

                showroom_total += driver_count


                ws.cell(
                    row=excel_row,
                    column=1,
                    value=sl_no
                )

                ws.cell(
                    row=excel_row,
                    column=2,
                    value=driver_name
                )

                # Vehicle intentionally blank.
                ws.cell(
                    row=excel_row,
                    column=3,
                    value=""
                )

                ws.cell(
                    row=excel_row,
                    column=4,
                    value=driver_count
                )


                for column_number in range(1, 5):

                    cell = ws.cell(
                        row=excel_row,
                        column=column_number
                    )

                    cell.border = border

                    cell.alignment = Alignment(
                        vertical="center"
                    )


                ws.cell(
                    row=excel_row,
                    column=1
                ).alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )


                ws.cell(
                    row=excel_row,
                    column=4
                ).alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )


                excel_row += 1

                sl_no += 1


            # ==================================================
            # TOTAL
            # ==================================================

            ws.cell(
                row=excel_row,
                column=1,
                value="TOTAL"
            )

            ws.cell(
                row=excel_row,
                column=4,
                value=showroom_total
            )


            for column_number in range(1, 5):

                cell = ws.cell(
                    row=excel_row,
                    column=column_number
                )

                cell.border = border

                cell.font = Font(
                    bold=True
                )

                cell.alignment = Alignment(
                    vertical="center"
                )


            ws.cell(
                row=excel_row,
                column=1
            ).alignment = Alignment(
                horizontal="center",
                vertical="center"
            )


            ws.cell(
                row=excel_row,
                column=4
            ).alignment = Alignment(
                horizontal="center",
                vertical="center"
            )


            # ==================================================
            # COLUMN WIDTHS
            # ==================================================

            ws.column_dimensions["A"].width = 12
            ws.column_dimensions["B"].width = 28
            ws.column_dimensions["C"].width = 22
            ws.column_dimensions["D"].width = 12


            # ==================================================
            # ROW HEIGHTS
            # ==================================================

            ws.row_dimensions[1].height = 22

            for row_number in range(
                2,
                excel_row + 1
            ):

                ws.row_dimensions[
                    row_number
                ].height = 20


            # ==================================================
            # FREEZE HEADER
            # ==================================================

            ws.freeze_panes = "A2"


            # ==================================================
            # TOTAL / GRAND TOTAL
            # ==================================================

            grand_total += showroom_total


    # ==========================================================
    # SAVE
    # ==========================================================

    wb.save(
        output_file
    )
