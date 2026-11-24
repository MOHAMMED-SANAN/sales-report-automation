import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


def _norm(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def _find_col(df, candidates, required=True):
    lookup = {_norm(c): c for c in df.columns}

    for candidate in candidates:
        key = _norm(candidate)
        if key in lookup:
            return lookup[key]

    for col in df.columns:
        key = _norm(col)
        if any(_norm(x) in key for x in candidates):
            return col

    if required:
        raise ValueError(
            f"Could not find column {candidates}. "
            f"Available columns: {list(df.columns)}"
        )
    return None


def _clean_key(series):
    return (
        series.fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .str.replace(r"\.0$", "", regex=True)
    )


def _is_warehouse(value):
    return bool(
        re.match(
            r"^WH(?:\b|[_\s-]|\d)",
            str(value).strip(),
            flags=re.IGNORECASE
        )
    )


def read_inventory(inventory_file):
    inv = pd.read_excel(inventory_file, dtype=object)

    part_col = _find_col(
        inv, ["PART NUMBER", "PART NO", "ITEM CODE", "ITEM", "SKU"]
    )
    brand_col = _find_col(inv, ["BRAND"])
    make_col = _find_col(inv, ["MAKE", "MANUFACTURER"])
    store_col = _find_col(
        inv, ["STORE", "LOCATION", "BRANCH", "WAREHOUSE"]
    )

    # Available Qty is Excel column M (13th column).
    if len(inv.columns) < 13:
        raise ValueError(
            "Inventory report has fewer than 13 columns. "
            "Available Qty is expected in column M."
        )

    qty_col = inv.columns[12]

    inv = inv.copy()

    inv["_part_key"] = _clean_key(inv[part_col])
    inv["_brand_key"] = _clean_key(inv[brand_col])
    inv["_make_key"] = _clean_key(inv[make_col])

    inv["_store_name"] = (
        inv[store_col].fillna("").astype(str).str.strip()
    )

    # Handle Excel errors, blanks and non-numeric quantities as zero.
    inv["_qty"] = pd.to_numeric(
        inv[qty_col], errors="coerce"
    ).fillna(0)

    inv = inv[
        (inv["_part_key"] != "")
        & (inv["_brand_key"] != "")
        & (inv["_make_key"] != "")
    ].copy()

    stores = sorted(
        {
            s for s in inv["_store_name"]
            if s and not _is_warehouse(s)
        },
        key=str.casefold
    )

    if not stores:
        raise ValueError("No store names found in Inventory Report.")

    return (
        inv,
        stores,
        part_col,
        brand_col,
        make_col,
        store_col,
        qty_col
    )


def create_fast_moving_report(
    sales_file,
    inventory_file,
    selected_store,
    output_file
):
    # 1. READ SALES REPORT
    sales = pd.read_excel(sales_file, dtype=object)

    part_col = _find_col(
        sales, ["PART NUMBER", "PART NO", "ITEM CODE", "ITEM", "SKU"]
    )
    brand_col = _find_col(sales, ["BRAND"])
    make_col = _find_col(sales, ["MAKE", "MANUFACTURER"])

    sales_qty_col = _find_col(
        sales,
        [
            "LAST 3 MONTHS SALES HAPPEN QUANTITY",
            "LAST 3 MONTH SALES HAPPEN QUANTITY",
            "LAST 3 MONTH SALES HAPPEN",
            "LAST 3 MONTHS SALES",
            "LAST 3 MONTH SALES",
            "3 MONTH SALES",
            "SALES HAPPEN QUANTITY",
            "SALES QTY",
            "QUANTITY"
        ],
        required=False
    )

    if sales_qty_col is None:
        if len(sales.columns) >= 4:
            sales_qty_col = sales.columns[3]
        else:
            raise ValueError(
                "Could not identify the last 3 months sales quantity column."
            )

    # 2. READ INVENTORY REPORT
    (
        inv,
        stores,
        inv_part_col,
        inv_brand_col,
        inv_make_col,
        inv_store_col,
        qty_col
    ) = read_inventory(inventory_file)

    if selected_store not in stores:
        raise ValueError(
            f"Selected store '{selected_store}' was not found."
        )

    # 3. FILTER SELECTED STORE
    selected_rows = inv[
        inv["_store_name"].str.casefold()
        == selected_store.casefold()
    ].copy()

    # 4. GET ALL WH WAREHOUSE ROWS
    warehouse_rows = inv[
        inv["_store_name"].map(_is_warehouse)
    ].copy()

    # 5. MATCH USING THREE COLUMNS TOGETHER
    matching_keys = [
        "_part_key",
        "_brand_key",
        "_make_key"
    ]

    # Sum Available Qty for selected store.
    store_totals = (
        selected_rows
        .groupby(matching_keys, dropna=False)["_qty"]
        .sum()
        .to_dict()
    )

    # Sum Available Qty from ALL WH warehouses.
    wh_totals = (
        warehouse_rows
        .groupby(matching_keys, dropna=False)["_qty"]
        .sum()
        .to_dict()
    )

    # 6. CREATE SALES MATCHING KEYS
    sales = sales.copy()

    sales["_part_key"] = _clean_key(sales[part_col])
    sales["_brand_key"] = _clean_key(sales[brand_col])
    sales["_make_key"] = _clean_key(sales[make_col])

    sales["_match_key"] = list(zip(
        sales["_part_key"],
        sales["_brand_key"],
        sales["_make_key"]
    ))

    # 7. GET STORE AND WH TOTAL QUANTITIES
    sales["_store_qty"] = [
        store_totals.get(key, 0)
        for key in sales["_match_key"]
    ]

    sales["_wh_qty"] = [
        wh_totals.get(key, 0)
        for key in sales["_match_key"]
    ]

    # 8. KEEP ONLY ZERO SELECTED-STORE STOCK
    sales = sales[
        (sales["_part_key"] != "")
        & (sales["_brand_key"] != "")
        & (sales["_make_key"] != "")
        & (sales["_store_qty"] == 0)
    ].copy()

    # 9. GET MAKE VALUES
    makes = sorted(
        sales[make_col]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda x: x != ""]
        .unique(),
        key=str.casefold
    )

    if not makes:
        raise ValueError(
            "No items found with zero stock in the selected store."
        )

    # 10. CREATE EXCEL WORKBOOK
    wb = Workbook()
    wb.remove(wb.active)

    yellow = PatternFill("solid", fgColor="FFFF00")
    blue = PatternFill("solid", fgColor="305496")

    white_bold = Font(
        name="Arial", size=10, bold=True, color="FFFFFF"
    )
    title_font = Font(
        name="Arial", size=12, bold=True, color="000000"
    )
    body_font = Font(
        name="Arial", size=10, bold=True, color="000000"
    )

    thin_black = Side(style="thin", color="000000")
    cell_border = Border(
        left=thin_black,
        right=thin_black,
        top=thin_black,
        bottom=thin_black
    )

    headers = [
        "PART NUMBER",
        "BRAND",
        "MAKE",
        "LAST 3 MONTH SALES HAPPEN",
        "STORE",
        "WH"
    ]

    # 11. CREATE ONE SHEET FOR EACH MAKE
    for make in makes:

        safe_name = re.sub(
            r'[\\/*?:\[\]]', "_", make
        )[:31] or "MAKE"

        base = safe_name
        sheet_name = base
        n = 2

        while sheet_name in wb.sheetnames:
            suffix = f"_{n}"
            sheet_name = base[:31 - len(suffix)] + suffix
            n += 1

        ws = wb.create_sheet(sheet_name)

        make_rows = sales[
            sales[make_col].astype(str).str.strip().str.casefold()
            == make.casefold()
        ]

        # TITLE
        ws.merge_cells("A1:F1")
        title_cell = ws["A1"]
        title_cell.value = f"{make.upper()} FAST MOVING REPORT"
        title_cell.fill = yellow
        title_cell.font = title_font
        title_cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )
        ws.row_dimensions[1].height = 24

        # HEADINGS
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_idx, value=header)
            cell.fill = blue
            cell.font = white_bold
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )
            cell.border = cell_border

        ws.row_dimensions[2].height = 35

        # DATA
        for row_idx, (_, row) in enumerate(
            make_rows.iterrows(), 3
        ):
            values = [
                row[part_col],
                row[brand_col],
                row[make_col],
                row[sales_qty_col],
                row["_store_qty"],
                row["_wh_qty"]
            ]

            for col_idx, value in enumerate(values, 1):
                if pd.isna(value):
                    value = ""

                cell = ws.cell(
                    row=row_idx,
                    column=col_idx,
                    value=value
                )
                cell.font = body_font
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )
                cell.border = cell_border

        # COLUMN WIDTHS
        widths = [22, 18, 18, 35, 17, 17]

        for i, width in enumerate(widths, 1):
            ws.column_dimensions[
                get_column_letter(i)
            ].width = width

        # EXCEL SETTINGS
        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A2:F{max(2, ws.max_row)}"
        ws.sheet_view.showGridLines = True

    # 12. SAVE REPORT
    wb.save(output_file)

    return len(makes)
