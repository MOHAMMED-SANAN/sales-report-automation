import re
from pathlib import Path

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Border,
    Side,
    Alignment
)
from openpyxl.utils import get_column_letter


def _norm(value):
    return re.sub(
        r'[^a-z0-9]+',
        '',
        str(value).strip().lower()
    )


def _find_col(df, candidates, required=True):
    lookup = {_norm(c): c for c in df.columns}

    for candidate in candidates:
        key = _norm(candidate)

        if key in lookup:
            return lookup[key]

    # Allow partial header matching
    for c in df.columns:
        key = _norm(c)

        if any(_norm(x) in key for x in candidates):
            return c

    if required:
        raise ValueError(
            f"Could not find required column: {candidates}. "
            f"Available columns: {list(df.columns)}"
        )

    return None


def _clean_key(series):
    return (
        series.fillna('')
        .astype(str)
        .str.strip()
        .str.upper()
    )


def _is_warehouse(value):
    return bool(
        re.match(
            r'^WH(?:_|\b|\s)',
            str(value).strip(),
            flags=re.IGNORECASE
        )
    )


# --------------------------------------------------
# READ INVENTORY REPORT
# --------------------------------------------------

def read_inventory(inventory_file):

    inv = pd.read_excel(
        inventory_file,
        dtype=object
    )

    part_col = _find_col(
        inv,
        ['PART NUMBER', 'PART NO', 'ITEM CODE', 'ITEM', 'SKU']
    )

    brand_col = _find_col(
        inv,
        ['BRAND']
    )

    make_col = _find_col(
        inv,
        ['MAKE', 'MANUFACTURER']
    )

    store_col = _find_col(
        inv,
        ['STORE', 'LOCATION', 'BRANCH', 'WAREHOUSE']
    )

    # Available Qty is column M (13th Excel column)
    if len(inv.columns) < 13:
        raise ValueError(
            'Inventory report has fewer than 13 columns. '
            'Available Qty is expected in column M.'
        )

    qty_col = inv.columns[12]

    inv = inv.copy()

    # Create matching keys
    inv['_part_key'] = _clean_key(inv[part_col])
    inv['_brand_key'] = _clean_key(inv[brand_col])
    inv['_make_key'] = _clean_key(inv[make_col])

    inv['_store_name'] = (
        inv[store_col]
        .fillna('')
        .astype(str)
        .str.strip()
    )

    inv['_qty'] = pd.to_numeric(
        inv[qty_col],
        errors='coerce'
    ).fillna(0)

    # Remove rows without matching keys
    inv = inv[
        (inv['_part_key'] != '') &
        (inv['_brand_key'] != '') &
        (inv['_make_key'] != '')
    ]

    # Get available stores, excluding WH warehouses
    stores = sorted(
        {
            s for s in inv['_store_name']
            if s and not _is_warehouse(s)
        },
        key=str.casefold
    )

    if not stores:
        raise ValueError(
            'No store names found in inventory report.'
        )

    return (
        inv,
        stores,
        part_col,
        brand_col,
        make_col,
        store_col,
        qty_col
    )


# --------------------------------------------------
# CREATE FAST MOVING REPORT
# --------------------------------------------------

def create_fast_moving_report(
    sales_file,
    inventory_file,
    selected_store,
    output_file
):

    # READ SALES REPORT
    sales = pd.read_excel(
        sales_file,
        dtype=object
    )

    part_col = _find_col(
        sales,
        ['PART NUMBER', 'PART NO', 'ITEM CODE', 'ITEM', 'SKU']
    )

    brand_col = _find_col(
        sales,
        ['BRAND']
    )

    make_col = _find_col(
        sales,
        ['MAKE', 'MANUFACTURER']
    )

    sales_qty_col = _find_col(
        sales,
        [
            'LAST 3 MONTH SALES HAPPEN QUANTITY',
            'LAST 3 MONTHS SALES HAPPEN QUANTITY',
            'LAST 3 MONTH SALES',
            'LAST 3 MONTHS SALES',
            '3 MONTH SALES',
            'SALES HAPPEN QUANTITY',
            'SALES QTY',
            'QUANTITY'
        ],
        required=False
    )

    if sales_qty_col is None:

        if len(sales.columns) >= 4:
            sales_qty_col = sales.columns[3]

        else:
            raise ValueError(
                'Could not identify sales quantity column.'
            )

    # READ INVENTORY
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
            f"Selected store '{selected_store}' "
            "was not found in inventory report."
        )

    # --------------------------------------------------
    # FILTER SELECTED STORE AND WH
    # --------------------------------------------------

    selected_rows = inv[
        inv['_store_name'].str.casefold()
        == selected_store.casefold()
    ]

    warehouse_rows = inv[
        inv['_store_name'].map(_is_warehouse)
    ]

    # --------------------------------------------------
    # GROUP BY PART NUMBER + BRAND + MAKE
    # --------------------------------------------------

    matching_keys = [
        '_part_key',
        '_brand_key',
        '_make_key'
    ]

    # Selected store quantity totals
    store_totals = (
        selected_rows
        .groupby(matching_keys)['_qty']
        .sum()
        .to_dict()
    )

    # All warehouse quantity totals
    wh_totals = (
        warehouse_rows
        .groupby(matching_keys)['_qty']
        .sum()
        .to_dict()
    )

    # --------------------------------------------------
    # CREATE SALES MATCHING KEYS
    # --------------------------------------------------

    sales = sales.copy()

    sales['_part_key'] = _clean_key(
        sales[part_col]
    )

    sales['_brand_key'] = _clean_key(
        sales[brand_col]
    )

    sales['_make_key'] = _clean_key(
        sales[make_col]
    )

    # --------------------------------------------------
    # FETCH STORE AND WH QUANTITIES
    # --------------------------------------------------

    sales['_match_key'] = list(
        zip(
            sales['_part_key'],
            sales['_brand_key'],
            sales['_make_key']
        )
    )

    sales['_store_qty'] = [
        store_totals.get(key, 0)
        for key in sales['_match_key']
    ]

    sales['_wh_qty'] = [
        wh_totals.get(key, 0)
        for key in sales['_match_key']
    ]

    # --------------------------------------------------
    # KEEP ONLY ZERO STORE STOCK
    # --------------------------------------------------

    sales = sales[
        sales['_store_qty'] == 0
    ].copy()

    # Get MAKE values
    makes = [
        x
        for x in (
            sales[make_col]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )
        if x
    ]

    if not makes:
        raise ValueError(
            'No MAKE values found after zero-stock filtering.'
        )

    # --------------------------------------------------
    # CREATE EXCEL WORKBOOK
    # --------------------------------------------------

    wb = Workbook()

    default = wb.active
    wb.remove(default)

    yellow = PatternFill(
        'solid',
        fgColor='FFFF00'
    )

    blue = PatternFill(
        'solid',
        fgColor='305496'
    )

    white_bold = Font(
        name='Arial',
        size=10,
        bold=True,
        color='FFFFFF'
    )

    title_font = Font(
        name='Arial',
        size=12,
        bold=True,
        color='000000'
    )

    body_font = Font(
        name='Arial',
        size=10,
        bold=True,
        color='000000'
    )

    thin_black = Side(
        style='thin',
        color='000000'
    )

    cell_border = Border(
        left=thin_black,
        right=thin_black,
        top=thin_black,
        bottom=thin_black
    )

    headers = [
        'PART NUMBER',
        'BRAND',
        'MAKE',
        'LAST 3 MONTHS SALES HAPPEN QUANTITY',
        'STORE',
        'WH'
    ]

    # --------------------------------------------------
    # CREATE SEPARATE SHEET FOR EACH MAKE
    # --------------------------------------------------

    for make in makes:

        safe_name = re.sub(
            r'[\\/*?:\[\]]',
            '_',
            make
        )[:31] or 'MAKE'

        # Avoid duplicate sheet names
        base = safe_name
        name = safe_name
        n = 2

        while name in wb.sheetnames:

            suffix = f'_{n}'

            name = (
                base[:31 - len(suffix)]
                + suffix
            )

            n += 1

        ws = wb.create_sheet(name)

        make_rows = sales[
            sales[make_col]
            .astype(str)
            .str.strip()
            .str.casefold()
            == make.casefold()
        ]

        # TITLE
        ws.merge_cells('A1:F1')

        title_cell = ws['A1']

        title_cell.value = (
            f'{make.upper()} FAST MOVING REPORT'
        )

        title_cell.fill = yellow
        title_cell.font = title_font

        title_cell.alignment = Alignment(
            horizontal='center',
            vertical='center'
        )

        ws.row_dimensions[1].height = 24

        # HEADERS
        for col_idx, header in enumerate(
            headers,
            1
        ):

            cell = ws.cell(
                row=2,
                column=col_idx,
                value=header
            )

            cell.fill = blue
            cell.font = white_bold

            cell.alignment = Alignment(
                horizontal='center',
                vertical='center',
                wrap_text=True
            )

            cell.border = cell_border

        ws.row_dimensions[2].height = 35

        # DATA
        for row_idx, (_, row) in enumerate(
            make_rows.iterrows(),
            3
        ):

            values = [
                row.get(part_col, ''),
                row.get(brand_col, ''),
                row.get(make_col, ''),
                row.get(sales_qty_col, ''),
                row['_store_qty'],
                row['_wh_qty']
            ]

            for col_idx, value in enumerate(
                values,
                1
            ):

                if pd.isna(value):
                    value = ''

                cell = ws.cell(
                    row=row_idx,
                    column=col_idx,
                    value=value
                )

                cell.font = body_font

                cell.alignment = Alignment(
                    horizontal='center',
                    vertical='center'
                )

                cell.border = cell_border

        # COLUMN WIDTHS
        widths = [
            22,  # PART NUMBER
            18,  # BRAND
            18,  # MAKE
            40,  # SALES QUANTITY
            17,  # STORE
            17   # WH
        ]

        for i, width in enumerate(
            widths,
            1
        ):

            ws.column_dimensions[
                get_column_letter(i)
            ].width = width

        # EXCEL SETTINGS
        ws.freeze_panes = 'A3'

        ws.auto_filter.ref = (
            f'A2:F{max(2, ws.max_row)}'
        )

        ws.sheet_view.showGridLines = True

    # SAVE WORKBOOK
    wb.save(output_file)

    return len(makes)
