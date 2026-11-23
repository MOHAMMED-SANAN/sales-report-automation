import re
from pathlib import Path
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


def _norm(value):
    return re.sub(r'[^a-z0-9]+', '', str(value).strip().lower())


def _find_col(df, candidates, required=True):
    lookup = {_norm(c): c for c in df.columns}
    for candidate in candidates:
        key = _norm(candidate)
        if key in lookup:
            return lookup[key]
    # allow partial matches for headers with extra descriptions
    for c in df.columns:
        key = _norm(c)
        if any(_norm(x) in key for x in candidates):
            return c
    if required:
        raise ValueError(f"Could not find required column. Expected one of: {', '.join(candidates)}. Available columns: {list(df.columns)}")
    return None


def _clean_key(series):
    return series.fillna('').astype(str).str.strip().str.upper()


def _is_warehouse(value):
    return bool(re.match(r'^WH(?:_|\b|\s)', str(value).strip(), flags=re.IGNORECASE))


def read_inventory(inventory_file):
    inv = pd.read_excel(inventory_file, dtype=object)
    part_col = _find_col(inv, ['PART NUMBER', 'PART NO', 'ITEM CODE', 'ITEM', 'SKU'])
    make_col = _find_col(inv, ['MAKE', 'MANUFACTURER'])
    store_col = _find_col(inv, ['STORE', 'LOCATION', 'BRANCH', 'WAREHOUSE'])
    # User confirmed Available Qty is column M (13th Excel column).
    if len(inv.columns) < 13:
        raise ValueError('Inventory report has fewer than 13 columns; Available Qty is expected in column M.')
    qty_col = inv.columns[12]
    inv = inv.copy()
    inv['_part_key'] = _clean_key(inv[part_col])
    inv['_make_key'] = _clean_key(inv[make_col])
    inv['_store_name'] = inv[store_col].fillna('').astype(str).str.strip()
    inv['_qty'] = pd.to_numeric(inv[qty_col], errors='coerce').fillna(0)
    inv = inv[(inv['_part_key'] != '') & (inv['_make_key'] != '')]
    stores = sorted({s for s in inv['_store_name'] if s and not _is_warehouse(s)}, key=str.casefold)
    if not stores:
        raise ValueError('No store names found in the inventory Store column (warehouse rows are excluded).')
    return inv, stores, part_col, make_col, store_col, qty_col


def create_fast_moving_report(sales_file, inventory_file, selected_store, output_file):
    sales = pd.read_excel(sales_file, dtype=object)
    part_col = _find_col(sales, ['PART NUMBER', 'PART NO', 'ITEM CODE', 'ITEM', 'SKU'])
    make_col = _find_col(sales, ['MAKE', 'MANUFACTURER'])
    sales_qty_col = _find_col(sales, [
        'LAST 3 MONTH SALES HAPPEN QUANTITY', 'LAST 3 MONTH SALES',
        '3 MONTH SALES', 'SALES HAPPEN QUANTITY', 'SALES QTY', 'QUANTITY'
    ], required=False)
    if sales_qty_col is None:
        # If source contains exactly the first four report fields, use its fourth column.
        if len(sales.columns) >= 4:
            sales_qty_col = sales.columns[3]
        else:
            raise ValueError('Could not identify the last 3 months sales quantity column in the sales report.')

    inv, stores, _, _, _, _ = read_inventory(inventory_file)
    if selected_store not in stores:
        raise ValueError(f"Selected store '{selected_store}' was not found in the inventory report.")

    selected_rows = inv[inv['_store_name'].str.casefold() == selected_store.casefold()]
    warehouse_rows = inv[inv['_store_name'].map(_is_warehouse)]
    store_totals = selected_rows.groupby(['_part_key', '_make_key'])['_qty'].sum().to_dict()
    wh_totals = warehouse_rows.groupby(['_part_key', '_make_key'])['_qty'].sum().to_dict()

    sales = sales.copy()
    sales['_part_key'] = _clean_key(sales[part_col])
    sales['_make_key'] = _clean_key(sales[make_col])
    sales['_store_qty'] = [store_totals.get((p, m), 0) for p, m in zip(sales['_part_key'], sales['_make_key'])]
    sales['_wh_qty'] = [wh_totals.get((p, m), 0) for p, m in zip(sales['_part_key'], sales['_make_key'])]

    # Keep only parts with zero available quantity in the selected store.
    # Then create one worksheet per Make, retaining original row order.
    sales = sales[sales['_store_qty'] == 0].copy()
    makes = [x for x in sales[make_col].dropna().astype(str).str.strip().unique() if x]
    if not makes:
        raise ValueError('No Make values found in the sales report.')

    wb = Workbook()
    default = wb.active
    wb.remove(default)
    yellow = PatternFill('solid', fgColor='FFFF00')
    blue = PatternFill('solid', fgColor='305496')
    white_bold = Font(name='Arial', size=10, bold=True, color='FFFFFF')
    title_font = Font(name='Arial', size=12, bold=True, color='000000')
    body_font = Font(name='Arial', size=10, bold=True, color='000000')
    thin_black = Side(style='thin', color='000000')
    cell_border = Border(left=thin_black, right=thin_black, top=thin_black, bottom=thin_black)
    headers = ['PART NUMBER', 'MAKE', 'LAST 3 MONTH SALES HAPPEN QUANTITY', 'STORE', 'WH']

    for make in makes:
        safe_name = re.sub(r'[\\/*?:\[\]]', '_', make)[:31] or 'MAKE'
        # Avoid duplicate worksheet names after Excel's 31-character limit.
        base, name, n = safe_name, safe_name, 2
        while name in wb.sheetnames:
            suffix = f'_{n}'
            name = base[:31-len(suffix)] + suffix
            n += 1
        ws = wb.create_sheet(name)
        make_rows = sales[sales[make_col].astype(str).str.strip().str.casefold() == make.casefold()]
        ws.merge_cells('A1:E1')
        title_cell = ws['A1']
        title_cell.value = f'{make.upper()} FAST MOVING REPORT'
        title_cell.fill = yellow
        title_cell.font = title_font
        title_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 24
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=2, column=col_idx, value=header)
            cell.fill = blue
            cell.font = white_bold
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = cell_border
        ws.row_dimensions[2].height = 32

        for row_idx, (_, row) in enumerate(make_rows.iterrows(), 3):
            values = [
                row.get(part_col, ''),
                row.get(make_col, ''),
                row.get(sales_qty_col, ''),
                row['_store_qty'],
                row['_wh_qty'],
            ]
            for col_idx, value in enumerate(values, 1):
                if pd.isna(value):
                    value = ''
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = body_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = cell_border

        widths = [20, 16, 38, 16, 16]
        for i, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        ws.freeze_panes = 'A3'
        ws.auto_filter.ref = f'A2:E{max(2, ws.max_row)}'
        ws.sheet_view.showGridLines = True

    wb.save(output_file)
    return len(makes)
