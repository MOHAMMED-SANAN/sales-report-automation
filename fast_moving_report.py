from __future__ import annotations

from io import BytesIO
import re

import pandas as pd


def _normalise_column(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def _find_column(columns, candidates, label):
    lookup = {_normalise_column(col): col for col in columns}
    for candidate in candidates:
        key = _normalise_column(candidate)
        if key in lookup:
            return lookup[key]
    raise ValueError(
        f"Could not find the {label} column. Expected one of: "
        + ", ".join(candidates)
    )


def _read_inventory(inventory_source):
    df = pd.read_excel(inventory_source)
    store_col = _find_column(
        df.columns, ["Store", "Branch", "Store Name", "Warehouse"], "Store"
    )
    part_col = _find_column(df.columns, ["Part Number", "Part No", "Item Code"], "Part Number")
    make_col = _find_column(df.columns, ["Make"], "Make")
    qty_col = _find_column(
        df.columns, ["Available Qty", "Available Quantity", "Available"], "Available Qty"
    )
    return df, store_col, part_col, make_col, qty_col


def get_inventory_store_options(inventory_bytes: bytes) -> list[str]:
    """Return store names, excluding warehouse rows whose Store starts with WH."""
    df, store_col, *_ = _read_inventory(BytesIO(inventory_bytes))
    names = df[store_col].dropna().astype(str).str.strip()
    names = names[names.ne("")]
    names = names[~names.str.upper().str.match(r"^WH(?:_|\\b|\\s)", na=False)]
    return sorted(names.unique().tolist(), key=str.casefold)


def _key_part(series):
    return series.fillna("").astype(str).str.strip().str.upper()


def generate_fast_moving_report(
    main_file: str,
    inventory_file: str,
    selected_store: str,
    output_file: str,
) -> None:
    """
    Fill STORE from the selected store and WH from all Store values beginning
    with WH. Both quantities use Available Qty. Match key: Part Number + Make.
    """
    main_df = pd.read_excel(main_file)
    inv_df, store_col, inv_part, inv_make, qty_col = _read_inventory(inventory_file)

    main_part = _find_column(main_df.columns, ["Part Number", "Part No"], "main Part Number")
    main_make = _find_column(main_df.columns, ["Make"], "main Make")

    # Match on Part Number + Make only. Brand is intentionally ignored.
    inv_keys = pd.DataFrame({
        "_part": _key_part(inv_df[inv_part]),
        "_make": _key_part(inv_df[inv_make]),
        "_store": inv_df[store_col].fillna("").astype(str).str.strip(),
        "_qty": pd.to_numeric(inv_df[qty_col], errors="coerce").fillna(0),
    })

    store_rows = inv_keys[
        inv_keys["_store"].str.casefold() == str(selected_store).strip().casefold()
    ]
    store_totals = (
        store_rows.groupby(["_part", "_make"], dropna=False)["_qty"]
        .sum()
        .to_dict()
    )

    wh_rows = inv_keys[
        inv_keys["_store"].str.upper().str.match(r"^WH(?:_|\\b|\\s)", na=False)
    ]
    wh_totals = (
        wh_rows.groupby(["_part", "_make"], dropna=False)["_qty"]
        .sum()
        .to_dict()
    )

    main_keys = list(zip(
        _key_part(main_df[main_part]),
        _key_part(main_df[main_make]),
    ))
    main_df["STORE"] = [store_totals.get(key, 0) for key in main_keys]
    main_df["WH"] = [wh_totals.get(key, 0) for key in main_keys]

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        main_df.to_excel(writer, index=False, sheet_name="Fast Moving Report")
