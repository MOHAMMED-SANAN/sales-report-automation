import os
import tempfile
from pathlib import Path

import streamlit as st
import pandas as pd

from target_report import (
    target_newfile,
    target_percenntage,
    format_report
)

from daywise_report import (
    split_date_into_columns
)

from salesman_report import (
    salesman_newfile,
    sales_person_report,
    create_excel_report
)

from mabellah_sales_report import (
    generate_mabellah_sales_report
)

from store_pick_report import (
    generate_store_pick_report
)

# NEW DELIVERY REPORT
from delivery_report import (
    generate_delivery_report
)


def _clean_column_name(value):
    return " ".join(str(value).strip().lower().replace("_", " ").split())


def _find_column(df, possible_names, required=True):
    lookup = {_clean_column_name(c): c for c in df.columns}
    for name in possible_names:
        found = lookup.get(_clean_column_name(name))
        if found is not None:
            return found
    if required:
        raise ValueError(
            "Missing required column. Expected one of: " + ", ".join(possible_names)
            + ". Available columns: " + ", ".join(map(str, df.columns))
        )
    return None


def generate_fast_moving_report(main_file, inventory_file, selected_store, output_file):
    """Add STORE and WH Available Qty to the main fast-moving report.

    Matching key is Part Number + Make. Brand is deliberately ignored.
    Duplicate inventory rows are summed.
    """
    main = pd.read_excel(main_file)
    inventory = pd.read_excel(inventory_file)

    main_part = _find_column(main, ["Part Number", "Part No", "Part#", "Part No."])
    main_make = _find_column(main, ["Make", "Manufacturer"])
    inv_part = _find_column(inventory, ["Part Number", "Part No", "Part#", "Part No."])
    inv_make = _find_column(inventory, ["Make", "Manufacturer"])
    store_col = _find_column(inventory, ["Store", "Branch", "Location", "Store Name"])
    qty_col = _find_column(inventory, ["Available Qty", "Available Quantity", "Avail Qty"])

    def key_part(series):
        return series.fillna("").astype(str).str.strip().str.upper()

    main_keys = key_part(main[main_part]) + "||" + key_part(main[main_make])
    inv = inventory.copy()
    inv["__part_key"] = key_part(inv[inv_part])
    inv["__make_key"] = key_part(inv[inv_make])
    inv["__key"] = inv["__part_key"] + "||" + inv["__make_key"]
    inv["__qty"] = pd.to_numeric(inv[qty_col], errors="coerce").fillna(0)
    stores = inv[store_col].fillna("").astype(str).str.strip()
    is_wh = stores.str.match(r"^WH(?:_|\b|\s)", case=False)

    store_totals = (inv.loc[stores.str.casefold() == str(selected_store).strip().casefold()]
                    .groupby("__key")["__qty"].sum())
    wh_totals = inv.loc[is_wh].groupby("__key")["__qty"].sum()

    main["STORE"] = main_keys.map(store_totals).fillna(0)
    main["WH"] = main_keys.map(wh_totals).fillna(0)
    main.to_excel(output_file, index=False)
    return main


# ==============================================================
# PAGE CONFIG
# ==============================================================

st.set_page_config(
    page_title="Sales Report Automation",
    page_icon="📊",
    layout="wide"
)


# ==============================================================
# TITLE
# ==============================================================

st.title("📊 Sales Report Automation")

st.write(
    "Upload your Excel sales report and select the reports "
    "you want to generate."
)


# =============================================================
# STEP 1: SELECT REPORT FIRST
# =============================================================

report_options = {
    "🎯 Target Percentage Report": "target",
    "📅 Day-wise Branch Report": "daywise",
    "👨‍💼 Salesman Daily Target Report": "salesman",
    "🏪 Mabellah Sales Report": "mabellah",
    "📦 Pick / Internal Transfer Report": "pick",
    "🚚 Store Wise Delivery Report": "delivery",
    "⚡ Fast Moving Report": "fast_moving",
}

st.subheader("1️⃣ Choose the report you want")
selected_label = st.selectbox(
    "Select Report",
    options=list(report_options.keys()),
    index=None,
    placeholder="Choose a report...",
    key="report_selector"
)
selected_report = report_options.get(selected_label) if selected_label else None

uploaded_file = None
inventory_file = None
selected_store = None

if selected_report:
    st.divider()
    st.subheader("2️⃣ Upload required Excel file(s)")

    if selected_report == "fast_moving":
        upload_col1, upload_col2 = st.columns(2)
        with upload_col1:
            uploaded_file = st.file_uploader(
                "Upload Fast Moving Sales Report", type=["xlsx", "xls"], key="fast_main_upload"
            )
        with upload_col2:
            inventory_file = st.file_uploader(
                "Upload Inventory Report", type=["xlsx", "xls"], key="fast_inventory_upload"
            )

        if inventory_file is not None:
            try:
                inventory_preview = pd.read_excel(inventory_file)
                inventory_store_col = _find_column(
                    inventory_preview, ["Store", "Branch", "Location", "Store Name"]
                )
                all_store_names = (
                    inventory_preview[inventory_store_col].dropna().astype(str).str.strip().unique().tolist()
                )
                store_names = sorted([
                    name for name in all_store_names
                    if not pd.Series([name]).str.match(r"^WH(?:_|\b|\s)", case=False).iloc[0]
                ])
                if store_names:
                    selected_store = st.selectbox("Select Store", store_names, key="fast_store_select")
                else:
                    st.error("No store names found in the inventory report (warehouse rows start with WH).")
            except Exception as exc:
                st.error(f"Unable to read Inventory Report: {exc}")

    else:
        uploaded_file = st.file_uploader(
            "Upload Excel Sales Report", type=["xlsx", "xls"], key=f"sales_upload_{selected_report}"
        )

if selected_report and uploaded_file is not None:
    st.success(f"Sales report uploaded: {uploaded_file.name}")
    if selected_report == "fast_moving" and inventory_file is not None and selected_store:
        st.success(f"Inventory uploaded: {inventory_file.name} | Store: {selected_store}")

    st.divider()

    target_selected = selected_report == "target"
    daywise_selected = selected_report == "daywise"
    salesman_selected = selected_report == "salesman"
    mabellah_selected = selected_report == "mabellah"
    pick_transfer_selected = selected_report == "pick"
    delivery_selected = selected_report == "delivery"
    fast_moving_selected = selected_report == "fast_moving"

    # ==========================================================
    # DELIVERY DATE FILTER
    #
    # THIS FILTER APPEARS ONLY WHEN DELIVERY REPORT IS SELECTED
    # ==========================================================

    delivery_start_date = None
    delivery_end_date = None


    if delivery_selected:

        st.divider()

        st.subheader(
            "🚚 Delivery Report Filters"
        )

        try:

            # --------------------------------------------------
            # Read file only to get Order Updated At dates
            # --------------------------------------------------

            preview_df = pd.read_excel(
                uploaded_file
            )


            # --------------------------------------------------
            # Check column
            # --------------------------------------------------

            if "Order Updated At" not in preview_df.columns:

                st.error(
                    "Delivery Report requires "
                    "'Order Updated At' column."
                )

            else:

                preview_df["Order Updated At"] = pd.to_datetime(
                    preview_df["Order Updated At"],
                    errors="coerce",
                    dayfirst=True
                )


                valid_dates = preview_df[
                    "Order Updated At"
                ].dropna()


                if valid_dates.empty:

                    st.error(
                        "No valid dates found in "
                        "'Order Updated At'."
                    )

                else:

                    min_date = valid_dates.min().date()
                    max_date = valid_dates.max().date()


                    # --------------------------------------------------
                    # DATE SELECTION
                    # --------------------------------------------------

                    date_col1, date_col2 = st.columns(2)


                    with date_col1:

                        delivery_start_date = st.date_input(
                            "From Date",
                            value=min_date,
                            key="delivery_start_date"
                        )


                    with date_col2:

                        delivery_end_date = st.date_input(
                            "To Date",
                            value=max_date,
                            key="delivery_end_date"
                        )


                    # --------------------------------------------------
                    # INVALID DATE CHECK
                    # --------------------------------------------------

                    if delivery_start_date > delivery_end_date:

                        st.error(
                            "From Date cannot be greater than To Date."
                        )


        except Exception as e:

            st.error(
                f"Unable to read delivery dates: {e}"
            )


    st.divider()


    # ==========================================================
    # GENERATE BUTTON
    # ==========================================================

    generate = st.button(
        "🚀 Generate Selected Reports",
        type="primary",
        use_container_width=True
    )


    # ==========================================================
    # GENERATE
    # ==========================================================

    if generate:


        # ======================================================
        # CHECK REPORT SELECTION
        # ======================================================

        if not any([
            target_selected,
            daywise_selected,
            salesman_selected,
            mabellah_selected,
            pick_transfer_selected,
            delivery_selected,
            fast_moving_selected
        ]):

            st.warning(
                "Please select at least one report."
            )

            st.stop()


        # ======================================================
        # CHECK DELIVERY DATES
        # ======================================================

        if delivery_selected:

            if (
                delivery_start_date is None
                or delivery_end_date is None
            ):

                st.error(
                    "Please select Delivery Report dates."
                )

                st.stop()


            if delivery_start_date > delivery_end_date:

                st.error(
                    "From Date cannot be greater than To Date."
                )

                st.stop()


        if fast_moving_selected and (inventory_file is None or not selected_store):
            st.error("Please upload the Inventory Report and select a store.")
            st.stop()

        # ======================================================
        # TEMP WORKING DIRECTORY
        # ======================================================

        workdir = Path(
            tempfile.mkdtemp(
                prefix="sales_report_"
            )
        )


        # ======================================================
        # SAVE UPLOADED FILE
        # ======================================================

        input_path = (
            workdir / uploaded_file.name
        )


        input_path.write_bytes(
            uploaded_file.getvalue()
        )

        inventory_path = None
        if fast_moving_selected:
            inventory_path = workdir / inventory_file.name
            inventory_path.write_bytes(inventory_file.getvalue())


        # ======================================================
        # GENERATED REPORTS
        # ======================================================

        generated = []

        errors = []


        # ======================================================
        # NUMBER OF SELECTED REPORTS
        # ======================================================

        selected_count = sum([
            target_selected,
            daywise_selected,
            salesman_selected,
            mabellah_selected,
            pick_transfer_selected,
            delivery_selected,
            fast_moving_selected
        ])


        progress = st.progress(0)

        done = 0


        # ======================================================
        # TARGET PERCENTAGE
        # ======================================================

        if target_selected:

            try:

                with st.spinner(
                    "Generating Target Percentage Report..."
                ):

                    df_target = target_newfile(
                        str(input_path)
                    )


                    target_df = target_percenntage(
                        df_target
                    )


                    report_date = (
                        pd.to_datetime(
                            df_target["DATE"],
                            errors="coerce"
                        )
                        .max()
                        .normalize()
                    )


                    output = (
                        workdir
                        / "Target_Percentage_Report.xlsx"
                    )


                    format_report(
                        target_df,
                        str(output),
                        report_date
                    )


                    generated.append(
                        (
                            "🎯 Target Percentage Report",
                            output
                        )
                    )


            except Exception as e:

                errors.append(
                    f"Target Percentage Report: {e}"
                )


            done += 1

            progress.progress(
                done / selected_count
            )


        # ======================================================
        # DAY-WISE BRANCH
        # ======================================================

        if daywise_selected:

            try:

                with st.spinner(
                    "Generating Day-wise Branch Report..."
                ):

                    output = (
                        workdir
                        / "Daywise_Sales_Report.xlsx"
                    )


                    split_date_into_columns(
                        str(input_path),
                        str(output)
                    )


                    generated.append(
                        (
                            "📅 Day-wise Branch Report",
                            output
                        )
                    )


            except Exception as e:

                errors.append(
                    f"Day-wise Branch Report: {e}"
                )


            done += 1

            progress.progress(
                done / selected_count
            )


        # ======================================================
        # SALESMAN DAILY TARGET
        # ======================================================

        if salesman_selected:

            try:

                with st.spinner(
                    "Generating Salesman Daily Target Report..."
                ):

                    df_salesman = salesman_newfile(
                        str(input_path)
                    )


                    salesman_df, salesman_date = (
                        sales_person_report(
                            df_salesman
                        )
                    )


                    output = (
                        workdir
                        / "Sales_Man_Daily_Target_Review.xlsx"
                    )


                    create_excel_report(
                        salesman_df,
                        salesman_date,
                        str(output)
                    )


                    generated.append(
                        (
                            "👨‍💼 Salesman Daily Target Report",
                            output
                        )
                    )


            except Exception as e:

                errors.append(
                    f"Salesman Daily Target Report: {e}"
                )


            done += 1

            progress.progress(
                done / selected_count
            )


        # ======================================================
        # MABELLAH SALES REPORT
        # ======================================================

        if mabellah_selected:

            try:

                with st.spinner(
                    "Generating MABELLAH Sales Report..."
                ):

                    output = (
                        workdir
                        / "Mabellah_Sales_Report.xlsx"
                    )


                    generate_mabellah_sales_report(
                        str(input_path),
                        str(output)
                    )


                    generated.append(
                        (
                            "🏪 Mabellah Sales Report",
                            output
                        )
                    )


            except Exception as e:

                errors.append(
                    f"Mabellah Sales Report: {e}"
                )


            done += 1

            progress.progress(
                done / selected_count
            )


        # ======================================================
        # STORE WISE PICK / INTERNAL TRANSFER REPORT
        # ======================================================

        if pick_transfer_selected:

            try:

                with st.spinner(
                    "Generating Store Wise Pick / Internal Transfer Report..."
                ):

                    output = (
                        workdir
                        / "Store_Wise_Pick_Count_Report.xlsx"
                    )


                    generate_store_pick_report(
                        str(input_path),
                        str(output)
                    )


                    generated.append(
                        (
                            "📦 Store Wise Pick / Internal Transfer Report",
                            output
                        )
                    )


            except Exception as e:

                errors.append(
                    f"Store Wise Pick / Internal Transfer Report: {e}"
                )


            done += 1

            progress.progress(
                done / selected_count
            )


        # ======================================================
        # STORE WISE DELIVERY REPORT
        # ======================================================

        if delivery_selected:

            try:

                with st.spinner(
                    "Generating Store Wise Delivery Report..."
                ):

                    output = (
                        workdir
                        / "Store_Wise_Delivery_Report.xlsx"
                    )


                    generate_delivery_report(
                        input_file=str(input_path),
                        output_file=str(output),
                        start_date=delivery_start_date,
                        end_date=delivery_end_date
                    )


                    generated.append(
                        (
                            "🚚 Store Wise Delivery Report",
                            output
                        )
                    )


            except Exception as e:

                errors.append(
                    f"Store Wise Delivery Report: {e}"
                )


            done += 1

            progress.progress(
                done / selected_count
            )


        # ======================================================
        # FAST MOVING REPORT
        # ======================================================

        if fast_moving_selected:
            try:
                with st.spinner("Generating Fast Moving Report..."):
                    output = workdir / "Fast_Moving_Report.xlsx"
                    generate_fast_moving_report(
                        str(input_path), str(inventory_path), selected_store, str(output)
                    )
                    generated.append(("⚡ Fast Moving Report", output))
            except Exception as e:
                errors.append(f"Fast Moving Report: {e}")


        # ======================================================
        # ERRORS
        # ======================================================

        if errors:

            for err in errors:

                st.error(err)


        # ======================================================
        # DOWNLOAD REPORTS
        # ======================================================

        if generated:

            st.success(
                "✅ Report generation completed!"
            )


            st.subheader(
                "📥 Download Reports"
            )


            for label, path in generated:

                st.download_button(

                    label=f"Download {label}",

                    data=path.read_bytes(),

                    file_name=path.name,

                    mime=(
                        "application/"
                        "vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),

                    key=f"download_{path.name}",

                    use_container_width=True
                )
