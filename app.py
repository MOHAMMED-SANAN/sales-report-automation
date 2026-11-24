
import tempfile
from pathlib import Path

import streamlit as st
import pandas as pd

from fast_moving_report import (
    read_inventory,
    create_fast_moving_report
)

# TARGET REPORT
try:
    from target_report import (
        target_newfile,
        target_percenntage,
        format_report
    )
except Exception:
    target_newfile = target_percenntage = format_report = None

# DAYWISE REPORT
try:
    from daywise_report import split_date_into_columns
except Exception:
    split_date_into_columns = None

# SALESMAN REPORT
try:
    from salesman_report import (
        salesman_newfile,
        sales_person_report,
        create_excel_report
    )
except Exception:
    salesman_newfile = sales_person_report = create_excel_report = None

# MABELLah REPORT
try:
    from mabellah_sales_report import generate_mabellah_sales_report
except Exception:
    generate_mabellah_sales_report = None

# STORE PICK REPORT
try:
    from store_pick_report import generate_store_pick_report
except Exception:
    generate_store_pick_report = None

# DELIVERY REPORT
try:
    from delivery_report import generate_delivery_report
except Exception:
    generate_delivery_report = None


st.set_page_config(
    page_title="Sales Report Automation",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Sales Report Automation")

st.write(
    "Select one report first, then upload the required Excel files."
)

report_options = [
    "Fast Moving Report",
    "Target Percentage Report",
    "Day-wise Branch Report",
    "Salesman Daily Target Report",
    "Mabellah Sales Report",
    "Store Wise Pick / Internal Transfer Report",
    "Store Wise Delivery Report"
]

selected_report = st.selectbox(
    "Choose which report you want to generate",
    report_options,
    index=0
)

st.divider()

EXCEL_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "spreadsheetml.sheet"
)


# ==================================================
# FAST MOVING REPORT
# ==================================================

if selected_report == "Fast Moving Report":

    st.subheader("⚡ Fast Moving Report")

    st.info(
        "Upload the Sales Report and Inventory Report. "
        "Select a store to generate separate MAKE-wise sheets. "
        "Only items with zero stock in the selected store "
        "will be included."
    )

    sales_upload = st.file_uploader(
        "1. Upload Fast Moving Sales Report",
        type=["xlsx", "xls"],
        key="fm_sales"
    )

    inventory_upload = st.file_uploader(
        "2. Upload Inventory Report",
        type=["xlsx", "xls"],
        key="fm_inventory"
    )

    selected_store = None

    if inventory_upload is not None:
        try:
            inv_df, stores, *_ = read_inventory(
                inventory_upload
            )

            selected_store = st.selectbox(
                "3. Select Store for Fast Moving Report",
                stores,
                key="fm_store"
            )

            st.caption(
                f"{len(stores)} stores found. "
                "WH warehouses are excluded from the store dropdown."
            )

        except Exception as exc:
            st.error(
                f"Unable to read Inventory Report: {exc}"
            )

    if st.button(
        "🚀 Generate Fast Moving Report",
        type="primary",
        use_container_width=True,
        key="generate_fast_moving"
    ):

        if sales_upload is None or inventory_upload is None:
            st.error(
                "Please upload both Sales Report and Inventory Report."
            )

        elif not selected_store:
            st.error("Please select a store.")

        else:
            try:
                with st.spinner(
                    "Generating MAKE-wise Fast Moving Report..."
                ):

                    workdir = Path(
                        tempfile.mkdtemp(prefix="fast_moving_")
                    )

                    sales_path = workdir / sales_upload.name
                    inventory_path = workdir / inventory_upload.name

                    safe_store = "".join(
                        c if c.isalnum() or c in "-_"
                        else "_"
                        for c in selected_store
                    )

                    output_path = (
                        workdir /
                        f"{safe_store}_Fast_Moving_Report.xlsx"
                    )

                    sales_path.write_bytes(
                        sales_upload.getvalue()
                    )

                    inventory_path.write_bytes(
                        inventory_upload.getvalue()
                    )

                    sheet_count = create_fast_moving_report(
                        sales_path,
                        inventory_path,
                        selected_store,
                        output_path
                    )

                st.success(
                    f"Fast Moving Report created successfully! "
                    f"{sheet_count} MAKE-wise sheets generated."
                )

                st.download_button(
                    "📥 Download Fast Moving Report",
                    data=output_path.read_bytes(),
                    file_name=output_path.name,
                    mime=EXCEL_MIME,
                    use_container_width=True
                )

            except Exception as exc:
                st.error(
                    f"Fast Moving Report generation failed: {exc}"
                )


# ==================================================
# OTHER REPORTS
# ==================================================

else:

    st.subheader(selected_report)

    uploaded_file = st.file_uploader(
        "Upload Excel Sales Report",
        type=["xlsx", "xls"],
        key="regular_sales"
    )

    delivery_start_date = None
    delivery_end_date = None
    date_error = None

    # DELIVERY DATE SELECTION
    if (
        uploaded_file is not None
        and selected_report == "Store Wise Delivery Report"
    ):

        try:
            preview = pd.read_excel(uploaded_file)

            if "Order Updated At" not in preview.columns:

                date_error = (
                    "Delivery Report requires "
                    "'Order Updated At' column."
                )

            else:

                dates = pd.to_datetime(
                    preview["Order Updated At"],
                    errors="coerce",
                    dayfirst=True
                ).dropna()

                if dates.empty:

                    date_error = (
                        "No valid dates found in "
                        "'Order Updated At'."
                    )

                else:

                    c1, c2 = st.columns(2)

                    with c1:
                        delivery_start_date = st.date_input(
                            "From Date",
                            value=dates.min().date(),
                            key="delivery_from"
                        )

                    with c2:
                        delivery_end_date = st.date_input(
                            "To Date",
                            value=dates.max().date(),
                            key="delivery_to"
                        )

                    if delivery_start_date > delivery_end_date:
                        date_error = (
                            "From Date cannot be greater than To Date."
                        )

            if date_error:
                st.error(date_error)

        except Exception as exc:
            st.error(
                f"Unable to read delivery dates: {exc}"
            )

    # GENERATE OTHER REPORTS
    if (
        uploaded_file is not None
        and st.button(
            "🚀 Generate Report",
            type="primary",
            use_container_width=True,
            key="generate_regular"
        )
    ):

        workdir = Path(
            tempfile.mkdtemp(prefix="sales_report_")
        )

        input_path = workdir / uploaded_file.name

        input_path.write_bytes(
            uploaded_file.getvalue()
        )

        output_path = None

        try:

            with st.spinner(
                f"Generating {selected_report}..."
            ):

                if selected_report == "Target Percentage Report":

                    if not all([
                        target_newfile,
                        target_percenntage,
                        format_report
                    ]):
                        raise ImportError(
                            "target_report.py is missing or incomplete."
                        )

                    df_target = target_newfile(str(input_path))

                    target_df = target_percenntage(df_target)

                    report_date = pd.to_datetime(
                        df_target["DATE"],
                        errors="coerce"
                    ).max().normalize()

                    output_path = (
                        workdir / "Target_Percentage_Report.xlsx"
                    )

                    format_report(
                        target_df,
                        str(output_path),
                        report_date
                    )

                elif selected_report == "Day-wise Branch Report":

                    if not split_date_into_columns:
                        raise ImportError(
                            "daywise_report.py is missing."
                        )

                    output_path = (
                        workdir / "Daywise_Sales_Report.xlsx"
                    )

                    split_date_into_columns(
                        str(input_path),
                        str(output_path)
                    )

                elif selected_report == "Salesman Daily Target Report":

                    if not all([
                        salesman_newfile,
                        sales_person_report,
                        create_excel_report
                    ]):
                        raise ImportError(
                            "salesman_report.py is missing or incomplete."
                        )

                    df = salesman_newfile(str(input_path))

                    report_df, report_date = sales_person_report(df)

                    output_path = (
                        workdir / "Sales_Man_Daily_Target_Review.xlsx"
                    )

                    create_excel_report(
                        report_df,
                        report_date,
                        str(output_path)
                    )

                elif selected_report == "Mabellah Sales Report":

                    if not generate_mabellah_sales_report:
                        raise ImportError(
                            "mabellah_sales_report.py is missing."
                        )

                    output_path = (
                        workdir / "Mabellah_Sales_Report.xlsx"
                    )

                    generate_mabellah_sales_report(
                        str(input_path),
                        str(output_path)
                    )

                elif selected_report == "Store Wise Pick / Internal Transfer Report":

                    if not generate_store_pick_report:
                        raise ImportError(
                            "store_pick_report.py is missing."
                        )

                    output_path = (
                        workdir / "Store_Wise_Pick_Count_Report.xlsx"
                    )

                    generate_store_pick_report(
                        str(input_path),
                        str(output_path)
                    )

                elif selected_report == "Store Wise Delivery Report":

                    if not generate_delivery_report:
                        raise ImportError(
                            "delivery_report.py is missing."
                        )

                    if (
                        date_error
                        or delivery_start_date is None
                        or delivery_end_date is None
                    ):
                        raise ValueError(
                            date_error
                            or "Please select valid delivery dates."
                        )

                    output_path = (
                        workdir / "Store_Wise_Delivery_Report.xlsx"
                    )

                    generate_delivery_report(
                        input_file=str(input_path),
                        output_file=str(output_path),
                        start_date=delivery_start_date,
                        end_date=delivery_end_date
                    )

            if output_path is None or not output_path.exists():
                raise FileNotFoundError(
                    "The report output file was not created."
                )

            st.success("Report generated successfully!")

            st.download_button(
                "📥 Download Report",
                data=output_path.read_bytes(),
                file_name=output_path.name,
                mime=EXCEL_MIME,
                use_container_width=True
            )

        except Exception as exc:

            st.error(
                f"{selected_report} generation failed: {exc}"
            )
