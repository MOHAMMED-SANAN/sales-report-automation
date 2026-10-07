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


# ==============================================================
# FILE UPLOAD
# ==============================================================

uploaded_file = st.file_uploader(
    "Upload Excel Sales Report",
    type=["xlsx", "xls"]
)


if uploaded_file is not None:

    st.success(
        f"File uploaded successfully: {uploaded_file.name}"
    )

    st.divider()

    st.subheader("📋 Select Reports")


    # ==========================================================
    # REPORT SELECTION
    # ==========================================================

    col1, col2, col3, col4, col5, col6 = st.columns(6)


    # ==========================================================
    # TARGET PERCENTAGE
    # ==========================================================

    with col1:

        st.markdown("### 🎯 Target Percentage")

        st.write(
            "Generate the target percentage report."
        )

        target_selected = st.checkbox(
            "Select Target Percentage Report"
        )


    # ==========================================================
    # DAY-WISE BRANCH
    # ==========================================================

    with col2:

        st.markdown("### 📅 Day-wise Branch")

        st.write(
            "Generate the day-wise sales report for each branch."
        )

        daywise_selected = st.checkbox(
            "Select Day-wise Branch Report"
        )


    # ==========================================================
    # SALESMAN TARGET
    # ==========================================================

    with col3:

        st.markdown("### 👨‍💼 Salesman Target")

        st.write(
            "Generate the salesman daily target report."
        )

        salesman_selected = st.checkbox(
            "Select Salesman Daily Target Report"
        )


    # ==========================================================
    # MABELLAH SALES
    # ==========================================================

    with col4:

        st.markdown("### 🏪 Mabellah Sales")

        st.write(
            "Generate the MABELLAH daily sales and quotation report."
        )

        mabellah_selected = st.checkbox(
            "Select Mabellah Sales Report"
        )


    # ==========================================================
    # PICK / TRANSFER
    # ==========================================================

    with col5:

        st.markdown("### 📦 Pick / Transfer")

        st.write(
            "Generate the store-wise Pick and Internal Transfer report."
        )

        pick_transfer_selected = st.checkbox(
            "Select Pick / Internal Transfer Report"
        )


    # ==========================================================
    # DELIVERY REPORT
    # ==========================================================

    with col6:

        st.markdown("### 🚚 Deliveries")

        st.write(
            "Generate the store-wise driver delivery report."
        )

        delivery_selected = st.checkbox(
            "Select Store Wise Delivery Report"
        )


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
                    errors="coerce"
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
                            min_value=min_date,
                            max_value=max_date,
                            key="delivery_start_date"
                        )


                    with date_col2:

                        delivery_end_date = st.date_input(
                            "To Date",
                            value=max_date,
                            min_value=min_date,
                            max_value=max_date,
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
            delivery_selected
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
            delivery_selected
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
