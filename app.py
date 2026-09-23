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


st.set_page_config(
    page_title="Sales Report Automation",
    page_icon="📊",
    layout="wide"
)


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

    col1, col2, col3, col4 = st.columns(4)


    with col1:

        st.markdown("### 🎯 Target Percentage")

        st.write(
            "Generate the target percentage report."
        )

        target_selected = st.checkbox(
            "Select Target Percentage Report"
        )


    with col2:

        st.markdown("### 📅 Day-wise Branch")

        st.write(
            "Generate the day-wise sales report for each branch."
        )

        daywise_selected = st.checkbox(
            "Select Day-wise Branch Report"
        )


    with col3:

        st.markdown("### 👨‍💼 Salesman Target")

        st.write(
            "Generate the salesman daily target report."
        )

        salesman_selected = st.checkbox(
            "Select Salesman Daily Target Report"
        )


    with col4:

        st.markdown("### 🏪 Mabellah Sales")

        st.write(
            "Generate the MABELLAH daily sales and quotation report."
        )

        mabellah_selected = st.checkbox(
            "Select Mabellah Sales Report"
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


    if generate:

        if not any([
            target_selected,
            daywise_selected,
            salesman_selected,
            mabellah_selected
        ]):

            st.warning(
                "Please select at least one report."
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


        input_path = (
            workdir / uploaded_file.name
        )


        input_path.write_bytes(
            uploaded_file.getvalue()
        )


        generated = []

        errors = []


        selected_count = sum([
            target_selected,
            daywise_selected,
            salesman_selected,
            mabellah_selected
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
        # SHOW ERRORS
        # ======================================================

        if errors:

            for err in errors:

                st.error(err)


        # ======================================================
        # DOWNLOADS
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

                    label=(
                        f"Download {label}"
                    ),

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
