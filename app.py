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
