# ==========================================================
# DELIVERY DATE FILTER
# ONLY APPEARS WHEN DELIVERY REPORT IS SELECTED
# ==========================================================

delivery_start_date = None
delivery_end_date = None


if delivery_selected:

    st.divider()

    st.subheader(
        "🚚 Delivery Report Filters"
    )

    try:

        preview_df = pd.read_excel(
            uploaded_file
        )


        if "Order Updated At" not in preview_df.columns:

            st.error(
                "Delivery Report requires "
                "'Order Updated At' column."
            )

        else:

            # IMPORTANT:
            # Your Excel date format is DD-MM-YYYY
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


                if delivery_start_date > delivery_end_date:

                    st.error(
                        "From Date cannot be greater than To Date."
                    )


    except Exception as e:

        st.error(
            f"Unable to read delivery dates: {e}"
        )
