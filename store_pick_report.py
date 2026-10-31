
    for store_name in store_list:

        sheet_name = str(
            store_name
        )[:31]


        if sheet_name not in wb.sheetnames:

            continue


        ws = wb[sheet_name]


        store_title = (

            f"{store_name} - "
            "PICK / INTERNAL TRANSFER REPORT"

        )


        format_report_sheet(

            ws,

            store_title

        )


    # ========================================================
    # 35. SAVE
    # ========================================================

    wb.save(
        output_file
    )


    # ========================================================
    # 36. SUCCESS MESSAGE
    # ========================================================

    print(
        "\n=============================================="
    )

    print(
        "REPORT CREATED SUCCESSFULLY"
    )

    print(
        "=============================================="
    )

    print(
        f"Output file: {output_file}"
    )

    print(
        "\nCreated sheets:"
    )

    print(
        " - Master Data"
    )


    for store_name in store_list:

        print(
            f" - {store_name}"
        )

