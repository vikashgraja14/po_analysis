from analyze_excelfinal import *
import base64
def tab6(RejactionRemarks):
    RejactionRemarks = RejactionRemarks[
        RejactionRemarks["Reason for Rejection"].notna()
    ]
    filtered_df = RejactionRemarks.copy()
    filtered_df.rename(
        columns={
            "Vendor Name": "Name",
            "Pstng Date": "Posting Date",
            "Type": "Doc.Type",
            "Vendor": "ReimbursementID",
            "Cost Ctr": "Cost Center",
            "Doc. Date": "Document Date",
            "Verified by": "Verifier ID",
            "Created": "Creator ID",
        },
        inplace=True,
    )
    filtered_df.sort_values(
        by=["ReimbursementID", "Amount", "Cost Center"], inplace=True
    )

    # Converting specific columns to strings
    columns_to_convert = [
        "Payable req.no",
        "Doc.Type",
        "Invoice Number",
        "Text",
        "Cost Center",
        "G/L",
        "Document No",
        "Creator ID",
        "Verifier ID",
        "Reason for Rejection",
        "HOG Approval by",
    ]
    filtered_df[columns_to_convert] = filtered_df[columns_to_convert].astype(str)

    # Select box for year and filtering data
    colu1, colu2, _, colu3, _ = st.columns([1, 4, 1, 4, 1])
    selected_year = colu1.selectbox(
        "Select a year", ["All"] + filtered_df["year"].unique().tolist()
    )
    if selected_year != "All":
        filtered_df = filtered_df[filtered_df["year"] == selected_year]

    # Displaying total amount allotted
    with colu2:
        Total_Amount_Alloted = filtered_df["Amount"].sum()
        total_amount_str = str(Total_Amount_Alloted)
        if len(total_amount_str) > 5:
            integer_part = int(float(total_amount_str))
            integer_length = len(str(integer_part))
            if 5 < integer_length <= 7:
                Total_Amount_Alloted /= 100000
                amount_display = f"₹ {Total_Amount_Alloted:,.2f} lakhs"
            elif integer_length > 7:
                Total_Amount_Alloted /= 10000000
                amount_display = f"₹ {Total_Amount_Alloted:,.2f} crores"
            else:
                amount_display = f"₹ {Total_Amount_Alloted:,.2f}"
        else:
            amount_display = f"₹ {Total_Amount_Alloted:,.2f}"
        st.markdown(f"<h3 style='text-align: center; font-size: 25px;'>Reimbursement Amount (in Rupees)</h3>",
                    unsafe_allow_html=True,
                    )
        st.markdown(
            f"<div style='{CARD1_STYLE}'>"
            f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>{amount_display}</h2>"
            "</div>",
            unsafe_allow_html=True,
        )

    # Displaying count of transactions
    with colu3:
        Total_Transaction = len(filtered_df)
        st.markdown(
            f"<h3 style='text-align: center; font-size: 25px;'>Count Of Transactions</h3>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div style='{CARD2_STYLE}'>"
            f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
            "</div>",
            unsafe_allow_html=True,
        )

    # Converting date columns and rounding 'Amount'
    filtered_df["Document Date"] = pd.to_datetime(
        filtered_df["Document Date"], format="%Y/%m/%d", errors="coerce"
    ).dt.date
    filtered_df["Posting Date"] = pd.to_datetime(
        filtered_df["Posting Date"], format="%Y/%m/%d", errors="coerce"
    ).dt.date
    filtered_df["Verified on"] = pd.to_datetime(
        filtered_df["Verified on"], format="%Y/%m/%d", errors="coerce"
    ).dt.date
    filtered_df["Amount"] = filtered_df["Amount"].round()

    # Converting columns to strings and additional cleaning
    columns_to_convert = [
        "Payable req.no",
        "Doc.Type",
        "Invoice Number",
        "Text",
        "Cost Center",
        "G/L",
        "Document No",
        "Creator ID",
        "Verifier ID",
        "HOD Apr/Rej by",
        "HOG Approval by",
    ]
    filtered_df[columns_to_convert] = filtered_df[columns_to_convert].astype(str)
    filtered_df["Document No"] = filtered_df["Document No"].apply(
        lambda x: str(x) if isinstance(x, str) else ""
    )
    filtered_df["Document No"] = filtered_df["Document No"].apply(
        lambda x: re.sub(r"\..*", "", x)
    )
    filtered_df.sort_values(
        by=["Amount", "Cost Center", "ReimbursementID"],
        ascending=False,
        inplace=True,
    )
    filtered_df["HOD Apr/Rej on"] = pd.to_datetime(
        filtered_df["HOD Apr/Rej on"], errors="coerce"
    ).dt.date
    filtered_df.reset_index(drop=True, inplace=True)
    filtered_df.index += 1  # Start index from 1

    # Displaying filtered dataframe and creating downloadable Excel link
    st.dataframe(
        filtered_df[
            [
                "Payable req.no",
                "Doc.Type",
                "ReimbursementID",
                "Amount",
                "Document Date",
                "Name",
                "Invoice Number",
                "Text",
                "Cost Center",
                "G/L",
                "Document No",
                "Posting Date",
                "Creator ID",
                "Verifier ID",
                "Verified on",
                "HOD Apr/Rej on",
                "Reason for Rejection",
                "HOD Apr/Rej by",
                "HOG Approval by",
                "HOG Approval on",
            ]
        ]
    )
    excel_buffer = BytesIO()
    filtered_df.to_excel(excel_buffer, index=False)
    excel_buffer.seek(0)
    excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
    download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Remarks.xlsx">Download Excel file</a>'
    st.markdown(download_link, unsafe_allow_html=True)