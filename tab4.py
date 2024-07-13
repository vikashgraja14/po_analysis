from analyze_excelfinal import *
def checkbox(radio1,dfholiday,dfholiday2):
    t41, t42 = st.tabs(["General Parameters", "Authorization Parameters"])
    css = """
        <style>
        .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
            font-size: 1.0rem;  # Adjust the font size as needed
            font-weight: bold;  # Make the font bold
        }
        </style>
        """

    st.markdown(css, unsafe_allow_html=True)

    with t41:
        radio = radio1.copy()
        radio["Doc. Date"] = pd.to_datetime(
            radio["Doc. Date"], format="%Y/%m/%d", errors="coerce"
        )
        radio["Pstng Date"] = pd.to_datetime(
            radio["Pstng Date"], format="%Y/%m/%d", errors="coerce"
        )
        radio["Verified on"] = pd.to_datetime(
            radio["Verified on"], format="%Y/%m/%d", errors="coerce"
        )
        dfholiday["date"] = pd.to_datetime(
            dfholiday["date"], format="%Y/%m/%d", errors="coerce"
        )
        # Renaming columns in the DataFrame
        radio.rename(
            columns={
                "Vendor Name": "Name",
                "Pstng Date": "Posting Date",
                "Type": "Doc.Type",
                "Vendor": "Reimbursement ID",
                "Cost Ctr": "Cost Center",
                "Doc. Date": "Document Date",
                "category": "Category",
                "year": "YEAR",
                "Verified by": "Verifier ID",
                "Created": "Creator ID",
                "HOG Approval by": "HOG(Approval) ID",
                "HOD Apr/Rej by": "HOD(Approval) ID",
            },
            inplace=True,
        )
        radio = radio[
            [
                "Payable req.no",
                "Doc.Type",
                "Reimbursement ID",
                "Amount",
                "Document Date",
                "Name",
                "YEAR",
                "Invoice Number",
                "Text",
                "Cost Center",
                "G/L",
                "Document No",
                "Posting Date",
                "Creator ID",
                "Verifier ID",
                "HOD(Approval) ID",
                "Category",
                "Reference invoice",
                "CostctrName",
                "G/L Name",
                "Profit Ctr",
                "GR/IC Reference",
                "Org.unit",
                "Status",
                "File 1",
                "File 2",
                "File 3",
                "Time",
                "Updated at",
                "Reason for Rejection",
                "Verified at",
                "Reference document",
                "Adv.doc year",
                "HOG(Approval) ID",
                "Request no (Advance mulitple selection)",
                "Invoice Reference Number",
                "HOG Approval at",
                "HOG Approval Req",
                "Requested HOG ID",
                "Month",
                "Vesselcode",
                "PEA Number",
                "Status of Request",
                "Clearing doc no.",
                "On",
                "Updated on",
                "Verified on",
                "HOG Approval on",
                "Clearing date",
            ]
        ]
        radio["Document No"] = radio["Document No"].astype(str)
        radio["Document No"] = radio["Document No"].apply(
            lambda x: str(x) if isinstance(x, str) else ""
        )
        radio["Document No"] = radio["Document No"].apply(
            lambda x: re.sub(r"\..*", "", x)
        )
        c11, c2, c3, c4 = st.columns(4)

        options = ["All"] + [yr for yr in radio["YEAR"].unique() if yr != "All"]
        selected_option = c11.selectbox("Choose a YEAR", options, index=0)

        if selected_option != "All":
            radio = radio[radio["YEAR"] == selected_option]
        # Custom CSS to inject
        custom_css = """
                       <style>
                           .st-eb {
                               font-size: 1.0rem; /* Adjust font size */
                               border: 1px solid #ced4da; /* Add border */

                           }
                       </style>
                       """

        # Apply custom CSS
        st.markdown(custom_css, unsafe_allow_html=True)
        options = ["All"] + [yr for yr in radio["Category"].unique() if yr != "All"]
        selected_option = c2.multiselect(
            "Choose a Category", options, default=["All"]
        )
        # Display the multiselect widget

        if "All" in selected_option:
            radio = radio
        else:
            radio = radio[radio["Category"].isin(selected_option)]

        filtered_df = radio.copy()
        columns_to_convert = ["Name", "Invoice Number", "Reimbursement ID"]
        filtered_df[columns_to_convert] = filtered_df[columns_to_convert].astype(
            str
        )
        filtered_df["Document No"] = filtered_df["Document No"].astype(str)
        filtered_df["Document No"] = filtered_df["Document No"].apply(
            lambda x: str(x) if isinstance(x, str) else ""
        )
        filtered_df["Document No"] = filtered_df["Document No"].apply(
            lambda x: re.sub(r"\..*", "", x)
        )

        st.markdown(
            """
            <style>
            .stCheckbox > label { font-weight: bold !important; }
            </style>
            """,
            unsafe_allow_html=True,
        )
        if "session_state" not in st.session_state:
            st.session_state["session_state"] = False
        st.markdown(
            f"<h3 style='text-align: left; font-size: 20px;'>General Parameters</h3>",
            unsafe_allow_html=True,
        )
        col1, col2, col3, col4, col5 = st.columns(5)
        checkbox_statesGen = {
            "Reimbursement ID": col1.checkbox(
                "Reimbursement ID", key="Reimbursement ID"
            ),
            "Amount": col2.checkbox("Amount", key="Amount"),
            "Document Date": col3.checkbox("Document Date", key="Document Date"),
            "Cost Center": col4.checkbox("Cost Center", key="Cost Center"),
            "G/L": col5.checkbox("G/L", key="G/ L"),
            "Invoice Number": col1.checkbox("Invoice Number", key="Invoice Number"),
            "Text": col2.checkbox("Text", key="Text"),
        }
        checked_columnsGen = [
            key for key, value in checkbox_statesGen.items() if value
        ]
        st.markdown(
            f"<h3 style='text-align: left; font-size: 20px;'>Authorization Parameters</h3>",
            unsafe_allow_html=True,
        )
        col1, col2, col3, col4, col5 = st.columns(5)
        checkbox_statesAuth = {
            "Reimbursement ID": col1.checkbox(
                "Reimbursement ID", key="ReimbursementID"
            ),
            "Creator ID": col2.checkbox("Creator ID", key="Creator ID"),
            "Verifier ID": col3.checkbox("Verifier ID", key="Verifier ID"),
            "HOG(Approval) ID": col5.checkbox(
                "HOG(Approval) ID", key="HOG(Approval) ID"
            ),
            "HOD(Approval) ID": col4.checkbox(
                "HOD(Approval) ID", key="HOD(Approval) ID"
            ),
        }
        checked_columns_auth = [
            key for key, value in checkbox_statesAuth.items() if value
        ]
        st.markdown(
            f"<h3 style='text-align: left; font-size: 20px;'>Special Parameters</h3>",
            unsafe_allow_html=True,
        )
        col1, col2, col3, col4, col5 = st.columns(5)
        checkbox_statesSpec = {
            "80 % Same Invoice": col1.checkbox(
                "80 % Same Invoice", key="80 % Same Invoice"
            ),
            "Holiday Transactions": col2.checkbox(
                "Holiday Transactions", key="Holiday Transactions"
            )
        }
        checked_columnsspec = [
            key for key, value in checkbox_statesSpec.items() if value
        ]

        filename = "Exceptions.xlsx"

        def filter_dataframe(
                filtered_df,
                checked_columnsGen,
                checked_columns_auth,
                checked_columnsspec,
                dfholiday,
                filename,
        ):
            filtered_df["Invoice Number"] = filtered_df["Invoice Number"].astype(
                str
            )
            if (
                    not checked_columnsGen
                    and not checked_columns_auth
                    and not checked_columnsspec
            ):
                st.error("Please select at least one Checkbox .")
                return None, None, None, None, None, None
            else:
                if (
                        not checked_columnsGen
                        and not checked_columnsspec
                        and checked_columns_auth
                ):
                    filtered_df, checked_columns_auth, filename = filter_auth(
                        filtered_df, checked_columns_auth, filename
                    )
                elif (
                        not checked_columns_auth
                        and not checked_columnsspec
                        and checked_columnsGen
                ):
                    filtered_df, checked_columnsGen, filename = filter_gen(
                        filtered_df, checked_columnsGen, filename)

                    # filtered_df, checked_columnsGen, filename = filter_gen(
                    #     filtered_df, checked_columnsGen, filename
                    # )
                elif (
                        not checked_columns_auth
                        and not checked_columnsGen
                        and checked_columnsspec
                ):
                    filtered_df, checked_columnsspec, dfholiday, filename = filter_spec(filtered_df,
                                                                                        checked_columnsspec,
                                                                                        dfholiday, filename)

                elif (
                        not checked_columnsGen
                        and checked_columns_auth
                        and checked_columnsspec
                ):
                    filtered_df, checked_columns_auth, filename = filter_auth(
                        filtered_df, checked_columns_auth, filename
                    )

                    if filtered_df is not None:
                        filtered_df, checked_columnsspec, dfholiday, filename = (
                            filter_spec(
                                filtered_df,
                                checked_columnsspec,
                                dfholiday,
                                filename,
                            )
                        )


                    else:
                        return None, None, None, None, None, None

                elif (
                        not checked_columns_auth
                        and checked_columnsGen
                        and checked_columnsspec
                ):
                    filtered_df, checked_columnsGen, filename = filter_gen(
                        filtered_df, checked_columnsGen, filename
                    )

                    if filtered_df is not None:
                        for group_name, group_df in filtered_df.groupby(checked_columnsGen):
                            # Process each group (group_df) using group_name
                            filtered_df, checked_columnsspec, dfholiday, filename = filter_spec(group_df,
                                                                                                checked_columnsspec,
                                                                                                dfholiday, filename)

                    else:
                        return None, None, None, None, None, None

                elif (
                        not checked_columnsspec
                        and checked_columns_auth
                        and checked_columnsGen
                ):
                    filtered_df, checked_columns_auth, filename = filter_auth(
                        filtered_df, checked_columns_auth, filename
                    )
                    if filtered_df is not None:
                        filtered_df, checked_columnsGen, filename = filter_gen(
                            filtered_df, checked_columnsGen, filename
                        )
                    else:
                        return None, None, None, None, None, None
                else:
                    filtered_df, checked_columns_auth, filename = filter_auth(
                        filtered_df, checked_columns_auth, filename
                    )
                    if filtered_df is not None:
                        filtered_df, checked_columnsspec, dfholiday, filename = (
                            filter_spec(
                                filtered_df,
                                checked_columnsspec,
                                dfholiday,
                                filename,
                            )
                        )
                        if filtered_df is not None:
                            filtered_df, checked_columnsGen, filename = filter_gen(
                                filtered_df, checked_columnsGen, filename
                            )
                        else:
                            return None, None, None, None, None, None
                    else:
                        return None, None, None, None, None, None

            filename = f"Transactions_with_same_column.xlsx"
            try:
                return (
                    filtered_df,
                    checked_columnsGen,
                    checked_columns_auth,
                    checked_columnsspec,
                    dfholiday,
                    filename,
                )
            except Exception as e:
                st.error(f"An error occurred: {e}")
                return None, None, None, None, None, None

        filename = "Exceptions.xlsx"
        if filtered_df is not None:
            (
                filtered_df,
                checked_columnsGen,
                checked_columns_auth,
                checked_columnsspec,
                dfholiday,
                filename,
            ) = filter_dataframe(
                filtered_df,
                checked_columnsGen,
                checked_columns_auth,
                checked_columnsspec,
                dfholiday,
                filename,
            )
        if filtered_df is not None:
            if filtered_df.empty:
                st.markdown(
                    "<div style='text-align: center; font-weight: bold;'>No entries</div>",
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f"<h2 style='text-align: center; font-size: 35px; font-weight: bold;'>Filtered Entries</h2>",
                    unsafe_allow_html=True,
                )
                c111, card1, middle_column, card2, c222 = st.columns(
                    [1, 4, 1, 4, 1]
                )
                with card1:
                    Total_Amount_Alloted = filtered_df["Amount"].sum()

                    # Convert the total amount to a string for length checking
                    total_amount_str = str(Total_Amount_Alloted)

                    # Check if the length of Total_Amount_Alloted is greater than 5
                    if len(total_amount_str) > 5:
                        # Get the integer part of the total amount
                        integer_part = int(float(total_amount_str))
                        # Calculate the length of the integer part
                        integer_length = len(str(integer_part))

                        # Divide by 1 lakh if the integer length is greater than 5 and less than or equal to 7
                        if integer_length > 5 and integer_length <= 7:
                            Total_Amount_Alloted /= 100000
                            amount_display = f"₹ {Total_Amount_Alloted:,.2f} lakhs"
                        # Divide by 1 crore if the integer length is greater than 7
                        elif integer_length > 7:
                            Total_Amount_Alloted /= 10000000
                            amount_display = f"₹ {Total_Amount_Alloted:,.2f} crores"
                        else:
                            amount_display = f"₹ {Total_Amount_Alloted:,.2f}"
                    else:
                        amount_display = f"₹ {Total_Amount_Alloted:,.2f}"

                    # Display the total amount spent
                    st.markdown(
                        f"<h3 style='text-align: center; font-size: 25px;'>Reimbursement Amount(in Rupees)</h3>",
                        unsafe_allow_html=True,
                    )
                    st.markdown(
                        f"<div style='{CARD1_STYLE}'>"
                        f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>{amount_display}</h2>"
                        "</div>",
                        unsafe_allow_html=True,
                    )

                with card2:
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
                filtered_df["Document Date"] = pd.to_datetime(
                    filtered_df["Document Date"], format="%Y/%m/%d", errors="coerce"
                )
                filtered_df["Posting Date"] = pd.to_datetime(
                    filtered_df["Posting Date"], format="%Y/%m/%d", errors="coerce"
                )
                filtered_df["Verified on"] = pd.to_datetime(
                    filtered_df["Verified on"], format="%Y/%m/%d", errors="coerce"
                )
                filtered_df["Posting Date"] = filtered_df["Posting Date"].dt.date
                filtered_df["Document Date"] = filtered_df["Document Date"].dt.date
                filtered_df["Verified on"] = filtered_df["Verified on"].dt.date
                # Ensure that 'filtered_df1' is a copy of 'filtered_df' with rounded 'Amount'
                filtered_df1 = filtered_df.copy()
                filtered_df1["Amount"] = filtered_df1["Amount"].round()

                filtered_df1.reset_index(drop=True, inplace=True)
                filtered_df1.index += 1  # Start index from 1
                st.dataframe(
                    filtered_df1[
                        [
                            "Payable req.no",
                            "Doc.Type",
                            "Reimbursement ID",
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
                            "HOD(Approval) ID",
                            "HOG(Approval) ID",
                        ]
                    ]
                )
                # Generate and provide a download link for the Excel file
                filename = f" Filtered entries.xlsx"
                filtered_df.reset_index(drop=True, inplace=True)
                filtered_df.index += 1  # Start index from 1
                excel_buffer = BytesIO()
                filtered_df.to_excel(excel_buffer, index=False)
                # Reset the buffer's position to the start for reading
                excel_buffer.seek(0)
                # Convert Excel buffer to base64
                excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()

                download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="{filename}">Download Excel file</a>'
                st.markdown(download_link, unsafe_allow_html=True)


    with t42:
        radio2 = radio1.copy()
        radio2["Doc. Date"] = pd.to_datetime(
            radio2["Doc. Date"], format="%Y/%m/%d", errors="coerce"
        )
        radio2["Pstng Date"] = pd.to_datetime(
            radio2["Pstng Date"], format="%Y/%m/%d", errors="coerce"
        )
        radio2["Verified on"] = pd.to_datetime(
            radio2["Verified on"], format="%Y/%m/%d", errors="coerce"
        )
        dfholiday2["date"] = pd.to_datetime(
            dfholiday2["date"], format="%Y/%m/%d", errors="coerce"
        )
        # Renaming columns in the DataFrame
        radio2.rename(
            columns={
                "Vendor Name": "Name",
                "Pstng Date": "Posting Date",
                "Type": "Doc.Type",
                "Vendor": "ReimbursementID",
                "Cost Ctr": "Cost Center",
                "Doc. Date": "Document Date",
                "Verified by": "Verifier ID",
                "Created": "Creator ID",
                "year": "Year",
                "HOG Approval by": "HOG(Approval) ID",
                "HOD Apr/Rej by": "HOD(Approval) ID",
            },
            inplace=True,
        )
        radio2 = radio2[
            [
                "Payable req.no",
                "Doc.Type",
                "ReimbursementID",
                "Amount",
                "Document Date",
                "Name",
                "Year",
                "Invoice Number",
                "Text",
                "Cost Center",
                "G/L",
                "Document No",
                "Posting Date",
                "Creator ID",
                "Verifier ID",
                "HOG(Approval) ID",
                "HOD(Approval) ID",
                "HOD Apr/Rej on",
                "category",
                "Reference invoice",
                "CostctrName",
                "G/L Name",
                "Profit Ctr",
                "GR/IC Reference",
                "Org.unit",
                "Status",
                "File 1",
                "File 2",
                "File 3",
                "Time",
                "Updated at",
                "Reason for Rejection",
                "Verified at",
                "Reference document",
                "Adv.doc year",
                "Request no (Advance mulitple selection)",
                "Invoice Reference Number",
                "HOG Approval at",
                "HOG Approval Req",
                "Requested HOG ID",
                "Month",
                "Vesselcode",
                "PEA Number",
                "Status of Request",
                "Clearing doc no.",
                "On",
                "Updated on",
                "Verified on",
                "HOG Approval on",
                "Clearing date",
            ]
        ]
        radio2["Document No"] = radio2["Document No"].astype(str)
        radio2["Document No"] = radio2["Document No"].apply(
            lambda x: str(x) if isinstance(x, str) else ""
        )
        radio2["Document No"] = radio2["Document No"].apply(
            lambda x: re.sub(r"\..*", "", x)
        )
        c11, c2, c3, c4 = st.columns(4)
        options = ["All"] + [
            yr for yr in radio2["Year"].unique() if yr != "All"
        ]
        selected_option = c11.selectbox("Choose an year", options, index=0)

        if selected_option != "All":
            radio2 = radio2[radio2["Year"] == selected_option]

        options = ["All"] + [
            yr for yr in radio2["category"].unique() if yr != "All"
        ]
        selected_option = c2.multiselect(
            "select a category", options, default=["All"]
        )

        if "All" in selected_option:
            radio2 = radio2
        else:
            radio2 = radio2[radio2["category"].isin(selected_option)]

        filtered_df = radio2.copy()
        filtered_df["Document No"] = filtered_df["Document No"].astype(str)
        filtered_df["Document No"] = filtered_df["Document No"].apply(
            lambda x: str(x) if isinstance(x, str) else ""
        )
        filtered_df["Document No"] = filtered_df["Document No"].apply(
            lambda x: re.sub(r"\..*", "", x)
        )
        filtered_df["Document Date"] = pd.to_datetime(
            filtered_df["Document Date"], format="%Y/%m/%d", errors="coerce"
        )
        filtered_df["Posting Date"] = pd.to_datetime(
            filtered_df["Posting Date"], format="%Y/%m/%d", errors="coerce"
        )
        filtered_df["Verified on"] = pd.to_datetime(
            filtered_df["Verified on"], format="%Y/%m/%d", errors="coerce"
        )
        dfholiday2["date"] = pd.to_datetime(
            dfholiday2["date"], format="%Y/%m/%d", errors="coerce"
        )

        if "sesion_state" not in st.session_state:
            st.session_state["sesion_state"] = False
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        checkbox_states2 = {
            "ReimbursementID": col1.checkbox(
                "ReimbursementID", key="Reimbursement  ID"
            ),
            "Creator ID": col2.checkbox("Creator ID", key="CreatorID"),
            "Verifier ID": col3.checkbox("Verifier ID", key="VerifierID"),
            "HOG(Approval) ID": col5.checkbox(
                "HOG(Approval) ID", key="HOG(ApprovalID"
            ),
            "HOD(Approval) ID": col4.checkbox(
                "HOD(Approval) ID", key="HOD(ApprovalID"
            ),
            "HolidayTransactions": col6.checkbox(
                "HolidayTransactions", key="Holidaytransactions"
            ),
        }
        css = """
                     <style>
                     [data-baseweb="checkbox"] [data-testid="stWidgetLabel"] p {
                         /* Styles for the label text for checkbox and toggle */
                         font-size: 1.0rem;
                         width: 300px;
                         margin-top: 0.01rem;
                     }

                     [data-baseweb="checkbox"] div {
                         /* Styles for the slider container */
                         height: 1rem;
                         width: 1.0rem;
                     }
                     [data-baseweb="checkbox"] div div {
                         /* Styles for the slider circle */
                         height: 1.8rem;
                         width: 1.8rem;
                     }
                     [data-testid="stCheckbox"] label span {
                         /* Styles the checkbox */
                         height: 1rem;
                         width: 1rem;
                     }
                     </style>
                     """

        st.markdown(css, unsafe_allow_html=True)

        checked_columns2 = [
            key for key, value in checkbox_states2.items() if value
        ]
        columns_to_check_for_duplicates2 = [
            column
            for column in checked_columns2
            if column not in ["HolidayTransactions"]
        ]

        # @st.experimental_fragment
        def filter_dataframe2(
                filtered_df,
                checked_columns2,
                columns_to_check_for_duplicates2,
                dfholiday2,
                filename,
        ):
            if not checked_columns2:
                st.error(
                    "Please select at least one column to check for duplicates."
                )
                return None, None, None, None, None
            else:
                if "HolidayTransactions" in checked_columns2:
                    if len(checked_columns2) == 1:
                        # Filter by holiday transactions
                        filtered_df["Document Date"] = pd.to_datetime(
                            filtered_df["Document Date"],
                            format="%Y/%m/%d",
                            errors="coerce",
                        )
                        filtered_df["Posting Date"] = pd.to_datetime(
                            filtered_df["Posting Date"],
                            format="%Y/%m/%d",
                            errors="coerce",
                        )
                        filtered_df["Verified on"] = pd.to_datetime(
                            filtered_df["Verified on"],
                            format="%Y/%m/%d",
                            errors="coerce",
                        )
                        dfholiday2["date"] = pd.to_datetime(
                            dfholiday2["date"],
                            format="%Y/%m/%d",
                            errors="coerce",
                        )
                        filtered_df = filtered_df[
                            filtered_df["Posting Date"].isin(dfholiday2["date"])
                        ]
                        filename = "HolidayTransactions.xlsx"
                    elif len(checked_columns2) == 2:
                        st.error("Please select another checkbox.")
                        return None, None, None, None, None
                    else:
                        # Filter by holiday transactions
                        filtered_df["Document Date"] = pd.to_datetime(
                            filtered_df["Document Date"],
                            format="%Y/%m/%d",
                            errors="coerce",
                        )
                        filtered_df["Posting Date"] = pd.to_datetime(
                            filtered_df["Posting Date"],
                            format="%Y/%m/%d",
                            errors="coerce",
                        )
                        filtered_df["Verified on"] = pd.to_datetime(
                            filtered_df["Verified on"],
                            format="%Y/%m/%d",
                            errors="coerce",
                        )
                        dfholiday2["date"] = pd.to_datetime(
                            dfholiday2["date"],
                            format="%Y/%m/%d",
                            errors="coerce",
                        )
                        filtered_df = filtered_df[
                            filtered_df["Posting Date"].isin(dfholiday2["date"])
                        ]
                        # Assuming 'columns_to_check_for_duplicates2' is a list of column names to check
                        for i in range(len(columns_to_check_for_duplicates2)):
                            for j in range(
                                    i + 1, len(columns_to_check_for_duplicates2)
                            ):
                                # Compare each column with every other column
                                col_i = columns_to_check_for_duplicates2[i]
                                col_j = columns_to_check_for_duplicates2[j]
                                filtered_df = filtered_df[
                                    filtered_df[col_i] == filtered_df[col_j]
                                    ]
                else:
                    if len(checked_columns2) == 1:
                        st.error("Please select another checkbox.")
                        return None, None, None, None, None
                    else:
                        for i in range(len(columns_to_check_for_duplicates2)):
                            for j in range(
                                    i + 1, len(columns_to_check_for_duplicates2)
                            ):
                                # Compare each column with every other column
                                col_i = columns_to_check_for_duplicates2[i]
                                col_j = columns_to_check_for_duplicates2[j]
                                filtered_df = filtered_df[
                                    filtered_df[col_i] == filtered_df[col_j]
                                    ]
                checked_columns2 = [
                    "Posting Date" if col == "HolidayTransactions" else col
                    for col in checked_columns2
                ]
                sort_columns2 = checked_columns2.copy()

                if (
                        "ReimbursementID" not in checked_columns2
                        and "Cost Center" not in checked_columns2
                ):
                    sort_columns2 += ["ReimbursementID", "Cost Center"]
                elif (
                        "ReimbursementID" in checked_columns2
                        and "Cost Center" not in checked_columns2
                ):
                    sort_columns2 += ["Cost Center"]
                elif (
                        "Cost Center" in checked_columns2
                        and "ReimbursementID" not in checked_columns2
                ):
                    sort_columns2 += ["ReimbursementID"]

                filtered_df = filtered_df.sort_values(by=sort_columns2)
                filtered_df.reset_index(drop=True, inplace=True)
                filtered_df.index += 1
                filename = f"Transactions_with_same_column.xlsx"

                try:
                    return (
                        filtered_df,
                        checked_columns2,
                        columns_to_check_for_duplicates2,
                        dfholiday2,
                        filename,
                    )
                except Exception as e:
                    st.error(f"An error occurred: {e}")
                    return None, None, None, None, None

        s1, s2, s3, s4, s5, s6, s7 = st.columns(7)
        if s1.button("ANALYZE"):
            st.session_state["sesion_state"] = True
            filename = "Exceptions.xlsx"
            (
                filtered_df,
                checked_columns2,
                columns_to_check_for_duplicates2,
                dfholiday2,
                filename,
            ) = filter_dataframe2(
                filtered_df,
                checked_columns2,
                columns_to_check_for_duplicates2,
                dfholiday2,
                filename,
            )
            # Check if 'filtered_df' is not None before proceeding
            if filtered_df is not None:
                if filtered_df.empty:
                    st.markdown(
                        "<div style='text-align: center; font-weight: bold;'>No entries</div>",
                        unsafe_allow_html=True,
                    )
                else:
                    c111, card1, middle_column, card2, c222 = st.columns(
                        [1, 4, 1, 4, 1]
                    )
                    with card1:
                        Total_Amount_Alloted = filtered_df["Amount"].sum()

                        # Convert the total amount to a string for length checking
                        total_amount_str = str(Total_Amount_Alloted)

                        # Check if the length of Total_Amount_Alloted is greater than 5
                        if len(total_amount_str) > 5:
                            # Get the integer part of the total amount
                            integer_part = int(float(total_amount_str))
                            # Calculate the length of the integer part
                            integer_length = len(str(integer_part))

                            # Divide by 1 lakh if the integer length is greater than 5 and less than or equal to 7
                            if integer_length > 5 and integer_length <= 7:
                                Total_Amount_Alloted /= 100000
                                amount_display = (
                                    f"₹ {Total_Amount_Alloted:,.2f} lakhs"
                                )
                            # Divide by 1 crore if the integer length is greater than 7
                            elif integer_length > 7:
                                Total_Amount_Alloted /= 10000000
                                amount_display = (
                                    f"₹ {Total_Amount_Alloted:,.2f} crores"
                                )
                            else:
                                amount_display = (
                                    f"₹ {Total_Amount_Alloted:,.2f}"
                                )
                        else:
                            amount_display = f"₹ {Total_Amount_Alloted:,.2f}"

                        # Display the total amount spent
                        st.markdown(
                            f"<h3 style='text-align: center; font-size: 25px;'>Reimbursement Amount(in Rupees)</h3>",
                            unsafe_allow_html=True,
                        )
                        st.markdown(
                            f"<div style='{CARD1_STYLE}'>"
                            f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>{amount_display}</h2>"
                            "</div>",
                            unsafe_allow_html=True,
                        )

                    with card2:
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
                    filtered_df["Document Date"] = pd.to_datetime(
                        filtered_df["Document Date"],
                        format="%Y/%m/%d",
                        errors="coerce",
                    )
                    filtered_df["Posting Date"] = pd.to_datetime(
                        filtered_df["Posting Date"],
                        format="%Y/%m/%d",
                        errors="coerce",
                    )
                    filtered_df["Verified on"] = pd.to_datetime(
                        filtered_df["Verified on"],
                        format="%Y/%m/%d",
                        errors="coerce",
                    )
                    filtered_df["Posting Date"] = filtered_df[
                        "Posting Date"
                    ].dt.date
                    filtered_df["Document Date"] = filtered_df[
                        "Document Date"
                    ].dt.date
                    filtered_df["Verified on"] = filtered_df[
                        "Verified on"
                    ].dt.date
                    # Ensure that 'filtered_df1' is a copy of 'filtered_df' with rounded 'Amount'
                    filtered_df1 = filtered_df.copy()
                    filtered_df1["Amount"] = (
                        filtered_df1["Amount"].round().astype(int)
                    )

                    filtered_df1.reset_index(drop=True, inplace=True)
                    filtered_df1.index += 1  # Start index from 1
                    filtered_df1["HOG(Approval) ID"] = filtered_df1[
                        "HOG(Approval) ID"
                    ].astype(str)
                    filtered_df1["HOD(Approval) ID"] = filtered_df1[
                        "HOD(Approval) ID"
                    ].astype(str)
                    filtered_df1["HOD(Approval) ID"] = filtered_df1[
                        "HOD(Approval) ID"
                    ].astype(str)
                    filtered_df1["HOD(Approval) ID"] = filtered_df1[
                        "HOD(Approval) ID"
                    ].astype(str)
                    filtered_df1["HOG(Approval) ID"] = filtered_df1[
                        "HOG(Approval) ID"
                    ].apply(lambda x: str(x) if isinstance(x, str) else "")
                    filtered_df1["HOG(Approval) ID"] = filtered_df1[
                        "HOG(Approval) ID"
                    ].apply(lambda x: re.sub(r"\..*", "", x))
                    st.dataframe(
                        filtered_df1[
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
                                "HOD(Approval) ID",
                                "HOG(Approval) ID",
                            ]
                        ].style.set_properties(**{"font-size": "16px"})
                    )

                    filename = f"SAME {checked_columns2}.xlsx"
                    filtered_df.reset_index(drop=True, inplace=True)
                    filtered_df.index += 1  # Start index from 1
                    excel_buffer = BytesIO()
                    filtered_df.to_excel(excel_buffer, index=False)
                    # Reset the buffer's position to the start for reading
                    excel_buffer.seek(0)
                    # Convert Excel buffer to base64
                    excel_b64 = base64.b64encode(
                        excel_buffer.getvalue()
                    ).decode()
                    download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="{filename}">Download Excel file</a>'
                    st.markdown(download_link, unsafe_allow_html=True)



