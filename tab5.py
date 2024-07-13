from analyze_excelfinal import *
import base64

def tab5(ApprovalExceptions,exceptions2,concatenated_df):
    t51, t52 = st.tabs(["Exceptions", "Unkown Verifier IDs"])

    # @st.experimental_fragment
    with t51:
        t511, t522 = st.tabs(
            ["HOD Approval Exceptions", "HOG Approval Exceptions"]
        )

        with t511:
            # ApprovalExceptions1 = ApprovalExceptions.copy()
            if concatenated_df is not None:
                ApprovalExceptions1 = ApprovalExceptions.copy()

                c11, c2, c3, c4 = st.columns(4)

                option = ["All"] + [yr for yr in ApprovalExceptions1["year"].unique() if yr != "All"]
                selected_option = c11.selectbox("SELECT year", option, index=0)

                if selected_option != "All":
                    ApprovalExceptions1 = ApprovalExceptions1[ApprovalExceptions1["year"] == selected_option]
                ApprovalExceptions = ApprovalExceptions1.copy()
                concatenated_df["Personnel No."] = concatenated_df[
                    "Personnel No."
                ].astype(str)
                concatenated_df["Personnel No."] = concatenated_df[
                    "Personnel No."
                ].apply(lambda x: str(x) if isinstance(x, str) else "")
                concatenated_df["Personnel No."] = concatenated_df[
                    "Personnel No."
                ].apply(lambda x: re.sub(r"\..*", "", x))
                concatenated_df["Date"] = pd.to_datetime(
                    concatenated_df["Date"], errors="coerce"
                )
                concatenated_df["Date"] = concatenated_df["Date"].dt.date
                concatenated = ApprovalExceptions[
                    ApprovalExceptions[["HOD Apr/Rej on", "HOD Apr/Rej by"]]
                    .apply(tuple, axis=1)
                    .isin(
                        concatenated_df[["Date", "Personnel No."]].apply(
                            tuple, axis=1
                        )
                    )
                ]

                if concatenated is not None:
                    if concatenated.empty:
                        st.markdown(
                            "<div style='text-align: center; font-weight: bold;'>No such entries</div>",
                            unsafe_allow_html=True,
                        )
                    else:
                        concatenated.rename(
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
                        concatenated[
                            ["HOD Apr/Rej on", "HOG Approval by", "year"]
                        ] = concatenated[
                            ["HOD Apr/Rej on", "HOG Approval by", "year"]
                        ].astype(
                            str
                        )

                        concatenated.reset_index(drop=True, inplace=True)
                        concatenated.index += 1  # Start index from 1
                        namehoD = concatenated_df.copy()
                        namehoD = namehoD.rename(
                            columns={
                                "Personnel No.": "HOD Apr/Rej by",
                                "Empl./appl.name": "HOD NAME",
                                "Name": "Department",
                            }
                        )

                        # Create mappings
                        hod_name_mapping = dict(
                            zip(namehoD["HOD Apr/Rej by"], namehoD["HOD NAME"])
                        )
                        dept_mapping = dict(
                            zip(
                                namehoD["HOD Apr/Rej by"], namehoD["Department"]
                            )
                        )

                        # Apply mappings to the concatenated DataFrame
                        concatenated["HOD NAME"] = concatenated[
                            "HOD Apr/Rej by"
                        ].map(hod_name_mapping)
                        concatenated["Department"] = concatenated[
                            "HOD Apr/Rej by"
                        ].map(dept_mapping)

                        namehoG = concatenated_df.copy()
                        namehoG = namehoG.rename(
                            columns={
                                "Personnel No.": "HOG Approval by",
                                "Empl./appl.name": "HOG NAME",
                                "Date": "HOG Approval on",
                            }
                        )
                        mapping = dict(
                            zip(namehoG["HOG Approval by"], namehoG["HOG NAME"])
                        )
                        concatenated["HOG NAME"] = concatenated[
                            "HOG Approval by"
                        ].map(mapping)

                        columns_to_convert = [
                            "Payable req.no",
                            "Doc.Type",
                            "HOD Apr/Rej on",
                            "HOG Approval by",
                            "year",
                            "HOG NAME",
                            "Invoice Number",
                            "Text",
                            "Cost Center",
                            "G/L",
                            "Document No",
                            "Creator ID",
                            "Verifier ID",
                        ]
                        concatenated[columns_to_convert] = concatenated[
                            columns_to_convert
                        ].astype(str)
                        concatenated = concatenated[
                            [
                                "Payable req.no",
                                "Invoice Number",
                                "Doc.Type",
                                "Name",
                                "ReimbursementID",
                                "category",
                                "Profit Ctr",
                                "BP",
                                "CostctrName",
                                "Cost Center",
                                "G/L",
                                "G/L Name",
                                "GR/IC Reference",
                                "Text",
                                "Org.unit",
                                "Status",
                                "File 1",
                                "File 2",
                                "File 3",
                                "Document No",
                                "Creator ID",
                                "Time",
                                "Updated at",
                                "HOD NAME",
                                "HOD Apr/Rej by",
                                "Department",
                                "HOD Apr/Rej at",
                                "Reason for Rejection",
                                "Verifier ID",
                                "Verified at",
                                "Reference document",
                                "Reference invoice",
                                "Adv.doc year",
                                "Request no (Advance mulitple selection)",
                                "ID",
                                "Invoice Reference Number",
                                "HOG NAME",
                                "HOG Approval by",
                                "HOG Approval at",
                                "HOG Approval Req",
                                "Requested HOG ID",
                                "Month",
                                "Vesselcode",
                                "PEA Number",
                                "Status of Request",
                                "Clearing doc no.",
                                "Document Date",
                                "Posting Date",
                                "Amount",
                                "On",
                                "Updated on",
                                "HOD Apr/Rej on",
                                "Verified on",
                                "HOG Approval on",
                                "Clearing date",
                                "year",
                            ]
                        ]



                        ccc, card01, c111, card1, middle_column, card2, c222 = (
                            st.columns([1, 2, 1, 2, 1, 2, 1])
                        )

                        with card01:
                            total_HODs = concatenated[
                                "HOD Apr/Rej by"
                            ].nunique()
                            total_HODs_str = str(total_HODs)
                            st.markdown(
                                f"<h3 style='text-align: center; font-size: 25px;'>NO Of HODs</h3>",
                                unsafe_allow_html=True,
                            )
                            st.write("")
                            st.markdown(
                                f"<div style='{CARD1_STYLE}'>"
                                f"<h2 style='color: #28a745; text-align: center; font-size: 35px;'>{total_HODs_str}</h2>"
                                "</div>",
                                unsafe_allow_html=True,
                            )

                        with card1:
                            Total_Amount_Alloted = concatenated["Amount"].sum()

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
                                amount_display = (
                                    f"₹ {Total_Amount_Alloted:,.2f}"
                                )

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
                            Total_Transaction = len(concatenated)
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
                        concatenated["Amount"] = concatenated["Amount"].round()
                        concatenated.reset_index(drop=True, inplace=True)
                        concatenated.index += 1  # Start index from 1
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
                            "HOG Approval by",
                        ]
                        concatenated[columns_to_convert] = concatenated[
                            columns_to_convert
                        ].astype(str)
                        concatenated["Document No"] = concatenated[
                            "Document No"
                        ].apply(lambda x: str(x) if isinstance(x, str) else "")
                        concatenated["Document No"] = concatenated[
                            "Document No"
                        ].apply(lambda x: re.sub(r"\..*", "", x))
                        concatenated["HOG Approval by"] = concatenated[
                            "HOG Approval by"
                        ].astype(str)
                        concatenated["HOG Approval by"] = concatenated[
                            "HOG Approval by"
                        ].apply(lambda x: str(x) if isinstance(x, str) else "")
                        concatenated["HOG Approval by"] = concatenated[
                            "HOG Approval by"
                        ].apply(lambda x: re.sub(r"\..*", "", x))

                        concatenated_show = concatenated.copy()
                        grouped_df = concatenated_show.groupby("HOD Apr/Rej by")
                        # Create 'Total transactions' column
                        concatenated_show["value"] = grouped_df[
                            "Amount"
                        ].transform("sum")
                        concatenated_show["Count of transactions"] = grouped_df[
                            "HOD Apr/Rej by"
                        ].transform(len)
                        # Create 'No of Days' column with unique values of 'HOG Approval on'
                        concatenated_show["No of Days"] = grouped_df[
                            "HOD Apr/Rej on"
                        ].transform("nunique")
                        concatenated_show = concatenated_show.rename(
                            columns={"HOD Apr/Rej by": "HOD ID"}
                        )
                        concatenated_show.sort_values(
                            by="value", ascending=False, inplace=True
                        )

                        concatenated_show = concatenated_show.drop_duplicates(
                            subset="HOD ID", keep="last"
                        )

                        cdf1, cdf2, cd3 = st.columns([2, 6, 2])
                        concatenated_show.reset_index(drop=True, inplace=True)
                        concatenated_show.index += 1  # Start index from 1
                        cdf2.dataframe(
                            concatenated_show[
                                [
                                    "HOD ID",
                                    "HOD NAME",
                                    "Department",
                                    "value",
                                    "Count of transactions",
                                    "No of Days",
                                ]
                            ]
                        )
                        space1, cdf1, cdf2, space2 = st.columns([1, 4, 4, 1])
                        grouped_data = (
                            concatenated_show.groupby("Department")["value"]
                            .sum()
                            .reset_index()
                        )

                        # Sort by value in descending order
                        grouped_data.sort_values(
                            by="value", ascending=False, inplace=True
                        )

                        # Select the top 10 departments
                        top_10 = grouped_data.head(10)

                        # Calculate the total value of the remaining departments
                        other_value = (
                                grouped_data["value"].sum() - top_10["value"].sum()
                        )

                        # Create a new DataFrame for the pie chart
                        pie_data = pd.concat(
                            [
                                top_10,
                                pd.DataFrame(
                                    {
                                        "Department": ["Others"],
                                        "value": [other_value],
                                    }
                                ),
                            ]
                        )
                        #
                        # # Create the pie chart
                        # fig = px.pie(pie_data, values='value', names='Department',
                        #              title='Top 10 Departments (Value Wise)')

                        # Create Pie chart for 'value' column
                        fig = px.pie(
                            pie_data,
                            values="value",
                            names="Department",
                            title="Value Wise",
                        )
                        fig.update_layout(width=400, height=400)
                        fig.update_traces(textinfo="percent")
                        fig.update(layout_title_text="")
                        cdf1.write("")
                        cdf1.markdown(
                            f"<h3 style='text-align: left; font-size: 20px;'>Value Wise</h3>",
                            unsafe_allow_html=True,
                        )
                        cdf1.plotly_chart(fig)

                        # Create Pie chart for 'No of Days' column
                        grouped_data2 = (
                            concatenated_show.groupby("Department")[
                                "No of Days"
                            ]
                            .sum()
                            .reset_index()
                        )

                        # Sort by sum of days in descending order
                        grouped_data2.sort_values(
                            by="No of Days", ascending=False, inplace=True
                        )

                        # Select the top 10 departments
                        top_10 = grouped_data2.head(10)

                        # Calculate the total sum of days for the remaining departments
                        other_sum1 = (
                                grouped_data2["No of Days"].sum()
                                - top_10["No of Days"].sum()
                        )

                        # Create a new DataFrame for the pie chart
                        pie_data1 = pd.concat(
                            [
                                top_10,
                                pd.DataFrame(
                                    {
                                        "Department": ["Others"],
                                        "No of Days": [other_sum1],
                                    }
                                ),
                            ]
                        )
                        fig2 = px.pie(
                            pie_data1,
                            values="No of Days",
                            names="Department",
                            title="Day Wise",
                        )
                        fig2.update_layout(width=400, height=400)
                        fig2.update_traces(textinfo="percent")
                        fig2.update(layout_title_text="")
                        cdf2.write("")
                        cdf2.markdown(
                            f"<h3 style='text-align: left; font-size: 20px;'>Day Wise</h3>",
                            unsafe_allow_html=True,
                        )
                        cdf2.plotly_chart(fig2)

                        concatenated.sort_values(
                            by=["HOD Apr/Rej by"], inplace=True
                        )
                        concatenated.reset_index(drop=True, inplace=True)
                        concatenated.index += 1  # Start index from 1
                        concatenated_show = concatenated_show[
                            [
                                "HOD ID",
                                "HOD NAME",
                                "Department",
                                "value",
                                "Count of transactions",
                                "No of Days",
                            ]
                        ]

                        excel_buffer = BytesIO()
                        concatenated.to_excel(excel_buffer, index=False)
                        with pd.ExcelWriter(
                                excel_buffer, engine="xlsxwriter"
                        ) as writer:
                            concatenated_show.to_excel(
                                writer, index=False, sheet_name="Summary"
                            )
                            concatenated.to_excel(
                                writer, index=False, sheet_name="Raw Data"
                            )
                        # Reset the buffer's position to the start for reading
                        excel_buffer.seek(0)
                        # Convert Excel buffer to base64
                        excel_b64 = base64.b64encode(
                            excel_buffer.getvalue()
                        ).decode()
                        # Download link for Excel file within a Markdown
                        download_link = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}" download="Approval Exceptions.xlsx">Download Excel file</a>'
                        st.markdown(download_link, unsafe_allow_html=True)
                        # Reset the buffer's position to the start for reading

    with t522:
        if concatenated_df is not None:
            concatenated_HOG = concatenated_df.copy()
            ApprovalExceptions_hog = ApprovalExceptions.copy()
            colu1, colu2, _, colu3, _ = st.columns([1, 4, 1, 4, 1])

            choice = ["All"] + [yr for yr in ApprovalExceptions_hog["year"].unique() if yr != "All"]
            choose_options = colu1.selectbox("select an option", choice, index=0)

            if choose_options != "All":
                ApprovalExceptions_hog = ApprovalExceptions_hog[ApprovalExceptions_hog["year"] == choose_options]
            else:
                ApprovalExceptions_hog = ApprovalExceptions_hog.copy()

            concatenated_HOG['Personnel No.'] = concatenated_HOG['Personnel No.'].astype(str)
            concatenated_HOG['Personnel No.'] = concatenated_HOG['Personnel No.'].apply(
                lambda x: str(x) if isinstance(x, str) else '')
            concatenated_HOG['Personnel No.'] = concatenated_HOG['Personnel No.'].apply(
                lambda x: re.sub(r'\..*', '', x))
            concatenated_HOG['Date'] = pd.to_datetime(concatenated_HOG['Date'], errors='coerce')
            concatenated_HOG['Date'] = concatenated_HOG['Date'].dt.date
            concatenated_hog = ApprovalExceptions_hog[

                ApprovalExceptions_hog[['HOG Approval on', 'HOG Approval by']].apply(tuple,
                                                                                     axis=1).isin(
                    concatenated_HOG[['Date', 'Personnel No.']].apply(tuple, axis=1))]

            if concatenated_hog is not None:
                if concatenated_hog.empty:
                    st.markdown(
                        "<div style='text-align: center; font-weight: bold;'>No such entries</div>",
                        unsafe_allow_html=True)
                else:
                    # concatenated['HOG Approval by'] = concatenated['HOG Approval by'].astype(str)
                    concatenated_hog.rename(
                        columns={
                            'Vendor Name': 'Name',
                            'Pstng Date': 'Posting Date',
                            'Type': "Doc.Type",
                            'Vendor': "ReimbursementID",
                            "Cost Ctr": "Cost Center",
                            'Doc. Date': 'Document Date',
                            "Verified by": 'Verifier ID',
                            'Created': 'Creator ID',
                        },
                        inplace=True
                    )
                    namehoD = concatenated_HOG.copy()
                    namehoD = namehoD.rename(
                        columns={'Personnel No.': 'HOD Apr/Rej by', 'Empl./appl.name': 'HOD NAME'
                                 })
                    mapping = dict(zip(namehoD['HOD Apr/Rej by'], namehoD['HOD NAME']))
                    concatenated_hog['HOD NAME'] = concatenated_hog['HOD Apr/Rej by'].map(mapping)
                    namehoG = concatenated_HOG.copy()
                    namehoG = namehoG.rename(
                        columns={'Personnel No.': 'HOG Approval by', 'Empl./appl.name': 'HOG NAME',
                                 'Name': 'Department',
                                 'Date': 'HOG Approval on'})

                    # Create mappings
                    hog_name_mapping = dict(zip(namehoG['HOG Approval by'], namehoG['HOG NAME']))
                    deptg_mapping = dict(zip(namehoG['HOG Approval by'], namehoG['Department']))

                    # Apply mappings to the concatenated DataFrame
                    concatenated_hog['HOG NAME'] = concatenated_hog['HOG Approval by'].map(hog_name_mapping)
                    concatenated_hog['Department'] = concatenated_hog['HOG Approval by'].map(deptg_mapping)

                    concatenated_hog[['HOD Apr/Rej on', 'HOG Approval by', 'year']] = concatenated_hog[
                        ['HOD Apr/Rej on', 'HOG Approval by', 'year']].astype(str)
                    concatenated_hog = concatenated_hog[["Payable req.no", "Invoice Number", "Doc.Type",
                                                         "Name",
                                                         "ReimbursementID",
                                                         "category",
                                                         "Profit Ctr",
                                                         "BP",
                                                         "CostctrName",
                                                         "Cost Center",
                                                         "G/L",
                                                         "G/L Name",
                                                         "GR/IC Reference",
                                                         "Text",
                                                         "Org.unit",
                                                         "Status",
                                                         "File 1",
                                                         "File 2",
                                                         "File 3",
                                                         "Document No",
                                                         "Creator ID",
                                                         "Time",
                                                         "Updated at",
                                                         "HOD NAME",
                                                         "HOD Apr/Rej by",
                                                         "HOD Apr/Rej at",
                                                         "Reason for Rejection",
                                                         "Verifier ID",
                                                         "Verified at",
                                                         "Reference document",
                                                         "Reference invoice",
                                                         "Adv.doc year",
                                                         "Request no (Advance mulitple selection)",
                                                         "ID",
                                                         "Invoice Reference Number",
                                                         "HOG NAME",
                                                         "HOG Approval by", 'Department',
                                                         "HOG Approval at",
                                                         "HOG Approval Req",
                                                         "Requested HOG ID",
                                                         "Month",
                                                         "Vesselcode",
                                                         "PEA Number",
                                                         "Status of Request",
                                                         "Clearing doc no.",
                                                         "Document Date",
                                                         "Posting Date",
                                                         "Amount",
                                                         "On",
                                                         "Updated on",
                                                         "HOD Apr/Rej on",
                                                         "Verified on",
                                                         "HOG Approval on",
                                                         "Clearing date", 'year'
                                                         ]]

                    concatenated_hog.reset_index(drop=True, inplace=True)
                    concatenated_hog.index += 1  # Start index from 1
                    # st.dataframe(concatenated)
                    columns_to_convert = ['Payable req.no', 'Doc.Type', 'HOD Apr/Rej on',
                                          'HOG Approval by',
                                          'year',
                                          'Invoice Number', 'Text', 'Cost Center', 'G/L',
                                          'Document No',
                                          'Creator ID', 'Verifier ID']
                    concatenated_hog[columns_to_convert] = concatenated_hog[columns_to_convert].astype(str)

                    ccc, card01, c111, card1, middle_column, card2, c222 = st.columns(
                        [1, 2, 1, 2, 1, 2, 1])
                    with card01:
                        total_HOGs = concatenated_hog['HOG Approval by'].nunique()
                        total_HOGs_str = str(total_HOGs)
                        st.markdown(
                            f"<h3 style='text-align: center; font-size: 25px;'>NO Of HOGs</h3>",
                            unsafe_allow_html=True
                        )
                        st.write("")
                        st.markdown(
                            f"<div style='{CARD1_STYLE}'>"
                            f"<h2 style='color: #28a745; text-align: center; font-size: 35px;'>{total_HOGs_str}</h2>"
                            "</div>",
                            unsafe_allow_html=True
                        )
                    with card1:
                        Total_Amount_Alloted = concatenated_hog['Amount'].sum()

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
                            unsafe_allow_html=True
                        )
                        st.markdown(
                            f"<div style='{CARD1_STYLE}'>"
                            f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>{amount_display}</h2>"
                            "</div>",
                            unsafe_allow_html=True
                        )

                    with card2:
                        Total_Transaction = len(concatenated_hog)
                        st.markdown(
                            f"<h3 style='text-align: center; font-size: 25px;'>Count Of Transactions</h3>",
                            unsafe_allow_html=True
                        )
                        st.markdown(
                            f"<div style='{CARD2_STYLE}'>"
                            f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                            "</div>",
                            unsafe_allow_html=True
                        )
                    concatenated_hog['Amount'] = concatenated_hog['Amount'].round()

                    concatenated_hog.reset_index(drop=True, inplace=True)
                    concatenated_hog.index += 1  # Start index from 1
                    columns_to_convert = ['Payable req.no', 'Doc.Type',
                                          'Invoice Number', 'Text', 'Cost Center', 'G/L',
                                          'Document No',
                                          'Creator ID', 'Verifier ID', 'HOG Approval by']
                    concatenated_hog[columns_to_convert] = concatenated_hog[columns_to_convert].astype(str)
                    concatenated_hog['Document No'] = concatenated_hog['Document No'].apply(
                        lambda x: str(x) if isinstance(x, str) else '')
                    concatenated_hog['Document No'] = concatenated_hog['Document No'].apply(
                        lambda x: re.sub(r'\..*', '', x))
                    concatenated_hog['HOG Approval by'] = concatenated_hog['HOG Approval by'].astype(str)
                    concatenated_hog['HOG Approval by'] = concatenated_hog['HOG Approval by'].apply(
                        lambda x: str(x) if isinstance(x, str) else '')
                    concatenated_hog['HOG Approval by'] = concatenated_hog['HOG Approval by'].apply(
                        lambda x: re.sub(r'\..*', '', x))

                    concatenated_show = concatenated_hog.copy()
                    grouped_df = concatenated_show.groupby('HOG Approval by')
                    # Create 'Total transactions' column
                    concatenated_show['value'] = grouped_df['Amount'].transform('sum')
                    concatenated_show['Count of transactions'] = grouped_df[
                        'HOG Approval by'].transform(
                        len)
                    # Create 'No of Days' column with unique values of 'HOG Approval on'
                    concatenated_show['No of Days'] = grouped_df['HOG Approval on'].transform(
                        'nunique')
                    concatenated_show = concatenated_show.rename(
                        columns={'HOG Approval by': 'HOG ID'
                                 })
                    concatenated_show.sort_values(by='value', ascending=False, inplace=True)

                    concatenated_show = concatenated_show.drop_duplicates(subset='HOG ID',
                                                                          keep='last')

                    concatenated_show.reset_index(drop=True, inplace=True)
                    concatenated_show.index += 1  # Start index from 1
                    ccc, cdf1, cdf2, cd3 = st.columns([1, 2, 6, 2])
                    cdf2.write("")
                    cdf2.dataframe(
                        concatenated_show[
                            ['HOG ID', 'HOG NAME', 'Department', 'value', 'Count of transactions',
                             'No of Days']]
                    )
                    concatenated_hog.sort_values(by=['HOG Approval by'], inplace=True)
                    concatenated_hog.reset_index(drop=True, inplace=True)
                    concatenated_hog.index += 1  # Start index from 1
                    space11, cdf1, cdf2, colg2 = st.columns([1, 4, 4, 1])

                    # Create Pie chart for 'value' column
                    fig = px.pie(concatenated_show, values='value', names='Department',
                                 title='Value Wise'
                                 )
                    fig.update_layout(width=400, height=400)
                    fig.update_traces(textinfo='percent+label')
                    fig.update(layout_title_text='',
                               layout_showlegend=False)
                    cdf1.write("")
                    cdf1.markdown(
                        f"<h3 style='text-align: left; font-size: 20px;'>Value Wise</h3>",
                        unsafe_allow_html=True
                    )
                    cdf1.plotly_chart(fig)

                    # Create Pie chart for 'No of Days' column
                    fig2 = px.pie(concatenated_show, values='No of Days', names='Department',
                                  title='Day Wise')
                    fig2.update_layout(width=400, height=400)
                    fig2.update_traces(textinfo='percent+label')
                    fig2.update(layout_title_text='',
                                layout_showlegend=False)
                    cdf2.write("")
                    cdf2.markdown(
                        f"<h3 style='text-align: left; font-size: 20px;'>Day Wise</h3>",
                        unsafe_allow_html=True
                    )
                    cdf2.plotly_chart(fig2)

                    concatenated_show = concatenated_show[
                        ['HOG ID', 'HOG NAME', 'Department', 'value', 'Count of transactions',
                         'No of Days']
                    ]

                    excel_buffer = BytesIO()
                    concatenated_hog.to_excel(excel_buffer, index=False)
                    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                        concatenated_show.to_excel(writer, index=False, sheet_name='Summary')
                        concatenated_hog.to_excel(writer, index=False, sheet_name='Raw Data')
                    excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
                    # Convert Excel buffer to base64
                    excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
                    # Download link for Excel file within a Markdown
                    download_link = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}" download="Approval Exceptions.xlsx">Download Excel file</a>'
                    st.markdown(download_link, unsafe_allow_html=True)
                    # excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
                    # # Convert Excel buffer to base64
                    # excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
                    # # Download link for Excel file within a Markdown
                    # download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Approval Exceptions.xlsx">Download Excel file</a>'
                    # st.markdown(download_link, unsafe_allow_html=True)

    with t52:
        filtered_df = exceptions2.copy()
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
                "year": "Year",
            },
            inplace=True,
        )
        filtered_df = filtered_df[filtered_df["Verifier ID"].str.contains(r'[a-zA-Z]')]
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
            "HOG Approval by",
        ]
        filtered_df[columns_to_convert] = filtered_df[columns_to_convert].astype(str)

        # Displaying total amount allotted
        colu1, colu2, _, colu3, _ = st.columns([1, 4, 1, 4, 1])
        Pick = ["All"] + [yr for yr in filtered_df["Year"].unique() if yr != "All"]
        Pickchoice = colu1.selectbox("Pick an option", Pick, index=0)

        if Pickchoice != "All":
            filtered_df = filtered_df[filtered_df["Year"] == Pickchoice]
        else:
            filtered_df = filtered_df.copy()
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
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'>Reimbursement Amount (in Rupees)</h3>",
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
            by=["ReimbursementID", "Amount", "Cost Center"], inplace=True
        )

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
        FILENAME = "Special Exceptions.xlsx"
        download_link = f'<a href="data:file/xls;base64,{excel_b64}" download=FILENAME>Download Excel file</a>'
        st.markdown(download_link, unsafe_allow_html=True)




