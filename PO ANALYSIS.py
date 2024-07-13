        with t41:
            def ChangeWidgetFontSize(wgt_txt, wch_font_size='12px'):
                htmlstr = """<script>var elements = window.parent.document.querySelectorAll('*'), i;
                                for (i = 0; i < elements.length; ++i) { if (elements[i].innerText == |wgt_txt|)
                                    { elements[i].style.fontSize='""" + wch_font_size + """';} } </script>  """

                htmlstr = htmlstr.replace('|wgt_txt|', "'" + wgt_txt + "'")
                components.html(f"{htmlstr}", height=0, width=0)

            # Renaming columns in the DataFrame
            radio['Doc. Date'] = pd.to_datetime(radio['Doc. Date'], format='%Y/%m/%d', errors='coerce')
            radio['Pstng Date'] = pd.to_datetime(radio['Pstng Date'], format='%Y/%m/%d', errors='coerce')
            radio['Verified on'] = pd.to_datetime(radio['Verified on'], format='%Y/%m/%d', errors='coerce')
            dfholiday['date'] = pd.to_datetime(dfholiday['date'], format='%Y/%m/%d', errors='coerce')

            # Renaming columns in the DataFrame
            radio.rename(
                columns={
                    'Vendor Name': 'Name',
                    'Pstng Date' :'Posting Date',
                    'Type': "Doc.Type",
                    'Vendor': "Reimbursement ID",
                    "Cost Ctr": "Cost Center",
                    'Doc. Date': 'Document Date',
                    "Verified by": 'Verifier',
                    'Created': 'Creator',
                    'year':'Year'
                },
                inplace=True
            )


            radio = radio[['Payable req.no', 'Doc.Type', 'Reimbursement ID', 'Amount', 'Document Date', 'Name','Year',
                           'Invoice Number', 'Text', 'Cost Center', 'G/L', 'Document No',
                           'Posting Date', 'Creator', 'Verifier', 'HOG Approval by', 'category',
                           'Reference invoice', 'CostctrName', 'G/L Name', 'Profit Ctr',
                           'GR/IC Reference', 'Org.unit', 'Status', 'File 1', 'File 2', 'File 3',
                           'Time', 'Updated at', 'Reason for Rejection', 'Verified at',
                           'Reference document', 'Adv.doc year',
                           'Request no (Advance mulitple selection)', 'Invoice Reference Number',
                           'HOG Approval at', 'HOG Approval Req', 'Requested HOG ID',
                           'Month', 'Vesselcode', 'PEA Number', 'Status of Request', 'Clearing doc no.',
                           'On', 'Updated on', 'Verified on',
                           'HOG Approval on', 'Clearing date']]
            radio['Document No'] = radio['Document No'].astype(str)
            radio['Document No'] = radio['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
            radio['Document No'] = radio['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
            c11, c2, c3, c4 = st.columns(4)
            options = ["All"] + [yr for yr in radio['Year'].unique() if yr != "All"]
            selected_option = c11.selectbox("Choose a Year", options, index=0)

            if selected_option != "All":
                radio = radio[radio['Year'] == selected_option]

            options = ["All"] + [yr for yr in radio['category'].unique() if yr != "All"]
            selected_option = c2.multiselect("Choose a category", options, default=["All"])

            if "All" in selected_option:
                radio = radio
            else:
                radio = radio[radio['category'].isin(selected_option)]

            filtered_df = radio.copy()
            columns_to_convert = ['Name', 'Invoice Number', "Reimbursement ID"]
            filtered_df[columns_to_convert] = filtered_df[columns_to_convert].astype(str)
            filtered_df['Document No'] = filtered_df['Document No'].astype(str)
            filtered_df['Document No'] = filtered_df['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
            filtered_df['Document No'] = filtered_df['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
            filtered_df['Document Date'] = pd.to_datetime(filtered_df['Document Date'], format='%Y/%m/%d', errors='coerce')
            filtered_df['Posting Date'] = pd.to_datetime(filtered_df['Posting Date'], format='%Y/%m/%d', errors='coerce')
            filtered_df['Verified on'] = pd.to_datetime(filtered_df['Verified on'], format='%Y/%m/%d', errors='coerce')
            dfholiday['date'] = pd.to_datetime(dfholiday['date'], format='%Y/%m/%d', errors='coerce')
            col1, col2, col3, col4, col5, col6, col7, col8, col9 = st.columns(9)
            checkbox_states = {
                'Reimbursement ID': col1.checkbox('Reimbursement ID', key='Reimbursement ID'),
                'Amount': col2.checkbox('Amount', key='Amount'),
                'Document Date': col3.checkbox('Document Date', key='Document Date'),
                'Cost Center': col4.checkbox('Cost Center', key='Cost Center'),
                'G/L': col5.checkbox('G/L', key='G/L'),
                'Invoice Number': col6.checkbox('Invoice Number', key='Invoice Number'),
                'Text': col7.checkbox('Text', key='Text'),
                'Inv-Special Character': col8.checkbox('Inv-Special Character', key='Inv-Special Character'),
                'Holiday Transactions': col9.checkbox('Holiday Transactions', key='Holiday Transactions')
            }
            @st.cache_resource(show_spinner=False)
            def has_special_characters(s):
                return re.search(r'[^A-Za-z0-9]+', s) is not None


            def is_similar(s1, s2):
                # Remove special characters from both strings
                s1_clean = re.sub(r'[^A-Za-z0-9]+', '', str(s1))
                s2_clean = re.sub(r'[^A-Za-z0-9]+', '', str(s2))
                # Check if the cleaned strings are equal
                return s1_clean == s2_clean

            checked_columns = [key for key, value in checkbox_states.items() if value]
            columns_to_check_for_duplicates = [column for column in checked_columns if
                                               column not in ['Holiday Transactions', 'Inv-Special Character']]
            filtered_df['Invoice Number'] = filtered_df['Invoice Number'].astype(str)
            @st.cache_resource(show_spinner=False)
            @st.experimental_fragment
            def filter_dataframe(filtered_df, checked_columns, columns_to_check_for_duplicates, dfholiday, filename):
                if not checked_columns:
                    st.error('Please select at least one column to check for duplicates.')
                    return None, None, None, None, None
                else:
                    if 'Holiday Transactions' in checked_columns:
                        if len(checked_columns) == 1:
                            filtered_df['Document Date'] = pd.to_datetime(filtered_df['Document Date'],
                                                                          format='%Y/%m/%d', errors='coerce')
                            filtered_df['Posting Date'] = pd.to_datetime(filtered_df['Posting Date'], format='%Y/%m/%d',
                                                                         errors='coerce')
                            filtered_df['Verified on'] = pd.to_datetime(filtered_df['Verified on'], format='%Y/%m/%d',
                                                                        errors='coerce')
                            dfholiday['date'] = pd.to_datetime(dfholiday['date'], format='%Y/%m/%d', errors='coerce')
                            filtered_df = filtered_df[filtered_df['Posting Date'].isin(dfholiday['date'])]
                            filename = "Holiday Transactions.xlsx"
                        elif len(checked_columns) == 2 and 'Inv-Special Character' in checked_columns:
                            filtered_df['Document Date'] = pd.to_datetime(filtered_df['Document Date'],
                                                                          format='%Y/%m/%d', errors='coerce')
                            filtered_df['Posting Date'] = pd.to_datetime(filtered_df['Posting Date'], format='%Y/%m/%d',
                                                                         errors='coerce')
                            filtered_df['Verified on'] = pd.to_datetime(filtered_df['Verified on'], format='%Y/%m/%d',
                                                                        errors='coerce')
                            dfholiday['date'] = pd.to_datetime(dfholiday['date'], format='%Y/%m/%d', errors='coerce')
                            filtered_df = filtered_df[filtered_df['Posting Date'].isin(dfholiday['date'])]
                            # Create a list of tuples with original and cleaned invoice numbers
                            invoice_pairs = [(invoice, re.sub(r'[^A-Za-z0-9]+', '', str(invoice)))
                                             for invoice in filtered_df['Invoice Number']]

                            # Initialize a set to store similar invoices
                            similar_invoices = set()

                            # Find all unique pairs where invoices are similar
                            for i, (inv1, clean_inv1) in enumerate(invoice_pairs):
                                for inv2, clean_inv2 in invoice_pairs[i + 1:]:
                                    if is_similar(clean_inv1, clean_inv2):
                                        similar_invoices.add(inv1)
                                        similar_invoices.add(inv2)

                            # Filter the dataframe to show only similar invoices
                            filtered_df = filtered_df[filtered_df['Invoice Number'].isin(similar_invoices)]

                            # Remove duplicates based on 'Invoice Number'
                            filtered_df = filtered_df[~filtered_df.duplicated(subset='Invoice Number', keep=False)]
                        elif len(checked_columns) > 2 and all(
                                item in checked_columns for item in ['Inv-Special Character', 'Holiday Transactions'] if
                                item != 'Reimbursement ID'):
                            filtered_df = filtered_df[filtered_df['Posting Date'].isin(dfholiday['date'])]
                            # Create a list of tuples with original and cleaned invoice numbers
                            # Group by 'Reimbursement ID'
                            grouped = filtered_df.groupby('Reimbursement ID')
                            similar_invoices = set()

                            for name, group in grouped:
                                if len(group) > 1:  # Apply logic only if group size is greater than one
                                    # Create a list of tuples with original and cleaned invoice numbers
                                    invoice_pairs = [(invoice, re.sub(r'[^A-Za-z0-9]+', '', str(invoice))) for invoice
                                                     in
                                                     group['Invoice Number']]
                                    # Find all unique pairs where invoices are similar within the group
                                    for i, (inv1, clean_inv1) in enumerate(invoice_pairs):
                                        for inv2, clean_inv2 in invoice_pairs[i + 1:]:
                                            if is_similar(clean_inv1, clean_inv2):
                                                similar_invoices.add(inv1)
                                                similar_invoices.add(inv2)

                            # Filter the dataframe to include only similar invoices
                            filtered_df = filtered_df[filtered_df['Invoice Number'].isin(similar_invoices)]
                            filtered_df = filtered_df[~filtered_df.duplicated(subset='Invoice Number', keep=False)]
                        elif len(checked_columns) == 3 and all(item in checked_columns for item in
                                                               ['Inv-Special Character', 'Holiday Transactions',
                                                                'Reimbursement ID']):
                            filtered_df = filtered_df[filtered_df['Posting Date'].isin(dfholiday['date'])]

                            # Group by 'Reimbursement ID'
                            grouped = filtered_df.groupby('Reimbursement ID')
                            similar_invoices = set()

                            for name, group in grouped:
                                if len(group) > 1:  # Apply logic only if group size is greater than one
                                    # Create a list of tuples with original and cleaned invoice numbers
                                    invoice_pairs = [(invoice, re.sub(r'[^A-Za-z0-9]+', '', str(invoice))) for invoice
                                                     in
                                                     group['Invoice Number']]
                                    # Find all unique pairs where invoices are similar within the group
                                    for i, (inv1, clean_inv1) in enumerate(invoice_pairs):
                                        for inv2, clean_inv2 in invoice_pairs[i + 1:]:
                                            if is_similar(clean_inv1, clean_inv2):
                                                similar_invoices.add(inv1)
                                                similar_invoices.add(inv2)

                            # Filter the dataframe to include only similar invoices
                            filtered_df = filtered_df[filtered_df['Invoice Number'].isin(similar_invoices)]
                            filtered_df = filtered_df[~filtered_df.duplicated(subset='Invoice Number', keep=False)]



                        else:

                            filtered_df = filtered_df[filtered_df['Posting Date'].isin(dfholiday['date'])]


                    elif len(checked_columns) == 1 and 'Inv-Special Character' in checked_columns :
                        # Create a list of tuples with original and cleaned invoice numbers
                        filtered_df.rename(
                            columns={
                                "Vendor": "Reimbursement ID",
                                'Pstng Date': "Posting Date"
                            },
                            inplace=True,
                        )
                        filtered_df = filtered_df.sort_values(
                            by=["Posting Date", "Invoice Number"])
                        # Group by 'Reimbursement ID'
                        grouped = filtered_df.groupby('Reimbursement ID')
                        similar_invoices = set()
                        for name, group in grouped:
                            if len(group) > 1:  # Apply logic only if group size is greater than one
                                # Create a list of tuples with original and cleaned invoice numbers
                                group = group.drop_duplicates(subset=['Reimbursement ID', 'Invoice Number'],
                                                              keep='last')
                                invoice_pairs = [(invoice, re.sub(r'[^A-Za-z0-9]+', '', str(invoice))) for invoice in
                                                 group['Invoice Number']]
                                # Find all unique pairs where invoices are similar within the group
                                for i, (inv1, clean_inv1) in enumerate(invoice_pairs):
                                    for inv2, clean_inv2 in invoice_pairs[i + 1:]:
                                        if is_similar(clean_inv1, clean_inv2):
                                            similar_invoices.add(inv1)
                                            similar_invoices.add(inv2)
                        filtered_df = filtered_df[filtered_df['Invoice Number'].isin(similar_invoices)]
                        filtered_df = filtered_df.sort_values(by="Invoice Number")
                        filtered_df.reset_index(drop=True, inplace=True)

                    elif len(
                            checked_columns) > 2 and 'Inv-Special Character' in checked_columns and 'Holiday Transactions' not in checked_columns:
                        # Filter the dataframe for duplicates based on specified columns

                        # Group by 'Reimbursement ID'
                        grouped = filtered_df.groupby('Reimbursement ID')
                        similar_invoices = set()

                        for name, group in grouped:
                            if len(group) > 1:  # Apply logic only if group size is greater than one
                                # Create a list of tuples with original and cleaned invoice numbers
                                invoice_pairs = [(invoice, re.sub(r'[^A-Za-z0-9]+', '', str(invoice))) for invoice in
                                                 group['Invoice Number']]
                                # Find all unique pairs where invoices are similar within the group
                                for i, (inv1, clean_inv1) in enumerate(invoice_pairs):
                                    for inv2, clean_inv2 in invoice_pairs[i + 1:]:
                                        if is_similar(clean_inv1, clean_inv2):
                                            similar_invoices.add(inv1)
                                            similar_invoices.add(inv2)

                        # Filter the dataframe to include only similar invoices
                        filtered_df = filtered_df[filtered_df['Invoice Number'].isin(similar_invoices)]
                        filtered_df = filtered_df[~filtered_df.duplicated(subset='Invoice Number', keep=False)]

                    elif 'Inv-Special Character' not in checked_columns and 'Holiday Transactions' not in checked_columns:
                         filtered_df = filtered_df[
                            filtered_df.duplicated(subset=columns_to_check_for_duplicates, keep=False)]

                    filtered_df['Posting Date'] = filtered_df['Posting Date'].dt.date
                    filtered_df['Document Date'] = filtered_df['Document Date'].dt.date
                    filtered_df['Verified on'] = filtered_df['Verified on'].dt.date
                # Sorting logic
                    checked_columns = [
                        'Posting Date' if col == 'Holiday Transactions'
                        else 'Invoice Number' if col == 'Inv-Special Character'
                        else col
                        for col in checked_columns
                    ]

                    sort_columns = checked_columns.copy()

                    if 'Reimbursement ID' not in checked_columns and 'Cost Center' not in checked_columns:
                        sort_columns += ['Reimbursement ID', 'Cost Center']
                    elif 'Reimbursement ID' in checked_columns and 'Cost Center' not in checked_columns:
                        sort_columns += ['Cost Center']
                    elif 'Cost Center' in checked_columns and 'Reimbursement ID' not in checked_columns:
                        sort_columns += ['Reimbursement ID']

                    filtered_df = filtered_df.sort_values(by=sort_columns)
                    filtered_df.reset_index(drop=True, inplace=True)
                    filtered_df.index += 1
                    filename = f"Transactions_with_same_column.xlsx"

                    try:
                        return filtered_df, checked_columns, columns_to_check_for_duplicates, dfholiday, filename
                    except Exception as e:
                        st.error(f'An error occurred: {e}')
                        return None, None, None, None, None
            filename = "Exceptions.xlsx"

            filtered_df, checked_columns2, columns_to_check_for_duplicates2, dfholiday2, filename = filter_dataframe(
                    filtered_df, checked_columns, columns_to_check_for_duplicates, dfholiday, filename
                )
            # Check if 'filtered_df' is not None before proceeding
            if filtered_df is not None:
                if filtered_df.empty:
                    st.markdown(
                    "<div style='text-align: center; font-weight: bold; color: black;'>No entries</div>",
                    unsafe_allow_html=True)
                else:
                    # st.markdown(
                    #     f"<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>ENTRIES with SAME {checked_columns}</h2>",
                    #     unsafe_allow_html=True)
                    c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
                    with card1:
                        Total_Amount_Alloted = filtered_df['Amount'].sum()

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
                            f"<h3 style='text-align: center; font-size: 25px;'>Amount of Exposure(in Rupees)</h3>",
                            unsafe_allow_html=True
                        )
                        st.markdown(
                            f"<div style='{card1_style}'>"
                            f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>{amount_display}</h2>"
                            "</div>",
                            unsafe_allow_html=True
                        )

                    with card2:
                        Total_Transaction = len(filtered_df)
                        st.markdown(
                            f"<h3 style='text-align: center; font-size: 25px;'>Count Of Transactions</h3>",
                            unsafe_allow_html=True
                        )
                        st.markdown(
                            f"<div style='{card2_style}'>"
                            f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                            "</div>",
                            unsafe_allow_html=True
                        )
                    filtered_df['Document Date'] = pd.to_datetime(filtered_df['Document Date'], format='%Y/%m/%d',
                                                                  errors='coerce')
                    filtered_df['Posting Date'] = pd.to_datetime(filtered_df['Posting Date'], format='%Y/%m/%d',
                                                                 errors='coerce')
                    filtered_df['Verified on'] = pd.to_datetime(filtered_df['Verified on'], format='%Y/%m/%d',
                                                                errors='coerce')
                    filtered_df['Posting Date'] = filtered_df['Posting Date'].dt.date
                    filtered_df['Document Date'] = filtered_df['Document Date'].dt.date
                    filtered_df['Verified on'] = filtered_df['Verified on'].dt.date
                    # Ensure that 'filtered_df1' is a copy of 'filtered_df' with rounded 'Amount'
                    filtered_df1 = filtered_df.copy()
                    filtered_df1['Amount'] = filtered_df1['Amount'].round()

                    filtered_df1.reset_index(drop=True, inplace=True)
                    filtered_df1.index += 1  # Start index from 1
                    st.dataframe(
                        filtered_df1[
                            ['Payable req.no', 'Doc.Type', 'Reimbursement ID', 'Amount', 'Document Date', 'Name',
                             'Invoice Number', 'Text', 'Cost Center', 'G/L', 'Document No',
                             'Posting Date', 'Creator', 'Verifier', 'HOG Approval by']]
                    )

                    # Generate and provide a download link for the Excel file
                    filename = f"SAME {checked_columns}.xlsx"
                    filtered_df.reset_index(drop=True, inplace=True)
                    filtered_df.index += 1  # Start index from 1
                    excel_buffer = BytesIO()
                    filtered_df.to_excel(excel_buffer, index=False)
                    excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
                    # Convert Excel buffer to base64
                    excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()

                    download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="{filename}">Download Excel file</a>'
                    st.markdown(download_link, unsafe_allow_html=True)
