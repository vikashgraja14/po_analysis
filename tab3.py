from analyze_excelfinal import *
import base64


@st.cache_resource(show_spinner=False)
def has_special_characters(s):
    return re.search(r'[^A-Za-z0-9]+', s) is not None

@st.cache_resource(show_spinner=False)
def is_similar(s1, s2):
    # Remove special characters from both strings
    s1_clean = re.sub(r'[^A-Za-z0-9]+', '', str(s1))
    s2_clean = re.sub(r'[^A-Za-z0-9]+', '', str(s2))
    # Check if the cleaned strings are equal
    return s1_clean == s2_clean

def Special_exceptions(specialexp):
    filtered_df = specialexp
    filtered_df['Invoice Number'] = filtered_df['Invoice Number'].astype(str)
    def special_rows(filtered_df):
        filtered_df.rename(
            columns={
                'Vendor Name': 'Name',
                'Pstng Date': 'Posting Date',
                'Type': "Doc.Type",
                'Vendor': "Reimbursement ID",
                "Cost Ctr": "Cost Center",
                'Doc. Date': 'Document Date',
                "Verified by": 'Verifier',
                'Created': 'Creator',
                'year': 'Year'
            },
            inplace=True
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
        c11, c2, c3, c4 = st.columns(4)
        filteryears = ["All"] + [yr for yr in filtered_df["Year"].unique() if yr != "All"]
        filtertheyears = c11.selectbox("SELECT year", filteryears, index=0)
        if filtertheyears != "All":
            filtered_df = filtered_df[filtered_df["Year"] == filtertheyears]
        else:
            filtered_df = filtered_df.copy()
        filtered_df = filtered_df.sort_values(by="Invoice Number")
        filtered_df.reset_index(drop=True, inplace=True)
        filtered_df.index += 1
        return filtered_df
    filtered_df = special_rows(filtered_df)
    if filtered_df is not None:
        if filtered_df.empty:
            st.markdown(
                "<div style='text-align: center; font-weight: bold; color: black;'>No entries</div>",
                unsafe_allow_html=True)
        else:
            st.write(filtered_df)
            # st.markdown(
            #     f"<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>No entries</h2>",
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
                    f"<div style='{CARD1_STYLE}'>"
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
                    f"<div style='{CARD2_STYLE}'>"
                    f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                    "</div>",
                    unsafe_allow_html=True
                )
            # filtered_df['Document Date'] = pd.to_datetime(filtered_df['Document Date'], format='%Y/%m/%d',
            #                                               errors='coerce')
            # filtered_df['Posting Date'] = pd.to_datetime(filtered_df['Posting Date'], format='%Y/%m/%d',
            #                                              errors='coerce')
            # filtered_df['Verified on'] = pd.to_datetime(filtered_df['Verified on'], format='%Y/%m/%d',
            #                                             errors='coerce')
            # filtered_df['Posting Date'] = filtered_df['Posting Date'].dt.date
            # filtered_df['Document Date'] = filtered_df['Document Date'].dt.date
            # filtered_df['Verified on'] = filtered_df['Verified on'].dt.date
            # Ensure that 'filtered_df1' is a copy of 'filtered_df' with rounded 'Amount'
            filtered_df1 = filtered_df.copy()
            filtered_df1['Amount'] = filtered_df1['Amount'].round()

            filtered_df1.reset_index(drop=True, inplace=True)
            filtered_df1.index += 1  # Start index from 1
            st.dataframe(
                filtered_df1
                    # [['Payable req.no', 'Doc.Type', 'Reimbursement ID', 'Amount', 'Document Date', 'Name',
                    #  'Invoice Number', 'Text', 'Cost Center', 'G/L', 'Document No',
                    #  'Posting Date', 'Creator', 'Verifier', 'HOD Apr/Rej by','HOG Approval by']]
            )

            # Generate and provide a download link for the Excel file
            filename = f"Invoice Special Charecters.xlsx"
            filtered_df.reset_index(drop=True, inplace=True)
            filtered_df.index += 1  # Start index from 1
            excel_buffer = BytesIO()
            filtered_df.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
            # Convert Excel buffer to base64
            excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()

            download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="{filename}">Download Excel file</a>'
            st.markdown(download_link, unsafe_allow_html=True)

