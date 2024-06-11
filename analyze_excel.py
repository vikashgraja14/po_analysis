import streamlit as st
import base64
from io import BytesIO
import pandas as pd
import plotly.express as px
import os
import re
directory = os.path.dirname(__file__)
os.chdir(directory)
@st.cache_resource(show_spinner=False)
def process_data(files):
    data = pd.read_excel(files)
    df1 = pd.read_excel(
        "Cost Center.xlsx",
        header=1)
    df2 = pd.read_excel(
        "GL Master_.xlsx",
        header=1)
    df = data.copy()
    dfcost = df1.copy()
    dfGL = df2.copy()
    dfcost = dfcost.rename(columns={'Short text': 'CostctrName'})
    dfcost = dfcost[['Cost Ctr', 'CostctrName']]
    dfGL = dfGL.rename(columns={'Long Text': 'G/L Name'})
    dfGL = dfGL.rename(columns={'G/L Acct': 'G/L'})
    mapping = dict(zip(dfGL['G/L'], dfGL['G/L Name']))
    df['G/L Name'] = df['G/L'].map(mapping)
    mapping = dict(zip(dfcost['Cost Ctr'], dfcost['CostctrName']))
    df['CostctrName'] = df['Cost Ctr'].map(mapping)
    df = df[df['Status of Request'] == 'ALL APPROVALS ARE DONE']
    df = df[df['Clearing doc no.'] != '']
    filtered_df = df
    filtered_df.reset_index(drop=True, inplace=True)
    filtered_df.index += 1
    def categorize_id(id):
        if id[0].isalpha() and id[:3] != 'N00':
            return 'Vendor'
        elif id[0].isdigit() and len(id) <= 4:
            return 'Others'
        elif (id[0] in ['5', '6', '9'] and len(id) >= 7) or (id[:3] == 'N00' and len(id) == 6) or len(id) >= 7:
            return 'Korean Expats'
        else:
            return 'Employee'

    filtered_df['category'] = filtered_df['Vendor'].apply(categorize_id)
    filtered_df['Time'] = pd.to_datetime(filtered_df['Time'], format='%H:%M:%S', errors='coerce')
    filtered_df['Updated at'] = pd.to_datetime(filtered_df['Updated at'], format='%H:%M:%S', errors='coerce')
    filtered_df['Verified at'] = pd.to_datetime(filtered_df['Verified at'], format='%H:%M:%S', errors='coerce')
    filtered_df['HOG Approval at'] = pd.to_datetime(filtered_df['HOG Approval at'], format='%H:%M:%S', errors='coerce')
    filtered_df['Doc. Date'] = pd.to_datetime(filtered_df['Doc. Date'], errors='coerce')
    filtered_df['Pstng Date'] = pd.to_datetime(filtered_df['Pstng Date'], errors='coerce')
    filtered_df['On'] = pd.to_datetime(filtered_df['On'], errors='coerce')
    filtered_df['Updated on'] = pd.to_datetime(filtered_df['Updated on'], errors='coerce')
    filtered_df['Verified on'] = pd.to_datetime(filtered_df['Verified on'], errors='coerce')
    filtered_df['HOG Approval on'] = pd.to_datetime(filtered_df['HOG Approval on'], errors='coerce')
    filtered_df['Clearing date'] = pd.to_datetime(filtered_df['Clearing date'], errors='coerce')
    filtered_df['year'] = filtered_df['Pstng Date'].dt.year
    filtered_df['Doc. Date'] = filtered_df['Doc. Date'].dt.date
    filtered_df['Pstng Date'] = filtered_df['Pstng Date'].dt.date
    filtered_df['On'] = filtered_df['On'].dt.date
    filtered_df['Updated on'] = filtered_df['Updated on'].dt.date
    filtered_df['Verified on'] = filtered_df['Verified on'].dt.date
    filtered_df['HOG Approval on'] = filtered_df['HOG Approval on'].dt.date
    filtered_df['Clearing date'] = filtered_df['Clearing date'].dt.date
    filtered_df['Time'] = filtered_df['Time'].dt.time
    filtered_df['Updated at'] = filtered_df['Updated at'].dt.time
    filtered_df['Verified at'] = filtered_df['Verified at'].dt.time
    filtered_df['HOG Approval at'] = filtered_df['HOG Approval at'].dt.time
    filtered_df.drop(columns=['Year'], inplace=True)
    df = filtered_df.copy()
    df.reset_index(drop=True, inplace=True)
    df['Vendor'] = df['Vendor'].astype(str)
    df['G/L'] = df['G/L'].astype(str)
    df['G/L'] = df['G/L'].apply(lambda x: str(x) if isinstance(x, str) else '')
    df['G/L'] = df['G/L'].apply(lambda x: re.sub(r'\..*', '', x))
    df.reset_index(drop=True, inplace=True)
    return df
@st.cache_resource(show_spinner=False)
def line_plot_overall_transactions(data, category, years, width=400, height=300):
    filtered_data = data[(data['category'] == category) & (data['year'].isin(years))]
    data_length = filtered_data.groupby('year').size().reset_index(name='data_len')

    # Create the line plot
    fig = px.line(data_length, x='year', y='data_len',
                  title='Transactions Count ',
                  labels={'year': 'Year', 'data_len': 'Transactions'},
                  markers=True)

    # Set mode to 'lines+markers'
    fig.update_traces(mode='lines+markers')

    # Update layout with width and height
    fig.update_layout(width=width, height=height)

    # Modify x-axis labels to include hyphens between years
    fig.update_xaxes(tickvals=years,
                     ticktext=[re.sub(r'(\d{4})(\d{2})', r'\1-\2', str(year)) for year in years])

    # Format y-axis ticks as integers
    fig.update_yaxes(tickformat=".0f")

    # Add text annotations to data points
    for year, count in zip(data_length['year'], data_length['data_len']):
        fig.add_annotation(
            x=year,
            y=count,
            text=str(count),
            showarrow=False,
            font=dict(size=12, color='black'),
            align='center',
            yshift=13
        )

    return fig
@st.cache_resource(show_spinner=False)
def line_plot_used_amount(data, category, years, width=400, height=300):
    # Filter the data
    data = data[(data['category'] == category) & (data['year'].isin(years))]

    # Group by year and sum the amounts
    amount_length = data.groupby('year')['Amount'].sum().reset_index(name='amount_length')

    # Create the line plot
    fig = px.line(amount_length, x='year', y='amount_length',
                  title='Transactions Value',
                  labels={'year': 'Year', 'amount_length': 'Amount'},
                  markers=True)

    # Set mode to 'lines+markers'
    fig.update_traces(mode='lines+markers')

    # Update layout with width and height
    fig.update_layout(width=width, height=height)

    # Modify x-axis labels to include hyphens between years
    fig.update_xaxes(tickvals=years,
                     ticktext=[re.sub(r'(\d{4})(\d{2})', r'\1-\2', str(year)) for year in years])

    # Define the format_amount function
    def format_amount(amount):
        integer_part = int(amount)
        integer_length = len(str(integer_part))
        if integer_length > 5 and integer_length <= 7:
            return f"₹ {amount / 100000:,.2f} lks"
        elif integer_length > 7:
            return f"₹ {amount / 10000000:,.2f} crs"
        else:
            return f"₹ {amount:,.2f}"

    # Add formatted annotations on top of data points
    for year, amount in zip(amount_length['year'], amount_length['amount_length']):
        formatted_amount = format_amount(amount)
        fig.add_annotation(
            x=year,
            y=amount,
            text=formatted_amount,
            showarrow=False,
            font=dict(size=12, color='black'),
            align='center',
            yshift=13
        )

    return fig
card1_style = """
             display: flex;
             flex-direction: column;
             justify-content: center;
             align-items: center;
             background-color: #ffffff;
             padding: 10px;
             border-radius: 10px;
             font-style: bold:
             box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
             max-width: 200px; /* Adjust the max-width as needed */
             margin-left: 50px; /* Set left margin to auto */
             # margin-right: auto; /* Set right margin to auto */
         """
card2_style = """
             display: flex;
             flex-direction: column;
             justify-content: center;
             align-items: center;
             background-color: #ffffff;
             padding: 10px;
             border-radius: 10px;
             box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
             max-width: 200px; /* Adjust the max-width as needed */
             margin-left: 50px; /* Set left margin to auto */
             # margin-right: auto; /* Set right margin to auto */
         """

def display_duplicate_invoices(exceptions):
    dfe = exceptions.copy()
    dfe['Document No'] = dfe['Document No'].astype(str)
    dfe['Document No'] = dfe['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
    dfe['Document No'] = dfe['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
    dfe.rename(columns={'Vendor Name': 'Name', 'Type': "Doc.Type", 'Vendor': "ID", "Cost Ctr": "Cost Center"},
               inplace=True)
    c1, c2, c3, c4, c5 = st.columns(5)
    options = ["All"] + [yr for yr in dfe['year'].unique() if yr != "All"]
    selected_option = c1.selectbox("Select an Year", options, index=0)
    # Filter the DataFrame based on the selected option
    if selected_option == "All":
        dfe = dfe  # Return the entire DataFrame
    else:
        dfe = dfe[dfe['year'] == selected_option]
    # Sorting by 'Invoice Number'
    dfe.sort_values(by='Invoice Number', ascending=True, inplace=True)
    columns_to_convert = ['Name', 'Amount', 'Doc. Date', 'Invoice Number', "ID"]
    dfe[columns_to_convert] = dfe[columns_to_convert].astype(str)
    selected_option = c2.selectbox("Select Duplicate type",
                                   ["Duplicate Invoice", "Reimbursement Amount_ID_DocDate_SAME",
                                    "Reimbursement Amount_CostCtr_70%Inv_SAME"], index=0)
    # Filter the DataFrame based on the selected option
    if selected_option == "Reimbursement Amount_ID_DocDate_SAME":
        dfe = dfe[dfe.duplicated(subset=['ID', 'Amount','Doc. Date'], keep=False)]
        dfe = dfe.sort_values(by=['ID', 'Amount'])
        filename = "Reimbursement Amount_ID_DocDate_SAME.xlsx"
    elif selected_option == "Reimbursement Amount_CostCtr_70%Inv_SAME":
        grouped_df = dfe.groupby(['Cost Center', 'Amount']).filter(lambda group: len(group) > 1)

        # Step 4: Compare values in column 'Invoice Number'
        def compare_strings(s1, s2):
            common_chars = set(s1) & set(s2)
            return len(common_chars) >= 0.7 * min(len(s1), len(s2))

        dfe = grouped_df[grouped_df.apply(
            lambda row: compare_strings(row['Invoice Number'], grouped_df['Invoice Number'].iloc[0]), axis=1)]
        dfe = dfe.groupby(['Amount', 'Cost Center']).filter(lambda group: len(group) > 1)
        # dfe = dfe.sort_values(by=['Amount','Invoice Number'])
        dfe.reset_index(drop=True, inplace=True)
        filename = "Reimbursement Amount_CostCtr_70%Inv_SAME.xlsx"
    elif selected_option == "Duplicate Invoice":
        dfe = dfe[dfe.duplicated(subset=['Invoice Number'], keep=False)]
        dfe = dfe.sort_values(by=['Invoice Number', 'ID', 'Name'], ascending=True)
        filename = "Duplicate Exceptions.xlsx"
        # dfe.sort_values(by='Invoice Number', ascending=True, inplace=True)
    # dfe = dfe.sort_values(by=['Invoice Number','ID'], ascending=True)
    # dfe.sort_values(by='Invoice Number', ascending=True, inplace=True)
    dfe.reset_index(drop=True, inplace=True)
    dfe.index += 1  # Start index from 1
    dfe['Amount'] = pd.to_numeric(dfe['Amount'], errors='coerce')
    dfe['Amount'] = dfe['Amount'].astype(int)
    st.write(
        "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>Entries with Duplicate Invoices</h2>",
        unsafe_allow_html=True)
    st.write("")
    if dfe.empty:
        st.write(
            "<div style='text-align: center; font-weight: bold; color: black;'>No entries with same Invoice Number</div>",
            unsafe_allow_html=True)
    else:
        c1, card1, middle_column, card2, c2 = st.columns([1, 4, 1, 4, 1])
        with card1:
            Total_Amount_Alloted = dfe['Amount'].sum()

            # Check if the length of Total_Amount_Alloted is greater than 5
            if len(str(Total_Amount_Alloted)) > 5:
                # Get the integer part of the total amount
                integer_part = int(Total_Amount_Alloted)
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
            st.write("")


        with card2:
            Total_Transaction = len(dfe)
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'> Count Of Transactions</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"<div style='{card2_style}'>"
                f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
            st.write("")
            # Display the DataFrame

        c1, c2, c3 = st.columns([1, 8, 1])
        dfe = dfe.drop(columns=['year'])
        dfe['Amount'] = dfe['Amount'].round()
        # dfe = dfe.sort_values(by=['Name', 'Amount', 'Doc. Date'])
        # dfe.sort_values(by='Invoice Number', ascending=True, inplace=True)
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        c2.write(dfe[['Payable req.no', 'Doc.Type', 'ID', 'Name', 'Invoice Number', 'Text', "Cost Center",
                      'G/L', 'Document No',
                      'Doc. Date', 'Pstng Date', 'Amount']])
        dfe = dfe[['Payable req.no', 'Doc.Type', 'ID', 'Name', 'category', 'Invoice Number',
                   'Reference invoice', 'Text', 'Document No',
                   'Doc. Date', 'Pstng Date', "Cost Center", 'CostctrName', 'G/L', 'G/L Name',
                   'Profit Ctr', 'GR/IC Reference', 'Org.unit', 'Status', 'File 1', 'File 2', 'File 3',
                   'Created', 'Time', 'Updated at', 'Reason for Rejection', 'Verified by', 'Verified at',
                   'Reference document', 'Reference invoice', 'Adv.doc year',
                   'Request no (Advance mulitple selection)', 'Invoice Reference Number',
                   'HOG Approval by', 'HOG Approval at', 'HOG Approval Req', 'Requested HOG ID',
                   'Month', 'Vesselcode', 'PEA Number', 'Status of Request', 'Clearing doc no.',
                   'Amount', 'On', 'Updated on', 'Verified on',
                   'HOG Approval on', 'Clearing date']]
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        excel_buffer = BytesIO()
        dfe.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
        # Convert Excel buffer to base64
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()

        download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="{filename}">Download Excel file</a>'
        st.markdown(download_link, unsafe_allow_html=True)

def same_Creator_Verified_HOG(exceptions):
    dfe = exceptions.copy()
    dfe = dfe[(dfe['Created'] == dfe['Verified by']) & (dfe['Verified by'] == dfe['HOG Approval by'])]
    dfe = dfe
    dfe['Document No'] = dfe['Document No'].astype(str)
    dfe['Document No'] = dfe['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
    dfe['Document No'] = dfe['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
    dfe.rename(columns={'Vendor Name': 'Name', 'Type': "Doc.Type", 'Vendor': "ID", "Cost Ctr": " Cost Center",
                        'Created': 'Created by'},
               inplace=True)
    c1, c2, c3, c4 = st.columns(4)
    options = ["All"] + [yr for yr in dfe['year'].unique() if yr != "All"]
    selected_option = c1.selectbox("Select an Year", options, index=0)
    # Filter the DataFrame based on the selected option
    if selected_option == "All":
        dfe = dfe  # Return the entire DataFrame
    else:
        dfe = dfe[dfe['year'] == selected_option]
    dfe.reset_index(drop=True, inplace=True)
    dfe.index += 1  # Start index from 1
    st.write(
        "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>Entries with same Creator ID , Verified ID and HOG Approval</h2>",
        unsafe_allow_html=True)
    st.write("")
    if dfe.empty:
        st.write(
            "<div style='text-align: center; font-weight: bold; color: black;'>No entries with same Creator ID, Verified ID and HOG Approval ID</div>",
            unsafe_allow_html=True)
    else:
        c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
        with card1:
            Total_Amount_Alloted = dfe['Amount'].sum()

            # Check if the length of Total_Amount_Alloted is greater than 5
            if len(str(Total_Amount_Alloted)) > 5:
                # Get the integer part of the total amount
                integer_part = int(Total_Amount_Alloted)
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
            st.write("")


        with card2:
            Total_Transaction = len(dfe)
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'> Count Of Transactions</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"<div style='{card2_style}'>"
                f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
            st.write("")
        c11, c22, c33 = st.columns([1, 8, 1])
        # Display the DataFrame
        dfe = dfe.drop(columns=['year'])
        # Display the DataFrame
        dfe['Amount'] = dfe['Amount'].round()
        dfe = dfe.drop(columns=['year'])
        dfe = dfe.sort_values(by=['Created by', 'Cost Center'])
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1

        # st.markdown(download_link, unsafe_allow_html=True)
        dfe['Amount'] = dfe['Amount'].round()
        c22.write(dfe[
                      ['Payable req.no', 'Type', 'Vendor', 'Vendor Name', 'Invoice Number', 'Text', "Cost Ctr",
                       'G/L', 'Document No',
                       'Doc. Date', 'Pstng Date', 'Amount', 'Created by', 'Verified by', 'HOG Approval by']])
        dfe = dfe[['Payable req.no', 'Doc.Type', 'ID', 'Name', 'category', 'Invoice Number',
                   'Reference invoice', 'Text', 'Document No',
                   'Doc. Date', 'Pstng Date', "Cost Center", 'CostctrName', 'G/L', 'G/L Name',
                   'Profit Ctr', 'GR/IC Reference', 'Org.unit', 'Status', 'File 1', 'File 2', 'File 3',
                   'Created by', 'Time', 'Updated at', 'Reason for Rejection', 'Verified by', 'Verified at',
                   'Reference document', 'Reference invoice', 'Adv.doc year',
                   'Request no (Advance mulitple selection)', 'Invoice Reference Number',
                   'HOG Approval by', 'HOG Approval at', 'HOG Approval Req', 'Requested HOG ID',
                   'Month', 'Vesselcode', 'PEA Number', 'Status of Request', 'Clearing doc no.',
                   'Amount', 'On', 'Updated on', 'Verified on',
                   'HOG Approval on', 'Clearing date']]
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        excel_buffer = BytesIO()
        dfe.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
        # Convert Excel buffer to base64
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
        # Download link for Excel file within a Markdown
        download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Creator_Verifier_Approver(HOD_HOD)_SAME.xlsx">Download Excel file</a>'
        st.markdown(download_link, unsafe_allow_html=True)


def same_Creator_Verified_HOGno(exceptions):
    dfe = exceptions.copy()
    # dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
    dfe = dfe[
        (dfe['Created'] == dfe['Verified by']) & (dfe['HOG Approval by'].isna())]
    dfe['Document No'] = dfe['Document No'].astype(str)
    dfe['Document No'] = dfe['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
    dfe['Document No'] = dfe['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
    dfe.rename(
        columns={'Vendor Name': 'Name', 'Type': "Doc.Type", 'Vendor': "ID", "Cost Ctr": "Cost Center",
                 'Created': 'Created by'},
        inplace=True)
    c1, c2, c3, c4 = st.columns(4)
    options = ["All"] + [yr for yr in dfe['year'].unique() if yr != "All"]
    selected_option = c1.selectbox("Select an Year", options, index=0)
    # Filter the DataFrame based on the selected option
    if selected_option == "All":
        dfe = dfe  # Return the entire DataFrame
    else:
        dfe = dfe[dfe['year'] == selected_option]
    dfe.reset_index(drop=True, inplace=True)
    dfe.index += 1  # Start index from 1
    st.write(
        "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>Entries with same Creator ID , Verified ID  and HOG Approval Is None</h2>",
        unsafe_allow_html=True)
    st.write("")
    if dfe.empty:
        st.write(
            "<div style='text-align: center; font-weight: bold; color: black;'>No entries with same Creator ID, Verified ID and With No HOG Approval</div>",
            unsafe_allow_html=True)
    else:
        c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
        with card1:
            Total_Amount_Alloted = dfe['Amount'].sum()

            # Check if the length of Total_Amount_Alloted is greater than 5
            if len(str(Total_Amount_Alloted)) > 5:
                # Get the integer part of the total amount
                integer_part = int(Total_Amount_Alloted)
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
            st.write("")
        with card2:
            Total_Transaction = len(dfe)
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'> Count Of Transactions</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"<div style='{card2_style}'>"
                f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
            st.write("")
        c11, c22, c33 = st.columns([1, 8, 1])
        # Display the DataFrame
        dfe['Amount'] = dfe['Amount'].round()
        dfe = dfe.drop(columns=['year'])
        dfe = dfe.sort_values(by=['Name', 'Created by', "Cost Center"])
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        c22.write(dfe[
                      ['Payable req.no', 'Doc.Type', 'ID', 'Name', 'Invoice Number', 'Text', "Cost Center",
                       'G/L', 'Document No',
                       'Doc. Date', 'Pstng Date', 'Amount', 'Created by', 'Verified by', 'HOG Approval by']])
        dfe = dfe[['Payable req.no', 'Doc.Type', 'ID', 'Name', 'category', 'Invoice Number',
                   'Reference invoice', 'Text', 'Document No',
                   'Doc. Date', 'Pstng Date', "Cost Center", 'CostctrName', 'G/L', 'G/L Name',
                   'Profit Ctr', 'GR/IC Reference', 'Org.unit', 'Status', 'File 1', 'File 2', 'File 3',
                   'Created by', 'Time', 'Updated at', 'Reason for Rejection', 'Verified by', 'Verified at',
                   'Reference document', 'Reference invoice', 'Adv.doc year',
                   'Request no (Advance mulitple selection)', 'Invoice Reference Number',
                   'HOG Approval by', 'HOG Approval at', 'HOG Approval Req', 'Requested HOG ID',
                   'Month', 'Vesselcode', 'PEA Number', 'Status of Request', 'Clearing doc no.',
                   'Amount', 'On', 'Updated on', 'Verified on',
                   'HOG Approval on', 'Clearing date']]
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        excel_buffer = BytesIO()
        dfe.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
        # Convert Excel buffer to base64
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
        # Download link for Excel file within a Markdown
        download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Creator_Verifier_SAME_NO APPROVER.xlsx">Download Excel file</a>'
        st.markdown(download_link, unsafe_allow_html=True)


def Creator_Verified_HOGno(exceptions):
    dfe = exceptions.copy()
    # dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
    dfe = dfe[(~dfe['Created'].isna()) & (~dfe['Verified by'].isna()) & dfe['HOG Approval by'].isna()]
    dfe['Document No'] = dfe['Document No'].astype(str)
    dfe['Document No'] = dfe['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
    dfe['Document No'] = dfe['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
    dfe.rename(
        columns={'Vendor Name': 'Name', 'Type': "Doc.Type", 'Vendor': "ID", "Cost Ctr": "Cost Center",
                 'Created': 'Created by'},
        inplace=True)
    c1, c2, c3, c4 = st.columns(4)
    options = ["All"] + [yr for yr in dfe['year'].unique() if yr != "All"]
    selected_option = c1.selectbox("Select an Year", options, index=0)
    # Filter the DataFrame based on the selected option
    if selected_option == "All":
        dfe = dfe  # Return the entire DataFrame
    else:
        dfe = dfe[dfe['year'] == selected_option]
    options = ["All"] + [gl for gl in dfe['G/L'].unique() if gl != "All"]
    selected_option = c2.selectbox("Select a G/L", options, index=0)
    # Filter the DataFrame based on the selected option
    if selected_option == "All":
        dfe = dfe  # Return the entire DataFrame
    else:
        dfe = dfe[dfe['G/L'] == selected_option]
    dfe.reset_index(drop=True, inplace=True)
    dfe.index += 1  # Start index from 1
    st.write(
        "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>Entries With No HOG Approval</h2>",
        unsafe_allow_html=True)
    st.write("")
    if dfe.empty:
        st.write(
            "<div style='text-align: center; font-weight: bold; color: black;'>No entries Without HOG Approval</div>",
            unsafe_allow_html=True)
    else:
        c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
        with card1:
            Total_Amount_Alloted = dfe['Amount'].sum()

            # Check if the length of Total_Amount_Alloted is greater than 5
            if len(str(Total_Amount_Alloted)) > 5:
                # Get the integer part of the total amount
                integer_part = int(Total_Amount_Alloted)
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
            st.write("")
        with card2:
            Total_Transaction = len(dfe)
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'> Count Of Transactions</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"<div style='{card2_style}'>"
                f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
            st.write("")
        c11, c22, c33 = st.columns([1, 8, 1])
        # Display the DataFrame
        dfe = dfe.drop(columns=['year'])
        dfe['Amount'] = dfe['Amount'].round()
        dfe = dfe.sort_values(by=['Name', "Cost Center"])
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        c22.write(dfe[
                      ['Payable req.no', 'Doc.Type', 'ID', 'Name', 'Invoice Number', 'Text', "Cost Center",
                       'G/L', 'Document No',
                       'Doc. Date', 'Pstng Date', 'Amount', 'Created by', 'Verified by', 'HOG Approval by']])
        dfe = dfe[['Payable req.no', 'Doc.Type', 'ID', 'Name', 'category', 'Invoice Number',
                   'Reference invoice', 'Text', 'Document No',
                   'Doc. Date', 'Pstng Date', "Cost Center", 'CostctrName', 'G/L', 'G/L Name',
                   'Profit Ctr', 'GR/IC Reference', 'Org.unit', 'Status', 'File 1', 'File 2', 'File 3',
                   'Created by', 'Time', 'Updated at', 'Reason for Rejection', 'Verified by', 'Verified at',
                   'Reference document', 'Reference invoice', 'Adv.doc year',
                   'Request no (Advance mulitple selection)', 'Invoice Reference Number',
                   'HOG Approval by', 'HOG Approval at', 'HOG Approval Req', 'Requested HOG ID',
                   'Month', 'Vesselcode', 'PEA Number', 'Status of Request', 'Clearing doc no.',
                   'Amount', 'On', 'Updated on', 'Verified on',
                   'HOG Approval on', 'Clearing date']]
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        excel_buffer = BytesIO()
        dfe.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
        # Convert Excel buffer to base64
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
        # Download link for Excel file within a Markdown
        download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="No approver(HOD_HOG).xlsx">Download Excel file</a>'
        st.markdown(download_link, unsafe_allow_html=True)


def Creator_HOG(exceptions):
    dfe = exceptions.copy()
    # dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
    dfe = dfe[
        (dfe['Created'] == dfe['HOG Approval by'])]

    dfe['Document No'] = dfe['Document No'].astype(str)
    dfe['Document No'] = dfe['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
    dfe['Document No'] = dfe['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
    dfe.rename(
        columns={'Vendor Name': 'Name', 'Type': "Doc.Type", 'Vendor': "ID", "Cost Ctr": " Cost Center",
                 'Created': 'Created by'},
        inplace=True)
    c1, c2, c3, c4 = st.columns(4)
    options = ["All"] + [yr for yr in dfe['year'].unique() if yr != "All"]
    selected_option = c1.selectbox("Select an Year", options, index=0)
    # Filter the DataFrame based on the selected option
    if selected_option == "All":
        dfe = dfe  # Return the entire DataFrame
    else:
        dfe = dfe[dfe['year'] == selected_option]
    dfe.reset_index(drop=True, inplace=True)
    dfe.index += 1  # Start index from 1
    st.write(
        "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>Entries With Same CreatorID and HOG ApprovalID</h2>",
        unsafe_allow_html=True)
    st.write("")
    if dfe.empty:
        st.write(
            "<div style='text-align: center; font-weight: bold; color: black;'>No entries With Same CreatorID and HOG ApprovalID</div>",
            unsafe_allow_html=True)
    else:
        c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
        with card1:
            Total_Amount_Alloted = dfe['Amount'].sum()

            # Check if the length of Total_Amount_Alloted is greater than 5
            if len(str(Total_Amount_Alloted)) > 5:
                # Get the integer part of the total amount
                integer_part = int(Total_Amount_Alloted)
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
            st.write("")
        with card2:
            Total_Transaction = len(dfe)
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'> Count Of Transactions</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"<div style='{card2_style}'>"
                f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
            st.write("")
        c11, c22, c33 = st.columns([1, 8, 1])
        # Display the DataFrame
        dfe['Amount'] = dfe['Amount'].round()
        dfe = dfe.drop(columns=['year'])
        dfe = dfe.sort_values(by=['Name', 'Created by'])
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        c22.write(dfe[
                      ['Payable req.no', 'Doc.Type', 'ID', 'Name', 'Invoice Number', 'Text', "Cost Center",
                       'G/L', 'Document No',
                       'Doc. Date', 'Pstng Date', 'Amount', 'Created by', 'Verified by', 'HOG Approval by']])
        dfe = dfe[['Payable req.no', 'Doc.Type', 'ID', 'Name', 'category', 'Invoice Number',
                   'Reference invoice', 'Text', 'Document No',
                   'Doc. Date', 'Pstng Date', "Cost Center", 'CostctrName', 'G/L', 'G/L Name',
                   'Profit Ctr', 'GR/IC Reference', 'Org.unit', 'Status', 'File 1', 'File 2', 'File 3',
                   'Created by', 'Time', 'Updated at', 'Reason for Rejection', 'Verified by', 'Verified at',
                   'Reference document', 'Reference invoice', 'Adv.doc year',
                   'Request no (Advance mulitple selection)', 'Invoice Reference Number',
                   'HOG Approval by', 'HOG Approval at', 'HOG Approval Req', 'Requested HOG ID',
                   'Month', 'Vesselcode', 'PEA Number', 'Status of Request', 'Clearing doc no.',
                   'Amount', 'On', 'Updated on', 'Verified on',
                   'HOG Approval on', 'Clearing date']]
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        excel_buffer = BytesIO()
        dfe.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
        # Convert Excel buffer to base64
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
        # Download link for Excel file within a Markdown
        download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Creator_Approver(HOD_HOD)_SAME.xlsx">Download Excel file</a>'
        st.markdown(download_link, unsafe_allow_html=True)


def same_Creator_Verified(exceptions):
    dfe = exceptions.copy()
    # dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
    dfe = dfe[
        (dfe['Created'] == dfe['Verified by'])]
    dfe['Document No'] = dfe['Document No'].astype(str)
    dfe['Document No'] = dfe['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
    dfe['Document No'] = dfe['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
    dfe.rename(
        columns={'Vendor Name': 'Name', 'Type': "Doc.Type", 'Vendor': "ID", "Cost Ctr": "Cost Center",
                 'Created': 'Created by'},
        inplace=True)
    c1, c2, c3, c4 = st.columns(4)
    options = ["All"] + [yr for yr in dfe['year'].unique() if yr != "All"]
    selected_option = c1.selectbox("Select an Year", options, index=0)
    # Filter the DataFrame based on the selected option
    if selected_option == "All":
        dfe = dfe  # Return the entire DataFrame
    else:
        dfe = dfe[dfe['year'] == selected_option]
    dfe.reset_index(drop=True, inplace=True)
    dfe.index += 1  # Start index from 1
    st.write(
        "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>Entries With Same Creator ID And Verified ID</h2>",
        unsafe_allow_html=True)
    st.write("")
    if dfe.empty:
        st.write(
            "<div style='text-align: center; font-weight: bold; color: black;'>No entries with same Creator ID And Verified ID </div>",
            unsafe_allow_html=True)
    else:
        c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
        with card1:
            Total_Amount_Alloted = dfe['Amount'].sum()

            # Check if the length of Total_Amount_Alloted is greater than 5
            if len(str(Total_Amount_Alloted)) > 5:
                # Get the integer part of the total amount
                integer_part = int(Total_Amount_Alloted)
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
            st.write("")

        with card2:
            Total_Transaction = len(dfe)
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'> Count Of Transactions</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"<div style='{card2_style}'>"
                f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
            st.write("")
        c11, c22, c33 = st.columns([1, 8, 1])
        # Display the DataFrame
        dfe['Amount'] = dfe['Amount'].round()
        dfe = dfe.sort_values(by=['Created by', 'Name', "Cost Center"])
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        dfe = dfe.drop(columns=['year'])
        c22.write(dfe[
                      ['Payable req.no', 'Doc.Type', 'ID', 'Name', 'Invoice Number', 'Text', "Cost Center",
                       'G/L', 'Document No',
                       'Doc. Date', 'Pstng Date', 'Amount', 'Created by', 'Verified by', 'HOG Approval by']])
        dfe = dfe[['Payable req.no', 'Doc.Type', 'ID', 'Name', 'category', 'Invoice Number',
                   'Reference invoice', 'Text', 'Document No',
                   'Doc. Date', 'Pstng Date', "Cost Center", 'CostctrName', 'G/L', 'G/L Name',
                   'Profit Ctr', 'GR/IC Reference', 'Org.unit', 'Status', 'File 1', 'File 2', 'File 3',
                   'Created by', 'Time', 'Updated at', 'Reason for Rejection', 'Verified by', 'Verified at',
                   'Reference document', 'Reference invoice', 'Adv.doc year',
                   'Request no (Advance mulitple selection)', 'Invoice Reference Number',
                   'HOG Approval by', 'HOG Approval at', 'HOG Approval Req', 'Requested HOG ID',
                   'Month', 'Vesselcode', 'PEA Number', 'Status of Request', 'Clearing doc no.',
                   'Amount', 'On', 'Updated on', 'Verified on',
                   'HOG Approval on', 'Clearing date']]
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        excel_buffer = BytesIO()
        dfe.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
        # Convert Excel buffer to base64
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
        # Download link for Excel file within a Markdown
        download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Creator_Verifier_SAME.xlsx">Download Excel file</a>'
        st.markdown(download_link, unsafe_allow_html=True)


df = pd.read_excel(
    "unlocked holiday.xlsx")


def Approval_holidays(exceptions):
    dfe = exceptions.copy()
    # dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
    dfe['Doc. Date'] = pd.to_datetime(dfe['Doc. Date'], format='%Y/%m/%d', errors='coerce')
    dfe['Pstng Date'] = pd.to_datetime(dfe['Pstng Date'], format='%Y/%m/%d', errors='coerce')
    dfe['Verified on'] = pd.to_datetime(dfe['Verified on'], format='%Y/%m/%d', errors='coerce')
    df['date'] = pd.to_datetime(df['date'], format='%Y/%m/%d', errors='coerce')
    # df['date'] = pd.to_datetime(df['date'], format='%Y-%m-%d')
    dfe = dfe[dfe['Pstng Date'].isin(df['date'])]
    # dfe = dfe[dfe['Pstng Date'].isin(df['date']) | dfe['Doc. Date'].isin(df['date']) | dfe['Verified on'].isin(df['date'])]
    # dfe = dfe[dfe['Pstng Date'].isin(df['date'])]
    dfe['Pstng Date'] = dfe['Pstng Date'].dt.date
    dfe['Doc. Date'] = dfe['Doc. Date'].dt.date
    dfe['Verified on'] = dfe['Verified on'].dt.date

    dfe['Document No'] = dfe['Document No'].astype(str)
    dfe['Document No'] = dfe['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
    dfe['Document No'] = dfe['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
    dfe.rename(
        columns={'Vendor Name': 'Name', 'Type': "Doc.Type", 'Vendor': "ID", "Cost Ctr": "Cost Center",
                 'Created': 'Created by'},
        inplace=True)
    c1, c2, c3, c4 = st.columns(4)
    options = ["All"] + [yr for yr in dfe['year'].unique() if yr != "All"]
    selected_option = c1.selectbox("Select an Year", options, index=0)
    # Filter the DataFrame based on the selected option
    if selected_option == "All":
        dfe = dfe  # Return the entire DataFrame


    else:
        dfe = dfe[dfe['year'] == selected_option]
    dfe.reset_index(drop=True, inplace=True)
    dfe.index += 1  # Start index from 1

    st.write(
        "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>ENTRIES MADE ON HOLIDAYS</h2>",
        unsafe_allow_html=True)
    st.write("")
    if dfe.empty:
        st.write(
            "<div style='text-align: center; font-weight: bold; color: black;'>No entries made on holidays</div>",
            unsafe_allow_html=True)
    else:

        c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
        with card1:
            with card1:
                Total_Amount_Alloted = dfe['Amount'].sum()

                # Check if the length of Total_Amount_Alloted is greater than 5
                if len(str(Total_Amount_Alloted)) > 5:
                    # Get the integer part of the total amount
                    integer_part = int(Total_Amount_Alloted)
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
                st.write("")
            # Total_alloted = Total_Amount_Alloted
            # st.markdown(
            #     f"<div style='{card1_style}'>"
            #     f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f} </h2>"
            #     "</div>",
            #     unsafe_allow_html=True
            # )
            # st.write("")
        with card2:
            Total_Transaction = len(dfe)
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'> Count Of Transactions</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"<div style='{card2_style}'>"
                f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
            st.write("")
        c11, c22, c33 = st.columns([1, 8, 1])
        # Display the DataFrame
        dfe['Amount'] = dfe['Amount'].round()
        dfe = dfe.drop(columns=['year'])
        dfe = dfe.sort_values(by=['Pstng Date', "Cost Center", 'Name'])
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        # dfe = dfe.drop(columns=['year'])
        c22.write(dfe[
                      ['Payable req.no', 'Doc.Type', 'ID', 'Name', 'Invoice Number', 'Text', "Cost Center",
                       'G/L', 'Document No',
                       'Doc. Date', 'Pstng Date', 'Amount', 'Created by', 'Verified by', 'HOG Approval by']])
        dfe = dfe[['Payable req.no', 'Doc.Type', 'ID', 'Name', 'category', 'Invoice Number',
                   'Reference invoice', 'Text', 'Document No',
                   'Doc. Date', 'Pstng Date', "Cost Center", 'CostctrName', 'G/L', 'G/L Name',
                   'Profit Ctr', 'GR/IC Reference', 'Org.unit', 'Status', 'File 1', 'File 2', 'File 3',
                   'Created by', 'Time', 'Updated at', 'Reason for Rejection', 'Verified by', 'Verified at',
                   'Reference document', 'Reference invoice', 'Adv.doc year',
                   'Request no (Advance mulitple selection)', 'Invoice Reference Number',
                   'HOG Approval by', 'HOG Approval at', 'HOG Approval Req', 'Requested HOG ID',
                   'Month', 'Vesselcode', 'PEA Number', 'Status of Request', 'Clearing doc no.',
                   'Amount', 'On', 'Updated on', 'Verified on',
                   'HOG Approval on', 'Clearing date']]
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        excel_buffer = BytesIO()
        dfe.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
        # Convert Excel buffer to base64
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
        # Download link for Excel file within a Markdown
        download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Holiday Transactions.xlsx">Download Excel file</a>'
        st.markdown(download_link, unsafe_allow_html=True)


def Pstingverified_holidays(exceptions):
    dfe = exceptions.copy()
    # dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
    dfe['Doc. Date'] = pd.to_datetime(dfe['Doc. Date'], format='%Y/%m/%d', errors='coerce')
    dfe['Pstng Date'] = pd.to_datetime(dfe['Pstng Date'], format='%Y/%m/%d', errors='coerce')
    dfe['Verified on'] = pd.to_datetime(dfe['Verified on'], format='%Y/%m/%d', errors='coerce')
    df['date'] = pd.to_datetime(df['date'], format='%Y/%m/%d', errors='coerce')
    # df['date'] = pd.to_datetime(df['date'], format='%Y-%m-%d')
    dfe = dfe[(dfe['Pstng Date'] == dfe['Verified on']) & dfe['Pstng Date'].isin(df['date'])]

    # dfe = dfe[dfe['Pstng Date'].isin(df['date']) | dfe['Doc. Date'].isin(df['date']) | dfe['Verified on'].isin(df['date'])]
    # dfe = dfe[dfe['Pstng Date'].isin(df['date'])]
    dfe['Pstng Date'] = dfe['Pstng Date'].dt.date
    dfe['Doc. Date'] = dfe['Doc. Date'].dt.date
    dfe['Verified on'] = dfe['Verified on'].dt.date

    dfe['Document No'] = dfe['Document No'].astype(str)
    dfe['Document No'] = dfe['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
    dfe['Document No'] = dfe['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
    dfe.rename(
        columns={'Vendor Name': 'Name', 'Type': "Doc.Type", 'Vendor': "ID", "Cost Ctr": "Cost Center",
                 'Created': 'Created by'},
        inplace=True)
    c1, c2, c3, c4 = st.columns(4)
    options = ["All"] + [yr for yr in dfe['year'].unique() if yr != "All"]
    selected_option = c1.selectbox("Select an Year", options, index=0)
    # Filter the DataFrame based on the selected option
    if selected_option == "All":
        dfe = dfe  # Return the entire DataFrame


    else:
        dfe = dfe[dfe['year'] == selected_option]
    dfe.reset_index(drop=True, inplace=True)
    dfe.index += 1  # Start index from 1

    st.write(
        "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>ENTRIES MADE ON HOLIDAYS</h2>",
        unsafe_allow_html=True)
    st.write("")
    if dfe.empty:
        st.write(
            "<div style='text-align: center; font-weight: bold; color: black;'>No entries made on holidays</div>",
            unsafe_allow_html=True)
    else:

        c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
        with card1:
            with card1:
                Total_Amount_Alloted = dfe['Amount'].sum()

                # Check if the length of Total_Amount_Alloted is greater than 5
                if len(str(Total_Amount_Alloted)) > 5:
                    # Get the integer part of the total amount
                    integer_part = int(Total_Amount_Alloted)
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
                st.write("")

        with card2:
            Total_Transaction = len(dfe)
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'> Count Of Transactions</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"<div style='{card2_style}'>"
                f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
            st.write("")
        c11, c22, c33 = st.columns([1, 8, 1])
        # Display the DataFrame
        dfe['Amount'] = dfe['Amount'].round()
        dfe = dfe.drop(columns=['year'])
        dfe = dfe.sort_values(by=['Pstng Date', "Cost Center", 'Name'])
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        c22.write(dfe[
                      ['Payable req.no', 'Doc.Type', 'ID', 'Name', 'Invoice Number', 'Text', "Cost Center",
                       'G/L', 'Document No',
                       'Doc. Date', 'Pstng Date', 'Amount', 'Created by', 'Verified by', 'HOG Approval by']])
        dfe = dfe[['Payable req.no', 'Doc.Type', 'ID', 'Name', 'category', 'Invoice Number',
                   'Reference invoice', 'Text', 'Document No',
                   'Doc. Date', 'Pstng Date', "Cost Center", 'CostctrName', 'G/L', 'G/L Name',
                   'Profit Ctr', 'GR/IC Reference', 'Org.unit', 'Status', 'File 1', 'File 2', 'File 3',
                   'Created by', 'Time', 'Updated at', 'Reason for Rejection', 'Verified by', 'Verified at',
                   'Reference document', 'Reference invoice', 'Adv.doc year',
                   'Request no (Advance mulitple selection)', 'Invoice Reference Number',
                   'HOG Approval by', 'HOG Approval at', 'HOG Approval Req', 'Requested HOG ID',
                   'Month', 'Vesselcode', 'PEA Number', 'Status of Request', 'Clearing doc no.',
                   'Amount', 'On', 'Updated on', 'Verified on',
                   'HOG Approval on', 'Clearing date']]
        dfe.reset_index(drop=True, inplace=True)
        dfe.index += 1  # Start index from 1
        excel_buffer = BytesIO()
        dfe.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
        # Convert Excel buffer to base64
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
        # Download link for Excel file within a Markdown
        download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Created_verified on Holidays.xlsx">Download Excel file</a>'
        st.markdown(download_link, unsafe_allow_html=True)

