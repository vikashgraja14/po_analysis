from streamlit_dynamic_filters import DynamicFilters
from functions import *
import streamlit.components.v1 as components
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


def on_upload_click():
    """
    Callback function for the upload button click event.
    """
    st.session_state.upload = True


def on_analyse_click():
    """
    Callback function for the analyse button click event.
    """
    st.session_state.analyse = True


# Initialize session state keys
if 'upload' not in st.session_state:
    st.session_state.upload = False
    st.session_state.upload_disabled = True

if 'analyse' not in st.session_state:
    st.session_state.analyse = False

st.set_page_config(layout="wide")
st.markdown("<h1 style='text-align: center; color: black;'><b>NON PO PAYMENT ANALYSIS </b></h1>",
            unsafe_allow_html=True)
with st.expander("Upload Excel Files", expanded=False):
    with st.form("my_form"):
        # File Uploader
        files = st.file_uploader("Active Assets", type="xlsx", accept_multiple_files=True, key="file_uploader")
        if not files:
            st.session_state.upload_disabled = True
            st.session_state.upload = False
        else:
            st.session_state.upload_disabled = False
        # Add the submit button
        submit_button = st.form_submit_button("UPLOAD")

# Automatically trigger upload when files are selected
if files:
    on_upload_click()

if st.session_state.upload:
    # Process only the first file from the list
    if len(files) == 1:
        grouped_data = process_data(files[0])
    else:
        # Process each file and concatenate the dataframes
        all_data = []
        for file in files:
            all_data.append(process_data(file))
        grouped_data = pd.concat(all_data)

    # Create copies of the data
    radio = grouped_data.copy()
    data = grouped_data.copy()


    exceptions = grouped_data.copy()
    exceptions2 = grouped_data.copy()
    t1, t2, t3, t4= st.tabs(["Overall Analysis", "Yearly Analysis", "Exceptions","Radio Button"])
    with t2:
        col1, col2, col3, col4 = st.columns(4)  # Split the page into two columns
        grouped_data = grouped_data.copy()
        selected_year = col1.selectbox('Select Year', grouped_data['year'].unique())
        filtered_data = grouped_data[grouped_data['year'] == selected_year]
        filtered_data['Cost Ctr'] = filtered_data['Cost Ctr'].astype(str)
        filtered_data['G/L'] = filtered_data['G/L'].astype(str)
        filtered_data['Vendor'] = filtered_data['Vendor'].astype(str)
        filtered_data['Cost Ctr'] = filtered_data['Cost Ctr'] + ' - ' + filtered_data[
            'CostctrName']
        filtered_data['G/L'] = filtered_data['G/L'] + ' - ' + filtered_data['G/L Name']
        filtered_data['Vendor'] = filtered_data['Vendor'] + ' - ' + filtered_data['Vendor Name']
        filtered_data['Cost Ctr'] = filtered_data['Cost Ctr'].astype(str)
        filtered_data['G/L'] = filtered_data['G/L'].astype(str)
        filtered_data['Vendor'] = filtered_data['Vendor'].astype(str)
        dynamic_filters = DynamicFilters(filtered_data, filters=['Cost Ctr', 'G/L', 'Vendor'])
        dynamic_filters.display_filters(location='columns', num_columns=4, gap='large')
        filtered_data = dynamic_filters.filter_df()
        filtered_data.reset_index(drop=True, inplace=True)
        filtered_data.index = filtered_data.index + 1
        filtered_data.rename_axis('S.NO', axis=1, inplace=True)
        for col in ['Cost Ctr', 'G/L', 'Vendor']:
            filtered_data[col] = filtered_data[col].str.split('-').str[0]
        df23 = filtered_data.copy()
        c1, card1, middle_column, card2, c2 = st.columns([1, 4, 1, 4, 1])
        with card1:
                Total_Amount_Alloted = filtered_data['Amount'].sum()

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
                    f"<h3 style='text-align: center; font-size: 25px;'>Total Amount Spent </h3>",
                    unsafe_allow_html=True
                )
                st.markdown(
                    f"<div style='{card1_style}'>"
                    f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>{amount_display}</h2>"
                    "</div>",
                    unsafe_allow_html=True
                )

                st.write(
                    "<h2 style='text-align: center; font-size: 25px; font-weight: bold; color: black;'>Amount Spent -Category</h2>",
                    unsafe_allow_html=True)

                # Assuming 'filtered_data' is a DataFrame that has been defined earlier
                category_amount = filtered_data.groupby('category')['Amount'].sum().reset_index()

                fig = px.bar(category_amount, x='category', y='Amount', color='category',
                             labels={'Amount': 'Amount (in Crores)'}, title='Amount Spent In Crores by Category',
                             width=400, height=525, template='plotly_white')


                def format_amount(amount):
                    integer_part = int(amount)
                    integer_length = len(str(integer_part))
                    if integer_length > 5 and integer_length <= 7:
                        return f"₹ {amount / 100000:,.2f} lks"
                    elif integer_length > 7:
                        return f"₹ {amount / 10000000:,.2f} crs"
                    else:
                        return f"₹ {amount:,.2f}"


                fig.update_traces(texttemplate='%{customdata}', textposition='outside')

                # Add a loop to iterate over each trace and update its 'customdata' with the correct amount
                for i, amount in enumerate(category_amount['Amount']):
                    fig.data[i].customdata = [format_amount(amount)]

                num_bars = len(category_amount)

                if num_bars == 1:
                    bargap_value = 0.8
                elif num_bars == 4:
                    bargap_value = 0.3
                elif num_bars == 2:
                    bargap_value = 0.7
                elif num_bars == 3:
                    bargap_value = 0.55
                else:
                    bargap_value = 0.55

                fig.update_layout(
                    xaxis_title='Category',
                    yaxis_title='Amount',
                    font=dict(size=14, color='black'),
                    showlegend=False,
                    bargap=bargap_value
                )

                st.plotly_chart(fig)

        with card2:
            Total_Transaction = df23['Amount'].count()  # Assuming filtered_data is defined
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'>Total Count Of Transactions</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"<div style='{card2_style}'>"
                f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
            st.write(
                "<h2 style='text-align: center; font-size: 25px; font-weight: bold; color: black;'>Transaction Count-Category</h2>",
                unsafe_allow_html=True)
            st.write("")
            category_amount = filtered_data.groupby('category')['Amount'].count().reset_index()

            fig = px.bar(category_amount, x='category', y='Amount', color='category',
                         labels={'Amount': 'Transactions'}, title='Transaction Count by Category',
                         width=400, height=500, template='plotly_white')

            # Add text annotations for each bar
            for trace in fig.data:
                for i, value in enumerate(trace.y):
                    fig.add_annotation(
                        x=trace.x[i],
                        y=value,
                        text=f"{value}",
                        showarrow=True,
                        font=dict(size=12, color='black'),
                        align='center',
                        yshift=5
                    )

            # Determine the number of bars
            num_bars = len(category_amount)

            # Set the bargap based on the number of bars
            if num_bars == 1:
                bargap_value = 0.8
            elif num_bars == 4:
                bargap_value = 0.3
            elif num_bars == 2:
                bargap_value = 0.7
            elif num_bars == 3:
                bargap_value = 0.55
            else:
                # Default value (you can adjust this as needed)
                bargap_value = 0.55

            # Customize the layout
            fig.update_layout(
                annotations=[dict(showarrow=False)],
                margin=dict(l=20, r=20, t=40, b=20),
                showlegend=False,
                xaxis_title='Category',
                yaxis_title='Transactions',
                font=dict(size=14, color='black'),
                bargap=bargap_value
            )
            # Show the plot in Streamlit
            st.plotly_chart(fig)
            st.write("")
        unique_categories = filtered_data['category'].unique()
        unique_categories = sorted(unique_categories, key=len, reverse=True)
        preferred_order = ['Vendor', 'Employee', 'Korean Expats', 'Others']

        # Sort categories according to preferred order
        unique_categories = sorted(unique_categories,
                                   key=lambda x: preferred_order.index(x) if x in preferred_order else len(
                                       preferred_order))

        # Define Cost_Ctr and G_L names
        Cost_Ctr = 'Cost Ctr'
        G_L = 'G/L'

        # Add 'All' option to the unique categories list
        unique_categories_with_options = ['All'] + [f'Top 25 {category} Transactions' for category in
                                                    unique_categories] + [f'Top 25 {Cost_Ctr} Transactions',
                                                                          f'Top 25 {G_L} Transactions']
        col1, col2, col3 = st.columns(3)
        # Selectbox to choose category or other options
        selected_category = col1.selectbox("Select Category or Option:", unique_categories_with_options)

        # Get all unique values in the 'Cost Ctr' column
        all_cost_ctrs = filtered_data['Cost Ctr'].unique()

        # Set selected_cost_ctr to represent all values in the 'Cost Ctr' column
        selected_cost_ctr = 'All'  # or all_cost_ctrs if you want the default to be all values

        # Filter the data based on the selected year and dropdown selections
        filtered_data = filtered_data[filtered_data['year'] == selected_year]
        # filtered_data = filtered_data.drop_duplicates(subset=['Vendor', 'year'], keep='first')
        df = filtered_data[filtered_data['year'] == selected_year]
        filtered_transactions = filtered_data[filtered_data['year'] == selected_year]
        # filtered_transactions = filtered_transactions.drop_duplicates(subset=['Vendor', 'year'], keep='first')

        if selected_category != 'All':
            if selected_category == f'Top 25 {Cost_Ctr} Transactions':
                filtered_data['overall_Alloted_Amount'] = filtered_data.groupby(['year'])['Amount'].transform('sum')
                filtered_data['Cumulative_Alloted/Cost Ctr'] = filtered_data.groupby(['Cost Ctr'])['Amount'].transform('sum')
                filtered_data['Cumulative_Alloted/Cost Ctr/Year'] = filtered_data.groupby(['Cost Ctr', 'year'])['Amount'].transform('sum')
                filtered_data['Percentage_Cumulative_Alloted/Cost Ctr'] = (filtered_data['Cumulative_Alloted/Cost Ctr/Year'] / filtered_data[
                    'overall_Alloted_Amount']) * 100
                yearly_total2 = filtered_data.groupby('year')['Amount'].sum().reset_index()
                yearly_total2.rename(columns={'Amount': 'Total_Alloted_Amount/year'}, inplace=True)
                filtered_data['Percentage_Cumulative_Alloted/Cost Ctr/Year'] = (filtered_data['Cumulative_Alloted/Cost Ctr/Year'] / filtered_data[
                    'overall_Alloted_Amount']) * 100

                filtered_data = filtered_data.sort_values(by='Cumulative_Alloted/Cost Ctr/Year', ascending=False)[
                    ['Cost Ctr', 'CostctrName',
                     'Cumulative_Alloted/Cost Ctr/Year',
                     'Percentage_Cumulative_Alloted/Cost Ctr/Year']].drop_duplicates(subset=['Cost Ctr'], keep='first')
                filtered_data.rename(columns={'Cumulative_Alloted/Cost Ctr/Year': 'Value (In ₹)','CostctrName':'Name','Cost Ctr':'Cost Center',
                                              'Percentage_Cumulative_Alloted/Cost Ctr/Year': '%total'},
                                     inplace=True)
                filtered_data.reset_index(drop=True, inplace=True)
                filtered_data.index = filtered_data.index + 1
                filtered_data.rename_axis('S.NO', axis=1, inplace=True)
                filtered_transactions['Cummulative_transactions2'] = len(filtered_transactions)
                filtered_transactions['Cumulative_transactions/Cost Ctr'] = filtered_transactions.groupby(['Cost Ctr'])['Cost Ctr'].transform('count')
                filtered_transactions['Cumulative_transactions/Cost Ctr/Year'] = filtered_transactions.groupby(['Cost Ctr'])['Cost Ctr'].transform('count')
                filtered_transactions['percentage Transcation/Cost Ctr/year'] = filtered_transactions['Cumulative_transactions/Cost Ctr/Year'] / filtered_transactions[
                    'Cummulative_transactions2'] * 100
                filtered_transactions = filtered_transactions.sort_values(by='percentage Transcation/Cost Ctr/year',
                                                                          ascending=False)[['Cost Ctr', 'CostctrName',
                                                                                            'Cumulative_transactions/Cost Ctr/Year',
                                                                                            'percentage Transcation/Cost Ctr/year'
                                                                                            ]].drop_duplicates(
                    subset=['Cost Ctr'], keep='first')
                filtered_transactions.rename(columns={'Cost Ctr':'Cost Center','CostctrName':'Name',
                                                      'Cumulative_transactions/Cost Ctr/Year': 'Transactions',
                                                      'percentage Transcation/Cost Ctr/year': '% total'},
                                             inplace=True)
                filtered_transactions.reset_index(drop=True, inplace=True)
                filtered_transactions.index = filtered_transactions.index + 1
                filtered_data.rename_axis('S.NO', axis=1, inplace=True)
                filtered_transactions.reset_index(drop=True, inplace=True)
                filtered_transactions.index = filtered_transactions.index + 1
                filtered_transactions.rename_axis('S.NO', axis=1, inplace=True)
                merged_df = pd.merge(filtered_data, filtered_transactions, on=['Cost Center','Name'])

            elif selected_category == f'Top 25 {G_L} Transactions':
                filtered_data['overall_Alloted_Amount'] = filtered_data.groupby(['year'])['Amount'].transform('sum')
                filtered_data['Cumulative_Alloted/G/L'] = filtered_data.groupby(['G/L'])['Amount'].transform('sum')
                filtered_data['Cumulative_Alloted/G/L/Year'] = filtered_data.groupby(['G/L', 'year'])['Amount'].transform('sum')
                filtered_data['Percentage_Cumulative_Alloted/G/L'] = (filtered_data['Cumulative_Alloted/G/L/Year'] / filtered_data[
                    'overall_Alloted_Amount']) * 100
                yearly_total2 = filtered_data.groupby('year')['Amount'].sum().reset_index()
                yearly_total2.rename(columns={'Amount': 'Total_Alloted_Amount/year'}, inplace=True)
                filtered_data['Percentage_Cumulative_Alloted/G/L/Year'] = (filtered_data['Cumulative_Alloted/G/L/Year'] / filtered_data[
                    'overall_Alloted_Amount']) * 100
                filtered_data = filtered_data.sort_values(by='Cumulative_Alloted/G/L/Year', ascending=False)[
                    ['G/L', 'G/L Name',
                     'Cumulative_Alloted/G/L/Year',
                     'Percentage_Cumulative_Alloted/G/L/Year']].drop_duplicates(subset=['G/L'], keep='first')
                filtered_data.rename(columns={'Cumulative_Alloted/G/L/Year': 'Value (In ₹)','G/LName':'Name',
                                              'Percentage_Cumulative_Alloted/G/L/Year': '%total'},
                                     inplace=True)
                filtered_data.reset_index(drop=True, inplace=True)
                filtered_data.index = filtered_data.index + 1
                filtered_data.rename_axis('S.NO', axis=1, inplace=True)
                filtered_transactions['Cummulative_transactions2'] = len(filtered_transactions)
                filtered_transactions['Cumulative_transactions/G/L'] = filtered_transactions.groupby(['G/L'])['G/L'].transform('count')
                filtered_transactions['Cumulative_transactions/G/L/Year'] = filtered_transactions.groupby(['G/L'])['G/L'].transform('count')
                filtered_transactions['percentage Transcation/G/L/year'] = filtered_transactions['Cumulative_transactions/G/L/Year'] / filtered_transactions[
                    'Cummulative_transactions2'] * 100
                filtered_transactions = filtered_transactions.sort_values(by='percentage Transcation/G/L/year',
                                                                          ascending=False)[['G/L', 'G/L Name',
                                                                                            'Cumulative_transactions/G/L/Year',
                                                                                            'percentage Transcation/G/L/year'
                                                                                            ]].drop_duplicates(
                    subset=['G/L'], keep='first')
                filtered_transactions.rename(columns={
                                                      'Cumulative_transactions/G/L/Year': 'Transactions',
                                                      'percentage Transcation/G/L/year': '% total'},
                                             inplace=True)
                filtered_transactions.reset_index(drop=True, inplace=True)
                filtered_transactions.index = filtered_transactions.index + 1
                filtered_transactions.rename_axis('S.NO', axis=1, inplace=True)
                merged_df = pd.merge(filtered_data, filtered_transactions, on=['G/L', 'G/L Name'])

            else:
                category = ' '.join(selected_category.split()[2:-1])
                filtered_data = filtered_data[filtered_data['category'] == category]
                filtered_transactions = filtered_data.copy()
                filtered_data['Amount_used/Year2'] = filtered_data.groupby(['Vendor', 'year', 'category'])['Amount'].transform('sum')
                filtered_data['Yearly_Alloted_Amount\Category2'] = filtered_data.groupby(['category', 'year'])['Amount'].transform('sum')
                filtered_data['percentage_of_amount/category_used/year2'] = (filtered_data['Amount_used/Year2'] / filtered_data[
                    'Yearly_Alloted_Amount\Category2']) * 100
                filtered_transactions = filtered_transactions[filtered_transactions['category'] == category]
                filtered_data = filtered_data.sort_values(by='percentage_of_amount/category_used/year2',
                                                          ascending=False)[['Vendor','Vendor Name', 'Amount_used/Year2',
                                                                            'percentage_of_amount/category_used/year2']]
                filtered_data = filtered_data.drop_duplicates(subset=['Vendor'], keep='first')
                filtered_data.rename(columns={'Amount_used/Year2': 'Value (In ₹)','Vendor':'ID','Vendor Name':'Name',
                                              'percentage_of_amount/category_used/year2': '%total'},
                                     inplace=True)

                filtered_data.reset_index(drop=True, inplace=True)
                filtered_data.index = filtered_data.index + 1
                filtered_data.rename_axis('S.NO', axis=1, inplace=True)
                filtered_transactions['Transations/year/Vendor2'] = filtered_transactions.groupby(['Vendor', 'year', 'category'])['Vendor'].transform('count')
                filtered_transactions['overall_transactions/category/year2'] = filtered_transactions.groupby(['category', 'year'])['category'].transform(
                    'count')
                filtered_transactions['percentransations_made/category/year2'] = (filtered_transactions['Transations/year/Vendor2'] / filtered_transactions[
                    'overall_transactions/category/year2']) * 100
                filtered_transactions = filtered_transactions.sort_values(by='percentransations_made/category/year2',
                                                                          ascending=False)[
                    ['Vendor','Vendor Name',
                     'Transations/year/Vendor2', 'percentransations_made/category/year2']]
                filtered_transactions = filtered_transactions.drop_duplicates(subset=['Vendor'], keep='first')
                filtered_transactions.rename(columns={'Vendor':'ID','Vendor Name':'Name',
                                                      'Transations/year/Vendor2': 'Transactions',
                                                      'percentransations_made/category/year2': '% total'},
                                             inplace=True)
                filtered_transactions.reset_index(drop=True, inplace=True)  # Reset index here
                filtered_transactions.index = filtered_transactions.index + 1
                filtered_transactions.rename_axis('S.NO', axis=1, inplace=True)
                merged_df = pd.merge(filtered_data, filtered_transactions, on=['ID', 'Name'])
        else:
            filtered_data['Amount_used/Year2'] = filtered_data.groupby(['Vendor', 'year'])[
                'Amount'].transform('sum')
            filtered_data['Yearly_Alloted_Amount\Category2'] = filtered_data.groupby(['year'])['Amount'].transform('sum')
            filtered_data['percentage_of_amount/category_used/year2'] = (filtered_data['Amount_used/Year2'] /
                                                                         filtered_data[
                                                                             'Yearly_Alloted_Amount\Category2']) * 100
            filtered_data = filtered_data.sort_values(by='percentage_of_amount/category_used/year2',
                                                      ascending=False)[['Vendor', 'Vendor Name', 'Amount_used/Year2',
                                                                        'percentage_of_amount/category_used/year2']]
            filtered_data = filtered_data.drop_duplicates(subset=['Vendor'], keep='first')
            filtered_data.rename(columns={'Amount_used/Year2': 'Value (In ₹)', 'Vendor': 'ID', 'Vendor Name': 'Name',
                                          'percentage_of_amount/category_used/year2': '%total'},
                                 inplace=True)

            filtered_data.reset_index(drop=True, inplace=True)
            filtered_data.index = filtered_data.index + 1
            filtered_data.rename_axis('S.NO', axis=1, inplace=True)
            filtered_transactions['Transations/year/Vendor2'] = \
            filtered_transactions.groupby(['Vendor', 'year'])['Vendor'].transform('count')
            filtered_transactions['overall_transactions/category/year2'] = \
            filtered_transactions['overall_transactions/category/year2'] = \
            filtered_transactions.groupby(['year'])['category'].transform(
                'count')
            filtered_transactions['percentransations_made/category/year2'] = (filtered_transactions[
                                                                                  'Transations/year/Vendor2'] /
                                                                              filtered_transactions[
                                                                                  'overall_transactions/category/year2']) * 100
            filtered_transactions = filtered_transactions.sort_values(by='percentransations_made/category/year2',
                                                                      ascending=False)[
                ['Vendor', 'Vendor Name',
                 'Transations/year/Vendor2', 'percentransations_made/category/year2']]
            filtered_transactions = filtered_transactions.drop_duplicates(subset=['Vendor'], keep='first')
            filtered_transactions.rename(columns={'Vendor':'ID','Vendor Name':'Name',
                                                  'Transations/year/Vendor2': 'Transactions',
                                                  'percentransations_made/category/year2': '% total'},
                                         inplace=True)
            filtered_transactions.reset_index(drop=True, inplace=True)  # Reset index here
            filtered_transactions.index = filtered_transactions.index + 1
            filtered_transactions.rename_axis('S.NO', axis=1, inplace=True)
            merged_df = pd.merge(filtered_data, filtered_transactions, on=['ID','Name'])
        col1, col2 = st.columns(2)

        with col1:

            st.write("Value wise (In \u20B9)")
            st.write(filtered_data.head(25))

        with col2:
            st.write("Transaction Count Wise")
            st.write(filtered_transactions.head(25))
        excel_buffer = BytesIO()
        merged_df.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)  # Reset the buffer's position to the start for reading

        # Convert Excel buffer to base64
        excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()

        # Download link for Excel file within a Markdown
        download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="PO Analysis.xlsx">Download Excel file</a>'
        st.markdown(download_link, unsafe_allow_html=True)
    with t1:
        filtered_df = data.copy()
        filtered_df['Cost Ctr'] = filtered_df['Cost Ctr'].astype(str)
        filtered_df['G/L'] = filtered_df['G/L'].astype(str)
        filtered_df['Vendor'] = filtered_df['Vendor'].astype(str)
        filtered_df['Cost Ctr'] = filtered_df['Cost Ctr'] + ' - ' + filtered_df[
            'CostctrName']
        filtered_df['G/L'] = filtered_df['G/L'] + ' - ' + filtered_df['G/L Name']
        filtered_df['Vendor'] = filtered_df['Vendor'] + ' - ' + filtered_df['Vendor Name']
        filtered_df['Cost Ctr'] = filtered_df['Cost Ctr'].astype(str)
        filtered_df['G/L'] = filtered_df['G/L'].astype(str)
        filtered_df['Vendor'] = filtered_df['Vendor'].astype(str)
        c1, c2, c3, c4 = st.columns(4)
        # Cost Center multiselect
        options_cost_center = ["All"] + [yr for yr in filtered_df['Cost Ctr'].unique() if yr != "All"]
        selected_cost_centers = c1.multiselect("Select Cost Centers", options_cost_center, default=["All"])

        # Filter the DataFrame based on selected Cost Centers
        if "All" in selected_cost_centers:
            filtered_df2 = filtered_df
        else:
            filtered_df2 = filtered_df[filtered_df['Cost Ctr'].isin(selected_cost_centers)]

        # G/L multiselect
        options_gl = ["All"] + [yr for yr in filtered_df2['G/L'].unique() if yr != "All"]
        selected_gl = c2.multiselect("Select G/L", options_gl, default=["All"])

        # Filter the DataFrame based on selected G/L
        if "All" in selected_gl:
            filtered_df1 = filtered_df2
        else:
            filtered_df1 = filtered_df2[filtered_df2['G/L'].isin(selected_gl)]

        # Vendor multiselect
        options_vendor = ["All"] + [yr for yr in filtered_df1['Vendor'].unique() if yr != "All"]
        selected_vendor = c3.multiselect("Select Vendor", options_vendor, default=["All"])

        # Filter the DataFrame based on selected Vendor
        if "All" in selected_vendor:
            filtered_data = filtered_df1
        else:
            filtered_data = filtered_df1[filtered_df1['Vendor'].isin(selected_vendor)]

        # dynamic_filters = DynamicFilters(filtered_df, filters=['Cost Ctr', 'G/L'])
        # dynaic_filters.display_filters()
        # filtered_data = dynammic_filters.filter_df()
        # # filtered_data = filtered_df.copy()
        filtered_data['Amount'] = filtered_data['Amount'].round()
        # df['G/L'] = df['G/L'].astype(str)
        # df['G/L'] = df['G/L'].apply(lambda x: str(x) if isinstance(x, str) else '')
        # df['G/L'] = df['G/L'].apply(lambda x: re.sub(r'\..*', '', x))
        filtered_data['Amount'] = filtered_data['Amount'].astype(int)
        card1, card2 = st.columns(2)
        for col in ['Cost Ctr', 'G/L', 'Vendor']:
            filtered_data[col] = filtered_data[col].str.split('-').str[0]
        with card1:
            Total_Amount_Alloted = filtered_data['Amount'].sum()

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
                f"<h3 style='text-align: center; font-size: 25px;'>Total Amount Spent </h3>",
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
            Total_Transaction = filtered_data['Payable req.no'].count()  # Assuming filtered_data is defined
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'>Total Count Of Transactions</h3>",
                unsafe_allow_html=True
            )
            st.markdown(
                f"<div style='{card2_style}'>"
                f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
        data = filtered_data.copy()
        data.reset_index(drop=True, inplace=True)
        data.index += 1  # Start index from 1

        years = (data['year'].unique())
        years_df = pd.DataFrame({'year': data['year'].unique()})
        filtered_years = (data['year'].unique())
        filtered_data = (data[data['year'].isin(filtered_years)].copy())

        st.write("## Employee Reimbursement Trend")

        c2,col1, col2,c5 = st.columns([1,4,4,1])
        c2.write("")
        c5.write("")
        fig_employee = line_plot_overall_transactions(data, "Employee",years_df['year'])
        fig_employee2 = line_plot_used_amount(data, "Employee", years)
        col2.plotly_chart(fig_employee)
        col1.plotly_chart(fig_employee2)

        # Line plot for category "Korean Expats"
        st.write("## Korean Expats Reimbursement Trend")
        c2,co1, co2,c5 = st.columns([1,4,4,1])
        fig_korean_expats = line_plot_overall_transactions(data, "Korean Expats", years_df['year'])
        fig_korean_expats1 = line_plot_used_amount(data, "Korean Expats", years)
        c2.write("")
        c5.write("")
        co2.plotly_chart(fig_korean_expats)
        co1.plotly_chart(fig_korean_expats1)

        # Line plot for category "Vendor"
        st.write("## Vendor Payment Trend")
        co2,c1, c2,c5 = st.columns([1,4,4,1])
        co2.write("")
        c5.write("")
        fig_vendor = line_plot_overall_transactions(data, "Vendor",years_df['year'])
        fig_vendor1 = line_plot_used_amount(data, "Vendor", years)
        c2.plotly_chart(fig_vendor)
        c1.plotly_chart(fig_vendor1)

        # Line plot for category "Others"
        st.write("## Other Payment Trend")
        c2,cl1, cl2,c5 = st.columns([1,4,4,1])
        fig_others = line_plot_overall_transactions(data, "Others",years_df['year'])
        fig_others1 = line_plot_used_amount(data, "Others", years)
        cl2.plotly_chart(fig_others)
        cl1.plotly_chart(fig_others1)
    with t3:
        # Define the functions
        c1, c42, c3, c4 = st.columns(4)
        option = c1.selectbox("Select an option",
                              ["Duplicate Exceptions", "No approver(HOD_HOG)","Creator_Verifier_SAME",
                               "Creator_Verified_SAME_NO Approver",
                               "Creator_Approver(HOD_HOD)_SAME", "Creator_Verifier_Approver(HOD_HOD)_SAME",
                                "Holiday Transactions", "Created_verified on Holidays"],
                              index=0)
        c42.write("")
        c3.write("")
        c4.write("")
        if option == "Duplicate Exceptions":
            display_duplicate_invoices(exceptions)
        elif option == "Creator_Verified_SAME_NO Approver":
            same_Creator_Verified_HOGno(exceptions)
        elif option == "No approver(HOD_HOG)":
            Creator_Verified_HOGno(exceptions)
        elif option == "Creator_Verifier_SAME":
            same_Creator_Verified(exceptions)
        elif option == "Creator_Approver(HOD_HOD)_SAME":
            Creator_HOG(exceptions)
        elif option == "Creator_Verifier_Approver(HOD_HOD)_SAME":
            same_Creator_Verified_HOG(exceptions)
        elif option == "Holiday Transactions":
            Approval_holidays(exceptions)
        elif option == "Created_verified on Holidays":
            Pstingverified_holidays(exceptions)
    with t4:
        t41, t42, t43, t44 = st.tabs(["Factor 1", "Factor 2", "Factor 3", "Factor 4"])

        with t41:
            def ChangeWidgetFontSize(wgt_txt, wch_font_size='12px'):
                htmlstr = """<script>var elements = window.parent.document.querySelectorAll('*'), i;
                                for (i = 0; i < elements.length; ++i) { if (elements[i].innerText == |wgt_txt|) 
                                    { elements[i].style.fontSize='""" + wch_font_size + """';} } </script>  """

                htmlstr = htmlstr.replace('|wgt_txt|', "'" + wgt_txt + "'")
                components.html(f"{htmlstr}", height=0, width=0)


            # Renaming columns in the DataFrame
            radio.rename(
                columns={
                    'Vendor Name': 'Name',
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
                           'Pstng Date', 'Creator', 'Verifier', 'HOG Approval by', 'category',
                           'Reference invoice', 'CostctrName', 'G/L Name', 'Profit Ctr',
                           'GR/IC Reference', 'Org.unit', 'Status', 'File 1', 'File 2', 'File 3',
                           'Time', 'Updated at', 'Reason for Rejection', 'Verified at',
                           'Reference document', 'Adv.doc year',
                           'Request no (Advance mulitple selection)', 'Invoice Reference Number',
                           'HOG Approval at', 'HOG Approval Req', 'Requested HOG ID',
                           'Month', 'Vesselcode', 'PEA Number', 'Status of Request', 'Clearing doc no.',
                           'On', 'Updated on', 'Verified on',
                           'HOG Approval on', 'Clearing date']]

            radio['Amount'] = radio['Amount'].round()


            columns_list = [
                'Reimbursement ID', 'Amount', "Document Date",
                "Cost Center", 'G/L', 'Invoice Number', 'Text'
            ]

            c11, c2, c3, c4 = st.columns(4)
            options = ["All"] + [yr for yr in radio['Year'].unique() if yr != "All"]
            selected_option = c11.selectbox("Choose an Year", options, index=0)
            # Filter the DataFrame based on the selected option
            if selected_option == "All":
                radio = radio  # Return the entire DataFrame


            else:
                radio = radio[radio['Year'] == selected_option]

            # Assuming 'radio' is your DataFrame

            # Function to filter DataFrame for duplicate values based on selected checkboxes

            def filter_dataframe(radio, columns_list):
                # Dictionary to hold the checkbox states
                checkbox_states = {}

                # Create columns in Streamlit with the same number as elements in columns_list
                cols = st.columns(len(columns_list))

                # Use a for loop to create a checkbox in each column and store the state
                for i, col in enumerate(cols):
                    with col:
                        # Create a checkbox and store its state
                        checkbox_states[columns_list[i]] = st.checkbox(f"{columns_list[i]}", key=i)

                # List to hold the names of columns to check for duplicates
                columns_to_check = [column_name for column_name, is_checked in checkbox_states.items() if is_checked]

                # Start with all rows
                filtered_df = radio.copy()

                # If any checkboxes are checked, filter the DataFrame for duplicates in those columns
                if columns_to_check:
                    # Filter the DataFrame for duplicates across the selected columns
                    filtered_df = filtered_df[filtered_df.duplicated(subset=columns_to_check, keep=False)]
                    filtered_df = filtered_df.sort_values(by=columns_to_check)
                    filtered_df.reset_index(drop=True, inplace=True)
                    filtered_df.index += 1
                return filtered_df, ', '.join(columns_to_check)


            filtered_df, columns_names = filter_dataframe(radio, columns_list)
            st.markdown(
                f"<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>ENTRIES with SAME {columns_names}</h2>",
                unsafe_allow_html=True)

            c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
            with card1:
                Total_Amount_Alloted = filtered_df['Amount'].sum()

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
                Total_Transaction = len(filtered_df)
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
            c11, c522, c33 = st.columns([1, 8, 1])
            # Display the DataFrame
            c522.write("")
            c522.write(filtered_df[['Payable req.no', 'Doc.Type', 'Reimbursement ID', 'Amount', 'Document Date', 'Name',
                                      'Invoice Number', 'Text', 'Cost Center', 'G/L', 'Document No',
                                      'Pstng Date', 'Creator', 'Verifier', 'HOG Approval by']])
            filename = f"SAME {columns_names}.xlsx"
            filtered_df.reset_index(drop=True, inplace=True)
            filtered_df.index += 1  # Start index from 1
            excel_buffer = BytesIO()
            filtered_df.to_excel(excel_buffer, index=False)
            excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
            # Convert Excel buffer to base64
            excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()

            download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="{filename}">Download Excel file</a>'
            st.markdown(download_link, unsafe_allow_html=True)


            # # Display the filtered DataFrame
            # st.dataframe(filtered_df)

            # col1, col2,c3,c4,c5,c6,c7 = st.columns(7)
            #
            # # Place each checkbox in one of the columns
            # with col1:
            #     coln_plus_1 = st.checkbox("70% SIMILAR INVOICE", key='coln_plus_1')
            #
            # with col2:
            #     coln_plus_2 = st.checkbox("Holiday Transactions", key='coln_plus_2')








