import streamlit as st
from analyze_excel import process_data
import re
import base64
from io import BytesIO
import pandas as pd
import plotly.express as px
from streamlit_dynamic_filters import DynamicFilters

@st.cache_resource(show_spinner=False)
def on_upload_click():
    """
    Callback function for the upload button click event.
    """
    st.session_state.upload = True
@st.cache_resource(show_spinner=False)
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
st.markdown("<h1 style='text-align: center; color: black;'><b>NON PO PAYMENT ANALYSIS </b></h1>", unsafe_allow_html=True)
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
    data = grouped_data.copy()
    exceptions = grouped_data.copy()
    t1, t2,t3 = st.tabs(["Overall Analysis","Yearly Analysis","Exceptions"])
    with t2:
        col1, col2, col3,col4 = st.columns(4)  # Split the page into two columns
        grouped_data = grouped_data.copy()
        selected_year = col1.selectbox('Select Year', grouped_data['year'].unique())
        filtered_data = grouped_data[grouped_data['year'] == selected_year]
        filtered_data['Cost Ctr'] = filtered_data['Cost Ctr'].astype(str)
        filtered_data['G/L'] = filtered_data['G/L'].astype(str)
        dynamic_filters = DynamicFilters(filtered_data, filters=['Cost Ctr', 'G/L'])
        dynamic_filters.display_filters()
        filtered_data = dynamic_filters.filter_df()
        filtered_data.reset_index(drop=True, inplace=True)
        filtered_data.index = filtered_data.index + 1
        filtered_data.rename_axis('S.NO', axis=1, inplace=True)
        df23 = filtered_data.copy()

        # Define card styles
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

        # Layout the cards
        c1, card1, middle_column, card2, c2 = st.columns([1, 4, 1, 4, 1])
        with card1:
            Total_Amount_Alloted = filtered_data['Amount'].sum() # Assuming filtered_data is defined

            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'>Total Amount Spent </h3>",
                unsafe_allow_html=True
            )
            Total_alloted = Total_Amount_Alloted
            st.markdown(
                f"<div style='{card1_style}'>"
                f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f}</h2>"
                "</div>",
                unsafe_allow_html=True
            )
            # st.write("")
            amount_length = filtered_data.groupby('year')['Amount'].sum().reset_index(name='amount_length')


            st.write(
                "<h2 style='text-align: center; font-size: 25px; font-weight: bold; color: black;'>Amount Spent -Category</h2>",
                unsafe_allow_html=True)
            category_amount = filtered_data.groupby('category')['Amount'].sum().reset_index()
            category_amount['Amount'] = category_amount['Amount'] / 10000000
            fig = px.bar(category_amount, x='category', y='Amount', color='category',
                         labels={'Amount': 'Amount (in Crores)'}, title='Amount Spent In Crores by Category',
                         width=400, height=525, template='plotly_white')

            # Add text labels on top of the bars with formatting in crores
            fig.update_traces(texttemplate='%{y:.2f} Cr', textposition='outside')

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
                xaxis_title='Category',
                yaxis_title='Amount (in Crores)',
                font=dict(size=14, color='black'),
                showlegend=False,
                bargap=bargap_value
            )

            # Show the plot in Streamlit
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
                yaxis_title='Amount (in Crores)',
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

                filtered_data = filtered_data.sort_values(by='Cumulative_Alloted/Cost Ctr/Year', ascending=False)[
                    ['Cost Ctr', 'CostctrName',
                     'Cumulative_Alloted/Cost Ctr/Year',
                     'Percentage_Cumulative_Alloted/Cost Ctr/Year']].drop_duplicates(subset=['Cost Ctr'], keep='first')
                filtered_data.rename(columns={'Cost Ctr':'Cost Center','CostctrName':'Name','Cumulative_Alloted/Cost Ctr/Year': 'Value (In ₹)',
                                              'Percentage_Cumulative_Alloted/Cost Ctr/Year': '% total'},
                                     inplace=True)
                filtered_data.reset_index(drop=True, inplace=True)
                filtered_data.index = filtered_data.index + 1
                filtered_data.rename_axis('S.NO', axis=1, inplace=True)

                filtered_transactions = filtered_transactions.sort_values(by='percentage Transcation/costctr/year',
                                                                          ascending=False)[['Cost Ctr', 'CostctrName',
                                                                                            'Cumulative_transactions/Cost Ctr/Year',
                                                                                            'percentage Transcation/costctr/year']].drop_duplicates(
                    subset=['Cost Ctr'], keep='first')
                filtered_transactions.rename(columns={'Cost Ctr':'Cost Center','CostctrName':'Name',
                                                      'Cumulative_transactions/Cost Ctr/Year': 'Transactions',
                                                      'percentage Transcation/costctr/year': '% total'},
                                             inplace=True)
                filtered_transactions.reset_index(drop=True, inplace=True)
                filtered_transactions.index = filtered_transactions.index + 1
                filtered_transactions.rename_axis('S.NO', axis=1, inplace=True)
                merged_df = pd.merge(filtered_data, filtered_transactions, on=['Cost Center','Name'])


            elif selected_category == f'Top 25 {G_L} Transactions':

                filtered_data['Cumulative_Alloted/G/L'] = filtered_data.groupby(['G/L'])['Amount'].transform('sum')
                filtered_data['Cumulative_Alloted/G/L/Year'] = filtered_data.groupby(['G/L', 'year'])['Amount'].transform('sum')
                filtered_data['Percentage_Cumulative_Alloted/G/L'] = (filtered_data['Cumulative_Alloted/G/L'] / filtered_data[
                    'cumulative_Alloted_Amount']) * 100
                yearly_total2 = filtered_data.groupby('year')['Amount'].sum().reset_index()
                yearly_total2.rename(columns={'Amount': 'Total_Alloted_Amount/year'}, inplace=True)
                filtered_data['Percentage_Cumulative_Alloted/G/L/Year'] = (filtered_data['Cumulative_Alloted/G/L/Year'] / filtered_data[
                    'Total_Alloted_Amount/year']) * 100
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
        dynamic_filters = DynamicFilters(filtered_df, filters=['Cost Ctr', 'G/L'])
        dynamic_filters.display_filters()
        filtered_data = dynamic_filters.filter_df()
        # filtered_data = filtered_df.copy()
        card1, card2 = st.columns(2)
        with card1:
            Total_Amount_Alloted = filtered_data['Amount'].sum() # Assuming filtered_data is defined
            # Calculate percentage of yearly allotted amount per category
            st.markdown(
                f"<h3 style='text-align: center; font-size: 25px;'>Total Amount Spent </h3>",
                unsafe_allow_html=True
            )
            Total_alloted = Total_Amount_Alloted
            st.markdown(
                f"<div style='{card1_style}'>"
                f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f}</h2>"
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


        @st.cache_resource(show_spinner=False)
        def line_plot_overall_transactions(data, category, years, width=500, height=300):
            filtered_data = data[(data['category'] == category) & (data['year'].isin(years))]
            data_length = filtered_data.groupby('year').size().reset_index(name='data_len')

            # Create the line plot
            fig = px.line(data_length, x='year', y='data_len',
                          title='Transactions Count by Year',
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
                    yshift=5
                )

            return fig


        @st.cache_resource(show_spinner=False)
        def line_plot_used_amount(data, category, years, width=500, height=300):
            data = data[(data['category'] == category) & (data['year'].isin(years))]

            amount_length = data.groupby('year')['Amount'].sum().reset_index(name='amount_length')

            # Create the line plot
            fig = px.line(amount_length, x='year', y='amount_length',
                          title='Transactions Count by Year',
                          labels={'year': 'Year', 'amount_length': 'Transaction Amount'},
                          markers=True)

            # Set mode to 'lines+markers'
            fig.update_traces(mode='lines+markers')

            # Update layout with width and height
            fig.update_layout(width=width, height=height)

            # Modify x-axis labels to include hyphens between years
            fig.update_xaxes(tickvals=years,
                             ticktext=[re.sub(r'(\d{4})(\d{2})', r'\1-\2', str(year)) for year in years])

            # Add titles on top of data points (annotations)
            for year, amount in zip(amount_length['year'], amount_length['amount_length']):
                fig.add_annotation(
                    x=year,
                    y=amount,
                    text=f"{amount:.2f}",
                    showarrow=False,
                    font=dict(size=12, color='black'),
                    align='center',
                    yshift=5
                )

            return fig
        years = (data['year'].unique())
        years_df = pd.DataFrame({'year': data['year'].unique()})
        filtered_years = (data['year'].unique())
        filtered_data = (data[data['year'].isin(filtered_years)].copy())



        st.write("## Employee Reimbursement Trend")


        col1, col2= st.columns(2)
        fig_employee = line_plot_overall_transactions(data, "Employee",years_df['year'])
        fig_employee2 = line_plot_used_amount(data, "Employee", years)
        col2.plotly_chart(fig_employee)
        col1.plotly_chart(fig_employee2)

        # Line plot for category "Korean Expats"
        st.write("## Korean Expats Reimbursement Trend")
        co1, co2= st.columns(2)
        fig_korean_expats = line_plot_overall_transactions(data, "Korean Expats", years_df['year'])
        fig_korean_expats1 = line_plot_used_amount(data, "Korean Expats", years)
        co2.plotly_chart(fig_korean_expats)
        co1.plotly_chart(fig_korean_expats1)

        # Line plot for category "Vendor"
        st.write("## Vendor Payment Trend")
        c1,c2=st.columns(2)
        fig_vendor = line_plot_overall_transactions(data, "Vendor",years_df['year'])
        fig_vendor1 = line_plot_used_amount(data, "Vendor", years)
        c2.plotly_chart(fig_vendor)
        c1.plotly_chart(fig_vendor1)

        # Line plot for category "Others"
        st.write("## Other Payment Trend")
        cl1,cl2=st.columns(2)
        fig_others = line_plot_overall_transactions(data, "Others",years_df['year'])
        fig_others1 = line_plot_used_amount(data, "Others", years)
        cl2.plotly_chart(fig_others)
        cl1.plotly_chart(fig_others1)
    with t3:
        # Define the functions
        def display_duplicate_invoices(exceptions):
            dfe = exceptions.copy()
            dfe = dfe[
                ['Payable req.no', 'Type', 'Vendor', 'Vendor Name', 'Invoice Number',"Text", "Cost Ctr", 'G/L', 'Document No',
                 'Doc. Date', 'Pstng Date', 'Amount','year']]

            dfe['Document No'] = dfe['Document No'].astype(str)
            dfe['Document No'] = dfe['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
            dfe['Document No'] = dfe['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
            dfe.rename(columns={'Vendor Name': 'Name', 'Type': "Doc.Type", 'Vendor': "ID", "Cost Ctr": "Cost Center"},
                       inplace=True)
            c1, c2, c3, c4,c5 = st.columns(5)
            options = ["All"] + [yr for yr in dfe['year'].unique() if yr != "All"]
            selected_option = c1.selectbox("Select an Year", options, index=0)
            # Filter the DataFrame based on the selected option
            if selected_option == "All":
                dfe = dfe  # Return the entire DataFrame
            else:
                dfe = dfe[dfe['year'] == selected_option]
            # Sorting by 'Invoice Number'
            dfe.sort_values(by='Invoice Number', ascending=True, inplace=True)
            dfe[['Name','Amount','Doc. Date','Invoice Number',"ID"]] = dfe[['Name','Amount','Doc. Date','Invoice Number',"ID"]].astype(str)
            selected_option = c2.selectbox("Select Duplicate type", ["Duplicate Invoice","Same amount,id,date", "Similar amount,costctr,invoice"], index=0)
            # Filter the DataFrame based on the selected option
            if selected_option == "Same amount,id,date":
                dfe = dfe[dfe.duplicated(subset=['Doc. Date', 'ID', 'Amount'], keep=False)]
            elif selected_option == "Similar amount,costctr,invoice":
                grouped_df = dfe.groupby(['Cost Center', 'Amount']).filter(lambda group: len(group) > 1)

                # Step 4: Compare values in column 'Invoice Number'
                def compare_strings(s1, s2):
                    common_chars = set(s1) & set(s2)
                    return len(common_chars) >= 0.7 * min(len(s1), len(s2))

                dfe = grouped_df[grouped_df.apply(
                    lambda row: compare_strings(row['Invoice Number'], grouped_df['Invoice Number'].iloc[0]), axis=1)]
                dfe = dfe.groupby(['Cost Center', 'Amount']).filter(lambda group: len(group) > 1)

                dfe.reset_index(drop=True, inplace=True)
            elif selected_option == "Duplicate Invoice":
                dfe = dfe[dfe.duplicated(subset=['Invoice Number'], keep=False)]
                dfe.sort_values(by='Invoice Number', ascending=True, inplace=True)
            dfe.reset_index(drop=True, inplace=True)
            dfe.index += 1  # Start index from 1
            dfe.sort_values(by='Invoice Number', ascending=True, inplace=True)
            dfe['Amount'] = pd.to_numeric(dfe['Amount'], errors='coerce')
            dfe['Amount'] = dfe['Amount'].astype(int)
            st.write(
                "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>Entries with Duplicate Invoices</h2>",
                unsafe_allow_html=True)
            st.write("")
            if dfe.empty:
                st.write("<div style='text-align: center; font-weight: bold; color: black;'>No entries with same Invoice Number</div>", unsafe_allow_html=True)
            else:
                 c1, card1, middle_column, card2, c2 = st.columns([1, 4, 1, 4, 1])
                 with card1:
                     Total_Amount_Alloted = dfe['Amount'].sum()
                     # Calculate percentage of yearly allotted amount per category

                     st.markdown(
                         f"<h3 style='text-align: center; font-size: 25px;'>Amount of Exposure (in Crores)</h3>",
                         unsafe_allow_html=True
                     )
                     Total_alloted = Total_Amount_Alloted / 10000000
                     st.markdown(
                         f"<div style='{card1_style}'>"
                         f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f}crores</h2>"
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
                 dfe = dfe.sort_values(by=['Name','Amount','Doc. Date'])
                 dfe.sort_values(by='Invoice Number', ascending=True, inplace=True)
                 dfe.reset_index(drop=True, inplace=True)
                 dfe.index += 1  # Start index from 1
                 c2.write(dfe)
                 excel_buffer = BytesIO()
                 dfe.to_excel(excel_buffer, index=False)
                 excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
            # Convert Excel buffer to base64
                 excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
            # Download link for Excel file within a Markdown
                 download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Duplicate Invoice.xlsx">Download Excel file</a>'
                 st.markdown(download_link, unsafe_allow_html=True)
        def same_approval(exceptions):
            dfe = exceptions.copy()
            dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
            dfe = dfe[dfe['Created'] == dfe['HOG Approval by']]
            dfe = dfe[
                ['Payable req.no', 'Type', 'Vendor', 'Vendor Name', 'Invoice Number','text', "Cost Ctr", 'G/L', 'Document No',
                 'Doc. Date', 'Pstng Date', 'Amount', 'Created', 'Verified by', 'HOG Approval by','year']]
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
                "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>Entries with Same Creator ID and HOG Approval ID</h2>",unsafe_allow_html=True)
            st.write("")
            if dfe.empty:
                st.write("<div style='text-align: center; font-weight: bold; color: black;'>No entries with same Creator ID and HOG Approval ID</div>", unsafe_allow_html=True)
            else:
                c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
                with card1:
                    Total_Amount_Alloted = dfe['Amount'].sum()
                    # Calculate percentage of yearly allotted amount per category

                    st.markdown(
                        f"<h3 style='text-align: center; font-size: 25px;'>Amount of Exposure(in Rupees)</h3>",
                        unsafe_allow_html=True
                    )
                    Total_alloted = Total_Amount_Alloted
                    st.markdown(
                        f"<div style='{card1_style}'>"
                        f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f} </h2>"
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
                c1, c2, c3 = st.columns([1, 8, 1])
                # Display the DataFrame
                dfe['Amount'] = dfe['Amount'].round()
                dfe = dfe.drop(columns=['year'])
                c2.write(dfe)
                excel_buffer = BytesIO()
                dfe.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
            # Convert Excel buffer to base64
                excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
            # Download link for Excel file within a Markdown
                download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Same CreatoR_HOG Approval ID.xlsx">Download Excel file</a>'
                st.markdown(download_link, unsafe_allow_html=True)
        def same_Creator_Verified_HOG(exceptions):
            dfe = exceptions.copy()
            dfe = dfe[(dfe['Created'] == dfe['Verified by']) & (dfe['Verified by'] == dfe['HOG Approval by'])]
            dfe = dfe[
                ['Payable req.no', 'Type', 'Vendor', 'Vendor Name', 'Invoice Number','Text',"Cost Ctr", 'G/L', 'Document No',
                 'Doc. Date', 'Pstng Date', 'Amount', 'Created', 'Verified by', 'HOG Approval by','year']]
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
                "<h2 style='text-align: center; font-size: 35px; font-weight: bold; color: black;'>Entries with same Creator ID , Verified ID and HOG Approval</h2>",unsafe_allow_html=True)
            st.write("")
            if dfe.empty:
                st.write("<div style='text-align: center; font-weight: bold; color: black;'>No entries with same Creator ID, Verified ID and HOG Approval ID</div>", unsafe_allow_html=True)
            else:
                c111, card1, middle_column, card2, c222 = st.columns([1, 4, 1, 4, 1])
                with card1:
                    Total_Amount_Alloted = dfe['Amount'].sum()
                    # Calculate percentage of yearly allotted amount per category

                    st.markdown(
                        f"<h3 style='text-align: center; font-size: 25px;'>Amount of Exposure(in Rupees)</h3>",
                        unsafe_allow_html=True
                    )
                    Total_alloted = Total_Amount_Alloted
                    st.markdown(
                        f"<div style='{card1_style}'>"
                        f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f} </h2>"
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
                c11,c22,c33 = st.columns([1, 8, 1])
                # Display the DataFrame
                dfe = dfe.drop(columns=['year'])
                dfe['Amount'] = dfe['Amount'].round()
                c22.write(dfe)
                excel_buffer = BytesIO()
                dfe.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
            # Convert Excel buffer to base64
                excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
            # Download link for Excel file within a Markdown
                download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="same_Creator_Verified_HOG Approval ID.xlsx">Download Excel file</a>'
                st.markdown(download_link, unsafe_allow_html=True)

        def same_Creator_Verified_HOGno(exceptions):
                dfe = exceptions.copy()
                dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
                dfe = dfe[
                          (dfe['Created'] == dfe['Verified by']) & (dfe['HOG Approval by'].isna())]
                dfe = dfe[
                    ['Payable req.no', 'Type', 'Vendor', 'Vendor Name', 'Invoice Number','Text', "Cost Ctr", 'G/L',
                     'Document No',
                     'Doc. Date', 'Pstng Date', 'Amount', 'Created', 'Verified by', 'HOG Approval by','year']]
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
                        # Calculate percentage of yearly allotted amount per category

                        st.markdown(
                            f"<h3 style='text-align: center; font-size: 25px;'>Amount of Exposure(in Rupees)</h3>",
                            unsafe_allow_html=True
                        )
                        Total_alloted = Total_Amount_Alloted
                        st.markdown(
                            f"<div style='{card1_style}'>"
                            f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f} </h2>"
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
                    c22.write(dfe)
                    excel_buffer = BytesIO()
                    dfe.to_excel(excel_buffer, index=False)
                    excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
                    # Convert Excel buffer to base64
                    excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
                    # Download link for Excel file within a Markdown
                    download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="same_Creator_Verified_NoHOG.xlsx">Download Excel file</a>'
                    st.markdown(download_link, unsafe_allow_html=True)

        def Creator_Verified_HOGno(exceptions):
                dfe = exceptions.copy()
                dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
                dfe = dfe[(~dfe['Created'].isna()) & (~dfe['Verified by'].isna()) & dfe['HOG Approval by'].isna()]
                dfe = dfe[
                    ['Payable req.no', 'Type', 'Vendor', 'Vendor Name', 'Invoice Number','Text', "Cost Ctr", 'G/L',
                     'Document No',
                     'Doc. Date', 'Pstng Date', 'Amount', 'Created', 'Verified by', 'HOG Approval by','year']]
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
                        # Calculate percentage of yearly allotted amount per category
                        st.markdown(
                            f"<h3 style='text-align: center; font-size: 25px;'>Amount of Exposure(in Rupees)</h3>",
                            unsafe_allow_html=True
                        )
                        Total_alloted = Total_Amount_Alloted
                        st.markdown(
                            f"<div style='{card1_style}'>"
                            f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f} </h2>"
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
                    c22.write(dfe)
                    excel_buffer = BytesIO()
                    dfe.to_excel(excel_buffer, index=False)
                    excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
                    # Convert Excel buffer to base64
                    excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
                    # Download link for Excel file within a Markdown
                    download_link = f'<a href="data:file/xls;base64,{excel_b64}" download=" Without HOG Approval.xlsx">Download Excel file</a>'
                    st.markdown(download_link, unsafe_allow_html=True)
        def Creator_HOG(exceptions):
                dfe = exceptions.copy()
                dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
                dfe = dfe[
                          (dfe['Created'] == dfe['HOG Approval by'])]

                dfe = dfe[
                    ['Payable req.no', 'Type', 'Vendor', 'Vendor Name', 'Invoice Number','Text', "Cost Ctr", 'G/L',
                     'Document No',
                     'Doc. Date', 'Pstng Date', 'Amount', 'Created', 'Verified by', 'HOG Approval by','year']]
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
                        # Calculate percentage of yearly allotted amount per category

                        st.markdown(
                            f"<h3 style='text-align: center; font-size: 25px;'>Amount of Exposure(in Rupees)</h3>",
                            unsafe_allow_html=True
                        )
                        Total_alloted = Total_Amount_Alloted
                        st.markdown(
                            f"<div style='{card1_style}'>"
                            f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f} </h2>"
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
                    c22.write(dfe)
                    excel_buffer = BytesIO()
                    dfe.to_excel(excel_buffer, index=False)
                    excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
                    # Convert Excel buffer to base64
                    excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
                    # Download link for Excel file within a Markdown
                    download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Same CreatorID and HOG ApprovalID.xlsx">Download Excel file</a>'
                    st.markdown(download_link, unsafe_allow_html=True)
        def same_Creator_Verified(exceptions):
                dfe = exceptions.copy()
                dfe = dfe.drop_duplicates(subset=['Vendor', 'year'], keep='first')
                dfe = dfe[
                          (dfe['Created'] == dfe['Verified by'])]
                dfe = dfe[
                    ['Payable req.no', 'Type', 'Vendor', 'Vendor Name', 'Invoice Number','Text', "Cost Ctr", 'G/L',
                     'Document No','year',
                     'Doc. Date', 'Pstng Date', 'Amount', 'Created', 'Verified by', 'HOG Approval by']]
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
                        # Calculate percentage of yearly allotted amount per category

                        st.markdown(
                            f"<h3 style='text-align: center; font-size: 25px;'>Amount of Exposure(in Rupees)</h3>",
                            unsafe_allow_html=True
                        )
                        Total_alloted = Total_Amount_Alloted
                        st.markdown(
                            f"<div style='{card1_style}'>"
                            f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f} </h2>"
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
                    c22.write(dfe)
                    excel_buffer = BytesIO()
                    dfe.to_excel(excel_buffer, index=False)
                    excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
                    # Convert Excel buffer to base64
                    excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
                    # Download link for Excel file within a Markdown
                    download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="same Creator ID And Verified ID.xlsx">Download Excel file</a>'
                    st.markdown(download_link, unsafe_allow_html=True)


        def all_sundays(year):
            return pd.date_range(start=str(year), end=str(year + 1), freq='W-SUN')


        # Create an empty DataFrame
        df = pd.DataFrame(columns=['date', 'type'])

        # Loop through years from 2020 to 2025
        for year in range(2020, 2026):
            sundays = all_sundays(year)
            sunday_df = pd.DataFrame({'date': sundays.strftime('%Y/%m/%d'), 'type': 'sunday'})

            # Add rows for specific dates and types
            gandhi_jayanti = pd.DataFrame({'date': [f'{year}/10/02'], 'type': ['gandhi_jayanti']})
            republic_day = pd.DataFrame({'date': [f'{year}/01/26'], 'type': ['republic_day']})
            independence_day = pd.DataFrame({'date': [f'{year}/08/15'], 'type': ['independence_day']})
            may_day = pd.DataFrame({'date': [f'{year}/05/01'], 'type': ['may_day']})

            df = pd.concat([df, sunday_df, gandhi_jayanti, republic_day, independence_day, may_day], ignore_index=True)


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

            dfe = dfe[
                    ['Payable req.no', 'Type', 'Vendor', 'Vendor Name', 'Invoice Number','Text', "Cost Ctr", 'G/L',
                     'Document No','Verified on',
                     'Doc. Date', 'Pstng Date', 'Amount', 'Created', 'Verified by', 'HOG Approval by','year']]
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
                    Total_Amount_Alloted = dfe['Amount'].sum()
                    # Calculate percentage of yearly allotted amount per category

                    st.markdown(
                        f"<h3 style='text-align: center; font-size: 25px;'>Amount of Exposure(in Rupees)</h3>",
                        unsafe_allow_html=True
                    )
                    Total_alloted = Total_Amount_Alloted
                    st.markdown(
                        f"<div style='{card1_style}'>"
                        f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f} </h2>"
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
                c22.write(dfe)
                excel_buffer = BytesIO()
                dfe.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
                # Convert Excel buffer to base64
                excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
                # Download link for Excel file within a Markdown
                download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Approval On Holidays.xlsx">Download Excel file</a>'
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

            dfe = dfe[
                    ['Payable req.no', 'Type', 'Vendor', 'Vendor Name', 'Invoice Number','Text', "Cost Ctr", 'G/L',
                     'Document No','Verified on',
                     'Doc. Date', 'Pstng Date', 'Amount', 'Created', 'Verified by', 'HOG Approval by','year']]
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
                    Total_Amount_Alloted = dfe['Amount'].sum()
                    # Calculate percentage of yearly allotted amount per category

                    st.markdown(
                        f"<h3 style='text-align: center; font-size: 25px;'>Amount of Exposure(in Rupees)</h3>",
                        unsafe_allow_html=True
                    )
                    Total_alloted = Total_Amount_Alloted
                    st.markdown(
                        f"<div style='{card1_style}'>"
                        f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>₹ {Total_alloted:,.2f} </h2>"
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
                c22.write(dfe)
                excel_buffer = BytesIO()
                dfe.to_excel(excel_buffer, index=False)
                excel_buffer.seek(0)  # Reset the buffer's position to the start for reading
                # Convert Excel buffer to base64
                excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
                # Download link for Excel file within a Markdown
                download_link = f'<a href="data:file/xls;base64,{excel_b64}" download="Posting&verified On Holidays.xlsx">Download Excel file</a>'
                st.markdown(download_link, unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        option = c1.selectbox("Select an option", ["Duplicate Invoices","No HOG Approval","Same Creator and Verified ID","Same Creator and HOG ID","Same Creator,Verified and HOG ID","Same Creator, Verified and NO HOG Approval","Posting on holidays","Posting and Verified During Holidays"],
                                      index=0)
        if option == "Duplicate Invoices":
            display_duplicate_invoices(exceptions)
        elif option == "Same Creator,Verified and HOG ID":
            same_Creator_Verified_HOG(exceptions)
        elif option == "Same Creator, Verified and NO HOG Approval":
            same_Creator_Verified_HOGno(exceptions)
        elif option == "No HOG Approval":
            Creator_Verified_HOGno(exceptions)
        elif option == "Same Creator and Verified ID":
            same_Creator_Verified(exceptions)
        elif option == "Same Creator and HOG ID":
            Creator_HOG(exceptions)
        elif option == "Posting on holidays":
            Approval_holidays(exceptions)
        elif option == "Posting and Verified During Holidays":
            Pstingverified_holidays(exceptions)
