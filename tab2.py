from io import BytesIO
import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit_dynamic_filters import DynamicFilters
import base64

CARD1_STYLE = """
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

# Style for card2 used in the Streamlit application.
CARD2_STYLE = """
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
def format_amount(amount):
    """
    Format the given amount into a human-readable currency string.

    Args:
    - amount (float or int): The amount to be formatted.

    Returns:
    - str: A formatted string representing the amount in
    Indian Rupees (₹).
           The formatting depends on the length of the integer
           part of the amount:
           - If the integer length is greater than 5 and less
           than or equal to 7, the amount is divided by
           100,000 and formatted as "₹ {amount / 100000:,.2f} lks".
           - If the integer length is greater than 7, the amount
            is divided by 10,000,000 and formatted as
            "₹ {amount / 10000000:,.2f} crs".
           - Otherwise, the amount is formatted directly as
           "₹ {amount:,.2f}".

    """
    integer_part = int(amount)
    integer_length = len(str(integer_part))

    if 5 < integer_length <= 7:
        return f"₹ {amount / 100000:,.2f} lks"
    elif integer_length > 7:
        return f"₹ {amount / 10000000:,.2f} crs"
    else:
        return f"₹ {amount:,.2f}"

def display_dashboard(data):
    filtered_data = data.copy()
    filtered_data["Vendor"] = filtered_data["Vendor"].astype(str)
    filtered_data["Vendor Name"] = filtered_data["Vendor Name"].astype(str)
    filtered_data["Created"] = filtered_data["Created"].astype(str)
    vendor_name_dict = dict(
        zip(filtered_data["Vendor"], filtered_data["Vendor Name"])
    )
    # Map 'Vendor Name' to a new column 'CreatedName' based on 'Created'
    filtered_data["CreatedName"] = filtered_data["Created"].map(
        vendor_name_dict
    )
    # Fill NaN values in 'CreatedName' with an empty string
    filtered_data["CreatedName"] = filtered_data["CreatedName"].fillna("")
    # Concatenate 'Created' with 'CreatedName', separated by ' - ', only if
    # 'CreatedName' is not empty
    filtered_data["Created"] = filtered_data.apply(
        lambda row: (
            row["Created"] + " - " + row["CreatedName"]
            if row["CreatedName"]
            else row["Created"]
        ),
        axis=1,
    )
    # Drop the now unnecessary 'CreatedName' column
    filtered_data.drop("CreatedName", axis=1, inplace=True)
    filtered_data["Cost Ctr"] = filtered_data["Cost Ctr"].astype(str)
    filtered_data["G/L"] = filtered_data["G/L"].astype(str)
    filtered_data["Cost Ctr"] = (
            filtered_data["Cost Ctr"] + " - " + filtered_data["CostctrName"]
    )
    filtered_data["Vendor"] = (
            filtered_data["Vendor"] + " - " + filtered_data["Vendor Name"]
    )
    filtered_data["G/L"] = (
            filtered_data["G/L"] + " - " + filtered_data["G/L Name"]
    )
    filtered_data["Cost Ctr"] = filtered_data["Cost Ctr"].astype(str)
    filtered_data["G/L"] = filtered_data["G/L"].astype(str)
    filtered_data["Created"] = filtered_data["Created"].astype(str)
    filtered_data["Vendor"] = filtered_data["Vendor"].astype(str)
    filtered_data["year"] = filtered_data["year"].astype(str)
    dynamic_filters = DynamicFilters(
        filtered_data, filters=["year", "Cost Ctr", "G/L", "Vendor", "Created"]
    )
    dynamic_filters.display_filters(
        location="columns", num_columns=5, gap="large"
    )
    filtered_data = dynamic_filters.filter_df()
    filtered_data.reset_index(drop=True, inplace=True)
    filtered_data.index = filtered_data.index + 1
    filtered_data.rename_axis("S.NO", axis=1, inplace=True)
    filtered_datasheet2 = dynamic_filters.filter_df()
    filtered_datasheet2.reset_index(drop=True, inplace=True)
    filtered_datasheet2.index = filtered_data.index + 1
    filtered_datasheet2.rename_axis("S.NO", axis=1, inplace=True)
    for col in ["Cost Ctr", "G/L", "Vendor", "Created"]:
        filtered_datasheet2[col] = (
            filtered_datasheet2[col].str.split("-").str[0]
        )
    for col in ["Cost Ctr", "G/L", "Vendor", "Created"]:
        filtered_data[col] = filtered_data[col].str.split("-").str[0]
    df23 = filtered_data.copy()
    c1, card1, middle_column, card2, c2 = st.columns([1, 4, 1, 4, 1])
    with card1:
        Total_Amount_Alloted = filtered_data["Amount"].sum()

        # Check if the length of Total_Amount_Alloted is greater than 5
        if len(str(Total_Amount_Alloted)) > 5:
            # Get the integer part of the total amount
            integer_part = int(Total_Amount_Alloted)
            # Calculate the length of the integer part
            integer_length = len(str(integer_part))
            # Divide by 1 lakh if the integer length is greater than
            # 5 and less than or equal to 7
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
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div style='{CARD1_STYLE}'>"
            f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>{amount_display}</h2>"
            "</div>",
            unsafe_allow_html=True,
        )
        st.write(
            "<h2 style='text-align: center; font-size: 25px; font-weight: bold; color: black;'>Amount Spent -Category</h2>",
            unsafe_allow_html=True,
        )

        # Assuming 'filtered_data' is a DataFrame that has been defined earlier
        category_amount = (
            filtered_data.groupby("category")["Amount"].sum().reset_index()
        )

        fig = px.bar(
            category_amount,
            x="category",
            y="Amount",
            color="category",
            labels={"Amount": "Amount (in Crores)"},
            title="Amount Spent In Crores by Category",
            width=400,
            height=525,
            template="plotly_white",
        )

        fig.update_traces(texttemplate="%{customdata}", textposition="outside")

        # Add a loop to iterate over each trace and update its 'customdata' with the correct amount
        for i, amount in enumerate(category_amount["Amount"]):
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
            xaxis_title="Category",
            yaxis_title="Amount",
            font=dict(size=14, color="black"),
            showlegend=False,
            bargap=bargap_value,
        )

        st.plotly_chart(fig)

    with card2:
        # Assuming filtered_data is defined
        Total_Transaction = df23["Amount"].count()
        st.markdown(
            f"<h3 style='text-align: center; font-size: 25px;'>Total Count Of Transactions</h3>",
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div style='{CARD2_STYLE}'>"
            f"<h2 style='color: #28a745; text-align: center;'>{Total_Transaction:,}</h2>"
            "</div>",
            unsafe_allow_html=True,
        )
        st.write(
            "<h2 style='text-align: center; font-size: 25px; font-weight: bold; color: black;'>Transaction Count-Category</h2>",
            unsafe_allow_html=True,
        )
        st.write("")
        category_amount = (
            filtered_data.groupby("category")["Amount"].count().reset_index()
        )

        fig = px.bar(
            category_amount,
            x="category",
            y="Amount",
            color="category",
            labels={"Amount": "Transactions"},
            title="Transaction Count by Category",
            width=400,
            height=500,
            template="plotly_white",
        )

        # Add text annotations for each bar
        for trace in fig.data:
            for i, value in enumerate(trace.y):
                fig.add_annotation(
                    x=trace.x[i],
                    y=value,
                    text=f"{value}",
                    showarrow=True,
                    font=dict(size=12, color="black"),
                    align="center",
                    yshift=5,
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
            xaxis_title="Category",
            yaxis_title="Transactions",
            font=dict(size=14, color="black"),
            bargap=bargap_value,
        )
        # Show the plot in Streamlit
        st.plotly_chart(fig)
        st.write("")
    unique_categories = filtered_data["category"].unique()
    unique_categories = sorted(unique_categories, key=len, reverse=True)
    preferred_order = ["Vendor", "Employee", "Korean Expats", "Others"]

    # Sort categories according to preferred order
    unique_categories = sorted(
        unique_categories,
        key=lambda x: (
            preferred_order.index(x)
            if x in preferred_order
            else len(preferred_order)
        ),
    )

    # Define Cost_Ctr and G_L names
    Cost_Ctr = "Cost Ctr"
    G_L = "G/L"

    # Add 'All' option to the unique categories list
    unique_categories_with_options = (
            ["All"]
            + [f"Top 25 {category} Transactions" for category in unique_categories]
            + [f"Top 25 {Cost_Ctr} Transactions", f"Top 25 {G_L} Transactions"]
    )
    col1, col2, col3 = st.columns(3)
    # Selectbox to choose category or other options
    selected_category = col1.selectbox(
        "Select Category or Option:", unique_categories_with_options
    )

    # Get all unique values in the 'Cost Ctr' column
    all_cost_ctrs = filtered_data["Cost Ctr"].unique()

    # Set selected_cost_ctr to represent all values in the 'Cost Ctr' column
    selected_cost_ctr = (
        "All"  # or all_cost_ctrs if you want the default to be all values
    )
    filtered_data = filtered_data.copy()
    filtered_transactions = filtered_data.copy()

    # Filter the data based on the selected year and dropdown selections

    # filtered_transactions = filtered_transactions.drop_duplicates(subset=['Vendor', 'year'], keep='first')

    if selected_category != "All":
        if selected_category == f"Top 25 {Cost_Ctr} Transactions":
            filtered_data["overall_Alloted_Amount"] = filtered_data.groupby(
                ["year"]
            )["Amount"].transform("sum")
            filtered_data["Cumulative_Alloted/Cost Ctr"] = (
                filtered_data.groupby(["Cost Ctr"])["Amount"].transform("sum")
            )
            filtered_data["Cumulative_Alloted/Cost Ctr/Year"] = (
                filtered_data.groupby(["Cost Ctr", "year"])["Amount"].transform(
                    "sum"
                )
            )
            filtered_data["Percentage_Cumulative_Alloted/Cost Ctr"] = (
                                                                              filtered_data[
                                                                                  "Cumulative_Alloted/Cost Ctr/Year"]
                                                                              / filtered_data[
                                                                                  "overall_Alloted_Amount"]
                                                                      ) * 100
            yearly_total2 = (
                filtered_data.groupby("year")["Amount"].sum().reset_index()
            )
            yearly_total2.rename(
                columns={"Amount": "Total_Alloted_Amount/year"}, inplace=True
            )
            filtered_data["Percentage_Cumulative_Alloted/Cost Ctr/Year"] = (
                                                                                   filtered_data[
                                                                                       "Cumulative_Alloted/Cost Ctr/Year"]
                                                                                   / filtered_data[
                                                                                       "overall_Alloted_Amount"]
                                                                           ) * 100

            filtered_data = filtered_data.sort_values(
                by="Cumulative_Alloted/Cost Ctr/Year", ascending=False
            )[
                [
                    "Cost Ctr",
                    "CostctrName",
                    "Cumulative_Alloted/Cost Ctr/Year",
                    "Percentage_Cumulative_Alloted/Cost Ctr/Year",
                ]
            ].drop_duplicates(
                subset=["Cost Ctr"], keep="first"
            )
            filtered_data.rename(
                columns={
                    "Cumulative_Alloted/Cost Ctr/Year": "Value (In ₹)",
                    "CostctrName": "Name",
                    "Cost Ctr": "Cost Center",
                    "Percentage_Cumulative_Alloted/Cost Ctr/Year": "%total",
                },
                inplace=True,
            )
            filtered_data.reset_index(drop=True, inplace=True)
            filtered_data.index = filtered_data.index + 1
            filtered_data.rename_axis("S.NO", axis=1, inplace=True)
            filtered_transactions["Cummulative_transactions2"] = len(
                filtered_transactions
            )
            filtered_transactions["Cumulative_transactions/Cost Ctr"] = (
                filtered_transactions.groupby(["Cost Ctr"])[
                    "Cost Ctr"
                ].transform("count")
            )
            filtered_transactions["Cumulative_transactions/Cost Ctr/Year"] = (
                filtered_transactions.groupby(["Cost Ctr"])[
                    "Cost Ctr"
                ].transform("count")
            )
            filtered_transactions["percentage Transcation/Cost Ctr/year"] = (
                    filtered_transactions["Cumulative_transactions/Cost Ctr/Year"]
                    / filtered_transactions["Cummulative_transactions2"]
                    * 100
            )
            filtered_transactions = filtered_transactions.sort_values(
                by="percentage Transcation/Cost Ctr/year", ascending=False
            )[
                [
                    "Cost Ctr",
                    "CostctrName",
                    "Cumulative_transactions/Cost Ctr/Year",
                    "percentage Transcation/Cost Ctr/year",
                ]
            ].drop_duplicates(
                subset=["Cost Ctr"], keep="first"
            )
            filtered_transactions.rename(
                columns={
                    "Cost Ctr": "Cost Center",
                    "CostctrName": "Name",
                    "Cumulative_transactions/Cost Ctr/Year": "Transactions",
                    "percentage Transcation/Cost Ctr/year": "% total",
                },
                inplace=True,
            )
            filtered_transactions.reset_index(drop=True, inplace=True)
            filtered_transactions.index = filtered_transactions.index + 1
            filtered_data.rename_axis("S.NO", axis=1, inplace=True)
            filtered_transactions.reset_index(drop=True, inplace=True)
            filtered_transactions.index = filtered_transactions.index + 1
            filtered_transactions.rename_axis("S.NO", axis=1, inplace=True)
            merged_df = pd.merge(
                filtered_data, filtered_transactions, on=["Cost Center", "Name"]
            )

        elif selected_category == f"Top 25 {G_L} Transactions":
            filtered_data["overall_Alloted_Amount"] = filtered_data.groupby(
                ["year"]
            )["Amount"].transform("sum")
            filtered_data["Cumulative_Alloted/G/L"] = filtered_data.groupby(
                ["G/L"]
            )["Amount"].transform("sum")
            filtered_data["Cumulative_Alloted/G/L/Year"] = (
                filtered_data.groupby(["G/L", "year"])["Amount"].transform(
                    "sum"
                )
            )
            filtered_data["Percentage_Cumulative_Alloted/G/L"] = (
                                                                         filtered_data[
                                                                             "Cumulative_Alloted/G/L/Year"]
                                                                         / filtered_data[
                                                                             "overall_Alloted_Amount"]
                                                                 ) * 100
            yearly_total2 = (
                filtered_data.groupby("year")["Amount"].sum().reset_index()
            )
            yearly_total2.rename(
                columns={"Amount": "Total_Alloted_Amount/year"}, inplace=True
            )
            filtered_data["Percentage_Cumulative_Alloted/G/L/Year"] = (
                                                                              filtered_data[
                                                                                  "Cumulative_Alloted/G/L/Year"]
                                                                              / filtered_data[
                                                                                  "overall_Alloted_Amount"]
                                                                      ) * 100
            filtered_data = filtered_data.sort_values(
                by="Cumulative_Alloted/G/L/Year", ascending=False
            )[
                [
                    "G/L",
                    "G/L Name",
                    "Cumulative_Alloted/G/L/Year",
                    "Percentage_Cumulative_Alloted/G/L/Year",
                ]
            ].drop_duplicates(
                subset=["G/L"], keep="first"
            )
            filtered_data.rename(
                columns={
                    "Cumulative_Alloted/G/L/Year": "Value (In ₹)",
                    "G/LName": "Name",
                    "Percentage_Cumulative_Alloted/G/L/Year": "%total",
                },
                inplace=True,
            )
            filtered_data.reset_index(drop=True, inplace=True)
            filtered_data.index = filtered_data.index + 1
            filtered_data.rename_axis("S.NO", axis=1, inplace=True)
            filtered_transactions["Cummulative_transactions2"] = len(
                filtered_transactions
            )
            filtered_transactions["Cumulative_transactions/G/L"] = (
                filtered_transactions.groupby(["G/L"])["G/L"].transform("count")
            )
            filtered_transactions["Cumulative_transactions/G/L/Year"] = (
                filtered_transactions.groupby(["G/L"])["G/L"].transform("count")
            )
            filtered_transactions["percentage Transcation/G/L/year"] = (
                    filtered_transactions["Cumulative_transactions/G/L/Year"]
                    / filtered_transactions["Cummulative_transactions2"]
                    * 100
            )
            filtered_transactions = filtered_transactions.sort_values(
                by="percentage Transcation/G/L/year", ascending=False
            )[
                [
                    "G/L",
                    "G/L Name",
                    "Cumulative_transactions/G/L/Year",
                    "percentage Transcation/G/L/year",
                ]
            ].drop_duplicates(
                subset=["G/L"], keep="first"
            )
            filtered_transactions.rename(
                columns={
                    "Cumulative_transactions/G/L/Year": "Transactions",
                    "percentage Transcation/G/L/year": "% total",
                },
                inplace=True,
            )
            filtered_transactions.reset_index(drop=True, inplace=True)
            filtered_transactions.index = filtered_transactions.index + 1
            filtered_transactions.rename_axis("S.NO", axis=1, inplace=True)
            merged_df = pd.merge(
                filtered_data, filtered_transactions, on=["G/L", "G/L Name"]
            )

        else:
            category = " ".join(selected_category.split()[2:-1])
            filtered_data = filtered_data[filtered_data["category"] == category]
            filtered_transactions = filtered_data.copy()
            filtered_data["Amount_used/Year2"] = filtered_data.groupby(
                ["Vendor", "year", "category"]
            )["Amount"].transform("sum")
            filtered_data["Yearly_Alloted_Amount\Category2"] = (
                filtered_data.groupby(["category", "year"])["Amount"].transform(
                    "sum"
                )
            )
            filtered_data["percentage_of_amount/category_used/year2"] = (
                                                                                filtered_data[
                                                                                    "Amount_used/Year2"]
                                                                                / filtered_data[
                                                                                    "Yearly_Alloted_Amount\Category2"]
                                                                        ) * 100
            filtered_transactions = filtered_transactions[
                filtered_transactions["category"] == category
                ]
            filtered_data = filtered_data.sort_values(
                by="percentage_of_amount/category_used/year2", ascending=False
            )[
                [
                    "Vendor",
                    "Vendor Name",
                    "Amount_used/Year2",
                    "percentage_of_amount/category_used/year2",
                ]
            ]
            filtered_data = filtered_data.drop_duplicates(
                subset=["Vendor"], keep="first"
            )
            filtered_data.rename(
                columns={
                    "Amount_used/Year2": "Value (In ₹)",
                    "Vendor": "ID",
                    "Vendor Name": "Name",
                    "percentage_of_amount/category_used/year2": "%total",
                },
                inplace=True,
            )

            filtered_data.reset_index(drop=True, inplace=True)
            filtered_data.index = filtered_data.index + 1
            filtered_data.rename_axis("S.NO", axis=1, inplace=True)
            filtered_transactions["Transations/year/Vendor2"] = (
                filtered_transactions.groupby(["Vendor", "year", "category"])[
                    "Vendor"
                ].transform("count")
            )
            filtered_transactions["overall_transactions/category/year2"] = (
                filtered_transactions.groupby(["category", "year"])[
                    "category"
                ].transform("count")
            )
            filtered_transactions["percentransations_made/category/year2"] = (
                                                                                     filtered_transactions[
                                                                                         "Transations/year/Vendor2"]
                                                                                     / filtered_transactions[
                                                                                         "overall_transactions/category/year2"]
                                                                             ) * 100
            filtered_transactions = filtered_transactions.sort_values(
                by="percentransations_made/category/year2", ascending=False
            )[
                [
                    "Vendor",
                    "Vendor Name",
                    "Transations/year/Vendor2",
                    "percentransations_made/category/year2",
                ]
            ]
            filtered_transactions = filtered_transactions.drop_duplicates(
                subset=["Vendor"], keep="first"
            )
            filtered_transactions.rename(
                columns={
                    "Vendor": "ID",
                    "Vendor Name": "Name",
                    "Transations/year/Vendor2": "Transactions",
                    "percentransations_made/category/year2": "% total",
                },
                inplace=True,
            )
            filtered_transactions.reset_index(
                drop=True, inplace=True
            )  # Reset index here
            filtered_transactions.index = filtered_transactions.index + 1
            filtered_transactions.rename_axis("S.NO", axis=1, inplace=True)
            merged_df = pd.merge(
                filtered_data, filtered_transactions, on=["ID", "Name"]
            )
    else:
        filtered_data["Amount_used/Year2"] = filtered_data.groupby(
            ["Vendor", "year"]
        )["Amount"].transform("sum")
        filtered_data["Yearly_Alloted_Amount\Category2"] = (
            filtered_data.groupby(["year"])["Amount"].transform("sum")
        )
        filtered_data["percentage_of_amount/category_used/year2"] = (
                                                                            filtered_data["Amount_used/Year2"]
                                                                            / filtered_data[
                                                                                "Yearly_Alloted_Amount\Category2"]
                                                                    ) * 100
        filtered_data = filtered_data.sort_values(
            by="percentage_of_amount/category_used/year2", ascending=False
        )[
            [
                "Vendor",
                "Vendor Name",
                "Amount_used/Year2",
                "percentage_of_amount/category_used/year2",
            ]
        ]
        filtered_data = filtered_data.drop_duplicates(
            subset=["Vendor"], keep="first"
        )
        filtered_data.rename(
            columns={
                "Amount_used/Year2": "Value (In ₹)",
                "Vendor": "ID",
                "Vendor Name": "Name",
                "percentage_of_amount/category_used/year2": "%total",
            },
            inplace=True,
        )

        filtered_data.reset_index(drop=True, inplace=True)
        filtered_data.index = filtered_data.index + 1
        filtered_data.rename_axis("S.NO", axis=1, inplace=True)
        filtered_transactions["Transations/year/Vendor2"] = (
            filtered_transactions.groupby(["Vendor", "year"])[
                "Vendor"
            ].transform("count")
        )
        filtered_transactions["overall_transactions/category/year2"] = (
            filtered_transactions["overall_transactions/category/year2"]
        ) = filtered_transactions.groupby(["year"])["category"].transform(
            "count"
        )
        filtered_transactions["percentransations_made/category/year2"] = (
                                                                                 filtered_transactions[
                                                                                     "Transations/year/Vendor2"]
                                                                                 / filtered_transactions[
                                                                                     "overall_transactions/category/year2"]
                                                                         ) * 100
        filtered_transactions = filtered_transactions.sort_values(
            by="percentransations_made/category/year2", ascending=False
        )[
            [
                "Vendor",
                "Vendor Name",
                "Transations/year/Vendor2",
                "percentransations_made/category/year2",
            ]
        ]
        filtered_transactions = filtered_transactions.drop_duplicates(
            subset=["Vendor"], keep="first"
        )
        filtered_transactions.rename(
            columns={
                "Vendor": "ID",
                "Vendor Name": "Name",
                "Transations/year/Vendor2": "Transactions",
                "percentransations_made/category/year2": "% total",
            },
            inplace=True,
        )
        filtered_transactions.reset_index(
            drop=True, inplace=True
        )  # Reset index here
        filtered_transactions.index = filtered_transactions.index + 1
        filtered_transactions.rename_axis("S.NO", axis=1, inplace=True)
        merged_df = pd.merge(
            filtered_data, filtered_transactions, on=["ID", "Name"]
        )
    col1, col2 = st.columns(2)

    with col1:

        st.write("Value wise (In \u20B9)")
        st.write(filtered_data.head(25))

    with col2:
        st.write("Transaction Count Wise")
        st.write(filtered_transactions.head(25))
    excel_buffer = BytesIO()
    filtered_datasheet2.rename(
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
    filtered_datasheet2 = filtered_datasheet2[
        [
            "Payable req.no",
            "Doc.Type",
            "ReimbursementID",
            "Name",
            "category",
            "Amount",
            "year",
            "Cost Center",
            "CostctrName",
            "G/L",
            "G/L Name",
            "Document Date",
            "Posting Date",
            "Document No",
            "Invoice Number",
            "Invoice Reference Number",
            "Text",
            "Creator ID",
            "Verifier ID",
            "Verified on",
            "HOG Approval by",
            "HOD Apr/Rej by",
            "HOG Approval on",
            "Reference invoice",
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
            "Clearing date",
        ]
    ]
    filtered_datasheet2 = filtered_datasheet2.sort_values(
        by=["ReimbursementID", "Posting Date", "Amount"]
    )
    filtered_datasheet2 = filtered_datasheet2.drop(columns=["year"])
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        merged_df.to_excel(writer, index=False, sheet_name="Summary")
        filtered_datasheet2.to_excel(writer, index=False, sheet_name="Raw Data")
    # Reset the buffer's position to the start for reading
    excel_buffer.seek(0)
    # Convert Excel buffer to base64
    excel_b64 = base64.b64encode(excel_buffer.getvalue()).decode()
    # Download link for Excel file within a Markdown
    download_link = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}" download="PO Analysis.xlsx">Download Excel file</a>'
    st.markdown(download_link, unsafe_allow_html=True)