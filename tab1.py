from analyze_excelfinal import *
def datatab1(data):
    filtered_df = data.copy()
    filtered_df["Cost Ctr"] = filtered_df["Cost Ctr"].astype(str)
    filtered_df["G/L"] = filtered_df["G/L"].astype(str)
    filtered_df["Vendor"] = filtered_df["Vendor"].astype(str)
    filtered_df["Created"] = filtered_df["Created"].astype(str)
    vendor_name_dict = dict(zip(filtered_df["Vendor"], filtered_df["Vendor Name"]))
    # Map 'Vendor Name' to a new column 'CreatedName' based on 'Created'
    filtered_df["CreatedName"] = filtered_df["Created"].map(vendor_name_dict)
    # Fill NaN values in 'CreatedName' with an empty string
    filtered_df["CreatedName"] = filtered_df["CreatedName"].fillna("")

    # Concatenate 'Created' with 'CreatedName', separated by ' - ', only if 'CreatedName' is not empty
    filtered_df["Created"] = filtered_df.apply(
        lambda row: (
            row["Created"] + " - " + row["CreatedName"]
            if row["CreatedName"]
            else row["Created"]
        ),
        axis=1,
    )

    filtered_df["G/L"] = filtered_df["G/L"] + " - " + filtered_df["G/L Name"]
    filtered_df["Vendor"] = (
            filtered_df["Vendor"] + " - " + filtered_df["Vendor Name"]
    )
    filtered_df["Cost Ctr"] = (
            filtered_df["Cost Ctr"] + " - " + filtered_df["CostctrName"]
    )
    filtered_df["Cost Ctr"] = filtered_df["Cost Ctr"].astype(str)
    filtered_df["G/L"] = filtered_df["G/L"].astype(str)
    filtered_df["Vendor"] = filtered_df["Vendor"].astype(str)
    filtered_df["Created"] = filtered_df["Created"].astype(str)
    vendor_name_dict = dict(zip(filtered_df["Vendor"], filtered_df["Vendor Name"]))
    filtered_df.drop("CreatedName", axis=1, inplace=True)
    filtered_data = filtered_df.copy()
    c1, c2, c3, c4, c5 = st.columns(5)

    # Cost Center multiselect
    options_cost_center = ["All"] + [
        yr for yr in filtered_data["Cost Ctr"].unique() if yr != "All"
    ]
    selected_cost_centers = c1.multiselect(
        "Select Cost Centers", options_cost_center, default=["All"]
    )
    if "All" not in selected_cost_centers:
        filtered_data = filtered_data[
            filtered_data["Cost Ctr"].isin(selected_cost_centers)
        ]

    # G/L multiselect
    options_gl = ["All"] + [
        yr for yr in filtered_data["G/L"].unique() if yr != "All"
    ]
    selected_gl = c2.multiselect("Select G/L", options_gl, default=["All"])
    if "All" not in selected_gl:
        filtered_data = filtered_data[filtered_data["G/L"].isin(selected_gl)]

    # Vendor multiselect
    options_vendor = ["All"] + [
        yr for yr in filtered_data["Vendor"].unique() if yr != "All"
    ]
    selected_vendor = c3.multiselect(
        "Select Reimbursing ID", options_vendor, default=["All"]
    )
    if "All" not in selected_vendor:
        filtered_data = filtered_data[filtered_data["Vendor"].isin(selected_vendor)]
    # creator multiselect
    options_creator = ["All"] + [
        yr for yr in filtered_data["Created"].unique() if yr != "All"
    ]
    selected_creator = c4.multiselect(
        "Select a Creator", options_creator, default=["All"]
    )
    if "All" not in selected_creator:
        filtered_data = filtered_data[
            filtered_data["Created"].isin(selected_creator)
        ]
    filtered_data["Amount"] = filtered_data["Amount"].round()
    filtered_data["Amount"] = filtered_data["Amount"].astype(int)
    card1, card2 = st.columns(2)
    for col in ["Cost Ctr", "G/L", "Vendor", "Created"]:
        filtered_data[col] = filtered_data[col].str.split("-").str[0]
    with card1:
        Total_Amount_Alloted = filtered_data["Amount"].sum()

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
            unsafe_allow_html=True,
        )
        st.markdown(
            f"<div style='{CARD1_STYLE}'>"
            f"<h2 style='color: #007bff; text-align: center; font-size: 35px;'>{amount_display}</h2>"
            "</div>",
            unsafe_allow_html=True,
        )
        st.write("")

    with card2:
        # Assuming filtered_data is defined
        Total_Transaction = filtered_data["Payable req.no"].count()
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
    data = filtered_data.copy()
    data.reset_index(drop=True, inplace=True)
    data.index += 1  # Start index from 1

    years = data["year"].unique()
    years_df = pd.DataFrame({"year": data["year"].unique()})
    filtered_years = data["year"].unique()
    filtered_data = data[data["year"].isin(filtered_years)].copy()

    st.write("## Employee Reimbursement Trend")

    c2, col1, col2, c5 = st.columns([1, 4, 4, 1])
    c2.write("")
    c5.write("")
    fig_employee = line_plot_overall_transactions(
        data, "Employee", years_df["year"]
    )
    fig_employee2 = line_plot_used_amount(data, "Employee", years)
    col2.plotly_chart(fig_employee)
    col1.plotly_chart(fig_employee2)

    # Line plot for category "Korean Expats"
    st.write("## Korean Expats Reimbursement Trend")
    c2, co1, co2, c5 = st.columns([1, 4, 4, 1])
    fig_korean_expats = line_plot_overall_transactions(
        data, "Korean Expats", years_df["year"]
    )
    fig_korean_expats1 = line_plot_used_amount(data, "Korean Expats", years)
    c2.write("")
    c5.write("")
    co2.plotly_chart(fig_korean_expats)
    co1.plotly_chart(fig_korean_expats1)

    # Line plot for category "Vendor"
    st.write("## Vendor Payment Trend")
    co2, c1, c2, c5 = st.columns([1, 4, 4, 1])
    co2.write("")
    c5.write("")
    fig_vendor = line_plot_overall_transactions(data, "Vendor", years_df["year"])
    fig_vendor1 = line_plot_used_amount(data, "Vendor", years)
    c2.plotly_chart(fig_vendor)
    c1.plotly_chart(fig_vendor1)

    # Line plot for category "Others"
    st.write("## Other Payment Trend")
    c2, cl1, cl2, c5 = st.columns([1, 4, 4, 1])
    fig_others = line_plot_overall_transactions(data, "Others", years_df["year"])
    fig_others1 = line_plot_used_amount(data, "Others", years)
    cl2.plotly_chart(fig_others)
    cl1.plotly_chart(fig_others1)
