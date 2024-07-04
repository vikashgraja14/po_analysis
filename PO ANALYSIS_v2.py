"""
This module imports the necessary libraries for data processing and visualization.

Modules:
    os: Provides a way of using operating system dependent functionality
    like reading or writing to the file system.
    re: Offers a set of functions that allows us to search a string for
     a match (regular expressions).
    difflib: Provides tools for comparing sequences, especially useful
    for comparing text files.
    pandas: A powerful data manipulation and analysis library for Python.
    Here it's used to read and manipulate Excel data.
    plotly.express: A high-level interface to Plotly, a graphing library.
    It's used for creating interactive plots and visualizations.
    streamlit: A framework for creating web apps in pure Python. Used for
    displaying and interacting with data and visualizations in a web interface.
"""
import os
import re
import difflib
from functools import partial
import pandas as pd
import plotly.express as px
import streamlit as st

directory = os.path.dirname(__file__)
os.chdir(directory)
dfholiday = pd.read_excel(
    "unlocked holiday.xlsx")
dfholiday2 = pd.read_excel(
    "unlocked holiday.xlsx")

@st.cache_resource(show_spinner=False)
def process_data(files):
    """
      Process Excel files containing non-PO payment data.

      Args:
      - files (list): List of file paths or file objects (xlsx format)
      containing non-PO payment data.

      Returns:
      - pandas.DataFrame: Concatenated DataFrame containing processed
      non-PO payment data.

      This function reads each Excel file from the provided list, processes them
      into a standardized format,
      and concatenates them into a single DataFrame for further analysis.

      """
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
    df['Clearing doc no.'] = df['Clearing doc no.'].astype(str)
    df = df[df['Clearing doc no.'].str.startswith('5')]
    df = df.query('`Clearing doc no.`.str.startswith("5")')

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
    filtered_df['HOD Apr/Rej on'] = pd.to_datetime(filtered_df['HOD Apr/Rej on'], errors='coerce')
    filtered_df['Clearing date'] = pd.to_datetime(filtered_df['Clearing date'], errors='coerce')
    filtered_df['year'] = filtered_df['Pstng Date'].dt.year
    filtered_df['HOD Apr/Rej on'] = filtered_df['HOD Apr/Rej on'].dt.date
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
    df['Clearing doc no.'] = df['Clearing doc no.'].astype(str)
    df['Clearing doc no.'] = df['Clearing doc no.'].apply(
        lambda x: str(x) if isinstance(x, str) else '')
    df['Clearing doc no.'] = df['Clearing doc no.'].apply(lambda x: re.sub(r'\..*', '', x))
    df['Clearing doc no.'] = df['Clearing doc no.'].astype(str)
    df['Clearing doc no.'] = df['Clearing doc no.'].apply(
        lambda x: str(x) if isinstance(x, str) else '')
    df['Clearing doc no.'] = df['Clearing doc no.'].apply(
        lambda x: re.sub(r'\..*', '', x))
    df['Document No'] = df['Document No'].astype(str)
    df['Document No'] = df['Document No'].apply(lambda x: str(x) if isinstance(x, str) else '')
    df['Document No'] = df['Document No'].apply(lambda x: re.sub(r'\..*', '', x))
    df['HOG Approval by'] = df['HOG Approval by'].astype(str)
    df['HOG Approval by'] = df['HOG Approval by'].apply(lambda x: str(x) if isinstance(x, str) else '')
    df['HOG Approval by'] = df['HOG Approval by'].apply(lambda x: re.sub(r'\..*', '', x))
    df.reset_index(drop=True, inplace=True)
    return df
def optimized_similar_invoices(filtered_df, column_name="Invoice Number", threshold=0.8):
  """
  Efficiently identifies rows with similar invoice numbers based on Jaccard similarity (without sets).

  Args:
      df (pd.DataFrame): Input DataFrame containing the column with invoice numbers.
      column_name (str, optional): Name of the column containing invoice numbers. Defaults to "column 1".
      threshold (float, optional): Similarity threshold (0 to 1) for considering invoices similar. Defaults to 0.8 (80%).

  Returns:
      pd.DataFrame: A DataFrame containing rows with potentially similar invoice numbers (including all columns).
  """

  def jaccard_similarity(s1, s2):
    """
    Calculates Jaccard similarity between two strings (without sets). Optimized for counting differences.

    Args:
        s1 (str): First string.
        s2 (str): Second string.

    Returns:
        float: Jaccard similarity score between 0 and 1.
    """
    min_len = min(len(s1), len(s2))
    diff_count = 0
    for i in range(min_len):
      if s1[i] != s2[i]:
        diff_count += 1

    # Early termination if too many differences
    if diff_count > threshold * min_len:
      return 0

    # Count remaining characters for Jaccard similarity (considering potential remaining differences)
    return float(min_len - diff_count) / (len(set(s1)) + len(set(s2)) - min_len + diff_count)

  # Add a new column for cleaned invoice numbers (remove non-alphanumeric characters)
  filtered_df["Clean Invoice Number"] = filtered_df[column_name].str.replace(r"[^\w]+", "", regex=True)

  # Create empty lists to store similar invoice numbers and already checked invoices
  similar_invoices = []
  checked_invoices = {}

  # Iterate through each row
  for i, row in filtered_df.iterrows():
    invoice_number = row["Clean Invoice Number"]

    # Skip already checked invoices
    if invoice_number in checked_invoices:
      continue

    # Compare with subsequent rows
    for j in range(i + 1, len(filtered_df)):
      other_invoice = filtered_df.loc[j, "Clean Invoice Number"]

      if other_invoice in checked_invoices:
        continue

      similarity = jaccard_similarity(invoice_number, other_invoice)

      if similarity >= threshold:
        similar_invoices.append(invoice_number)
        similar_invoices.append(other_invoice)
        checked_invoices.update([invoice_number, other_invoice])
        break  # No need to check further for this invoice

  # Filter the DataFrame based on similar invoice numbers
  return filtered_df[filtered_df["Clean Invoice Number"].isin(similar_invoices)]


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



@st.cache_resource(show_spinner=False)
def has_special_characters(s):
    return re.search(r'[^A-Za-z0-9]+', s) is not None

def is_similar(s1, s2):
    # Remove special characters from both strings
    s1_clean = re.sub(r'[^A-Za-z0-9]+', '', str(s1))
    s2_clean = re.sub(r'[^A-Za-z0-9]+', '', str(s2))
    # Check if the cleaned strings are equal
    return s1_clean == s2_clean


def identify_and_return_all_duplicates(filtered_df):
  """
  Efficiently identifies rows with potentially duplicate invoice numbers, considering all special characters,
  and returns a DataFrame containing all rows (including duplicates) from the original DataFrame.

  Args:
      df (pd.DataFrame): Input DataFrame containing an "Invoice Number" column.

  Returns:
      pd.DataFrame: A DataFrame containing all rows (including duplicates) from the original DataFrame.
  """

  # Regular expression to match any character that is not alphanumeric or hyphen
  special_char_pattern = r"[^\w-]+"

  # Vectorized removal of special characters using pandas
  filtered_df["Clean Invoice Number"] = filtered_df["Invoice Number"].str.replace(
      pat=special_char_pattern,
      repl="",
      regex=True
  )

  # Efficient duplicate check with vectorized operations
  duplicates = filtered_df[filtered_df["Clean Invoice Number"].duplicated(keep="first")]

  # Merge based on "Invoice Number" column (assuming it's unique)
  filtered_df = filtered_df.merge(duplicates, how='outer', on='Invoice Number')

  return filtered_df

@st.cache_resource(show_spinner=False)
def line_plot_overall_transactions(data, category, years, width=400, height=300):
    """
    Create a line plot showing the count of transactions over
     the specified years for a given category.

    Args:
        data (pd.DataFrame): The input data containing transaction records.
        category (str): The category of transactions to filter by.
        years (list of int): The years to include in the plot.
        width (int, optional): The width of the plot. Defaults to 400.
        height (int, optional): The height of the plot. Defaults to 300.

    Returns:
        plotly.graph_objs._figure.Figure: The Plotly figure object
        containing the line plot.

    """
    filtered_data = data[(data["category"] == category) & (data["year"].isin(years))]
    data_length = filtered_data.groupby("year").size().reset_index(name="data_len")

    # Create the line plot
    fig = px.line(
        data_length,
        x="year",
        y="data_len",
        title="Transactions Count",
        labels={"year": "Year", "data_len": "Transactions"},
        markers=True,
    )

    # Set mode to 'lines+markers'
    fig.update_traces(mode="lines+markers")

    # Update layout with width and height
    fig.update_layout(width=width, height=height)

    # Modify x-axis labels to include hyphens between years
    fig.update_xaxes(
        tickvals=years,
        ticktext=[re.sub(r"(\d{4})(\d{2})", r"\1-\2", str(year)) for year in years],
    )

    # Format y-axis ticks as integers
    fig.update_yaxes(tickformat=".0f")

    # Add text annotations to data points
    for year, count in zip(data_length["year"], data_length["data_len"]):
        fig.add_annotation(
            x=year,
            y=count,
            text=str(count),
            showarrow=False,
            font={"size": 12, "color": "black"},  # Use dictionary literal directly
            align="center",
            yshift=13,
        )
    return fig


@st.cache_resource(show_spinner=False)
def line_plot_used_amount(data, category, years, width=400, height=300):
    """
    Create a line plot showing the total amount of transactions over
     the specified years for a given category.

    Args:
        data (pd.DataFrame): The input data containing transaction records.
        category (str): The category of transactions to filter by.
        years (list of int): The years to include in the plot.
        width (int, optional): The width of the plot. Defaults to 400.
        height (int, optional): The height of the plot. Defaults to 300.

    Returns:
        plotly.graph_objs._figure.Figure: The Plotly figure object containing the line plot.

    """
    # Filter the data
    data = data[(data["category"] == category) & (data["year"].isin(years))]

    # Group by year and sum the amounts
    amount_length = (
        data.groupby("year")["Amount"].sum().reset_index(name="amount_length")
    )

    # Create the line plot
    fig = px.line(
        amount_length,
        x="year",
        y="amount_length",
        title="Transactions Value",
        labels={"year": "Year", "amount_length": "Amount"},
        markers=True,
    )

    # Set mode to 'lines+markers'
    fig.update_traces(mode="lines+markers")

    # Update layout with width and height
    fig.update_layout(width=width, height=height)

    # Modify x-axis labels to include hyphens between years
    fig.update_xaxes(
        tickvals=years,
        ticktext=[re.sub(r"(\d{4})(\d{2})", r"\1-\2", str(year)) for year in years],
    )

    # Define the format_amount function
    def format_amount(amount_value):
        """
        Format the given amount based on its value.

        Parameters:
        - amount_value (float or int): The amount to format.

        Returns:
        - str: Formatted string representation of the amount.
        """
        integer_part = int(amount_value)
        integer_length = len(str(integer_part))

        if 5 < integer_length <= 7:
            return f"₹ {amount_value / 100000:,.2f} lks"

        if integer_length > 7:
            return f"₹ {amount_value / 10000000:,.2f} crs"

        return f"₹ {amount_value:,.2f}"

    # Add formatted annotations on top of data points
    for year, amount in zip(amount_length["year"], amount_length["amount_length"]):
        formatted_amount = format_amount(amount)
        fig.add_annotation(
            x=year,
            y=amount,
            text=formatted_amount,
            showarrow=False,
            font={"size": 12, "color": "black"},
            align="center",
            yshift=13,
        )

    return fig



@st.cache_resource(show_spinner=False)
def check_similarity(s1, s2, threshold=0.7):
    """
    Check if two strings are similar based on a similarity threshold.

    Parameters:
    - s1 (str): First string to compare.
    - s2 (str): Second string to compare.
    - threshold (float, optional): Minimum similarity ratio to consider strings similar.
    Defaults to 0.8.

    Returns:
    - bool: True if the similarity ratio between s1 and s2 is greater than or equal
    to the threshold, False otherwise.
    """
    s1 = str(s1)
    s2 = str(s2)
    similarity = difflib.SequenceMatcher(None, s1, s2).ratio()
    return similarity >= threshold

def optimize_similarity_check(filtered_df, threshold=0.7):
    """
    Efficiently identifies invoice numbers with high character resemblance in a DataFrame.

    Parameters:
    - filtered_df (pd.DataFrame): Input DataFrame containing invoice numbers.
    - threshold (float, optional): Minimum similarity score to consider strings close.
    Defaults to 0.8.

    Returns:
    - pd.DataFrame: A new DataFrame containing only rows with potentially matching invoice numbers.
    """

    # Create a vectorized function to check resemblance efficiently
    vectorized_check = partial(check_similarity, threshold=threshold)

    # Efficiently compare each invoice number with all others based on character similarity
    similar_pairs = filtered_df["Invoice Number"].apply(vectorized_check, args=(filtered_df["Invoice Number"],))

    # Mark rows with potentially matching invoice numbers using a more descriptive name
    filtered_df["potentially_matching"] = similar_pairs  # New column name

    return filtered_df

def filter_auth(filtered_df, checked_columns_auth, filename):
    """
    Filter DataFrame based on checked authorization columns and sort results.

    Parameters:
    - filtered_df (pd.DataFrame): DataFrame to filter.
    - checked_columns_auth (list): List of column names to check for authorization.
    - filename (str): Name of the output file.

    Returns:
    - tuple or None: Tuple containing filtered DataFrame, checked columns list,
    and filename if successful,
      or None, None, None if an error occurs.
    """
    if len(checked_columns_auth) == 1:
        st.error("Please select another checkbox to verify Authorization Parameters.")
        return None, None, None
    else:
        for i in range(len(checked_columns_auth) - 1):
            col_i = checked_columns_auth[i]
            col_j = checked_columns_auth[i + 1]
            filtered_df = filtered_df[filtered_df[col_i] == filtered_df[col_j]]

    sort_columns = checked_columns_auth.copy()
    if "Reimbursement ID" not in checked_columns_auth and "Cost Center" not in checked_columns_auth:
        sort_columns += ["Reimbursement ID", "Cost Center"]
    elif "Reimbursement ID" in checked_columns_auth and "Cost Center" not in checked_columns_auth:
        sort_columns += ["Cost Center"]
    elif "Cost Center" in checked_columns_auth and "Reimbursement ID" not in checked_columns_auth:
        sort_columns += ["Reimbursement ID"]

    filtered_df = filtered_df.sort_values(by=sort_columns)
    filtered_df.reset_index(drop=True, inplace=True)
    filtered_df.index += 1

    filename = "Transactions_with_same_column.xlsx"

    try:
        return filtered_df, checked_columns_auth, filename
    except FileNotFoundError as e:
        # Handle specific exception (e.g., file not found)
        st.error(f"File not found: {e}")
        return None, None, None


def filter_gen(filtered_df, checked_columns_gen, filename):
    """
    Filter and sort DataFrame based on checked columns.

    Parameters:
    - filtered_df (pd.DataFrame): DataFrame to filter.
    - checked_columns_gen (list): List of columns to check for duplicates.
    - filename (str): Filename to save results.

    Returns:
    - pd.DataFrame or None: Filtered DataFrame, checked_columns_gen, filename.
    """
    filtered_df = filtered_df[
        filtered_df.duplicated(subset=checked_columns_gen, keep=False)
    ]

    sort_columns = checked_columns_gen.copy()
    if "Reimbursement ID" not in checked_columns_gen and "Cost Center" not in checked_columns_gen:
        sort_columns += ["Reimbursement ID", "Cost Center"]
    elif "Reimbursement ID" in checked_columns_gen and "Cost Center" not in checked_columns_gen:
        sort_columns += ["Cost Center"]
    elif "Cost Center" in checked_columns_gen and "Reimbursement ID" not in checked_columns_gen:
        sort_columns += ["Reimbursement ID"]

    filtered_df = filtered_df.sort_values(by=sort_columns)
    filtered_df.reset_index(drop=True, inplace=True)
    filtered_df.index += 1

    filename = "Transactions_with_same_column.xlsx"

    try:
        return filtered_df, checked_columns_gen, filename

    except pd.errors.EmptyDataError as e:
        st.error(f"An error occurred: {e}")
        return None, None, None


# Example usage (assuming filtered_df and checked_columns_gen are defined)
# try:
#     filtered_df, checked_columns_gen, filename = filter_groups_by_size(filtered_df, checked_columns_gen, "Transactions_with_same_column.xlsx")
# except pd.errors.EmptyDataError as e:
#     print(f"An error occurred: {e}")
#     filtered_df, checked_columns_gen, filename = None, None, None

def filter_spec(filtered_df, checked_columnsspec, dfholiday, filename):
    """
    Filter the dataframe based on specified columns and holiday dates,
     and perform various transformations
    including finding similar invoices and handling special characters
     in invoice numbers.

    Args:
        filtered_df (pd.DataFrame): The dataframe to be filtered and processed.
        checked_columnsspec (list of str): The columns based on which
        filtering and processing will be applied.
        dfholiday (pd.DataFrame): The dataframe containing holiday dates.
        filename (str): The filename for saving the processed dataframe.

    Returns:
        pd.DataFrame: The filtered and processed dataframe.
        list of str: The updated list of checked columns.
        pd.DataFrame: The dataframe containing holiday dates.
        str: The filename for saving the processed dataframe.

    Raises:
        Exception: If an error occurs during processing, an error message
         is displayed and None is returned
                   for all return values.

    """
    filtered_df["Invoice Number"] = filtered_df["Invoice Number"].astype(str)
    if "Holiday Transactions" in checked_columnsspec:
        if len(checked_columnsspec) == 1:
            filtered_df["Document Date"] = pd.to_datetime(
                filtered_df["Document Date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df["Posting Date"] = pd.to_datetime(
                filtered_df["Posting Date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df["Verified on"] = pd.to_datetime(
                filtered_df["Verified on"], format="%Y/%m/%d", errors="coerce"
            )
            dfholiday["date"] = pd.to_datetime(
                dfholiday["date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df = filtered_df[
                filtered_df["Posting Date"].isin(dfholiday["date"])
            ]
            checked_columnsspec = [
                (
                    "Posting Date"
                    if col == "Holiday Transactions"
                    else (
                        "Invoice Number"
                        if col in ("Inv-Special Character", "80 % Same Invoice")
                        else col
                    )
                )
                for col in checked_columnsspec
            ]
        elif (
            len(checked_columnsspec) == 2
            and "Inv-Special Character" in checked_columnsspec
        ):
            filtered_df["Document Date"] = pd.to_datetime(
                filtered_df["Document Date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df["Posting Date"] = pd.to_datetime(
                filtered_df["Posting Date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df["Verified on"] = pd.to_datetime(
                filtered_df["Verified on"], format="%Y/%m/%d", errors="coerce"
            )
            dfholiday["date"] = pd.to_datetime(
                dfholiday["date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df = filtered_df[
                filtered_df["Posting Date"].isin(dfholiday["date"])
            ]
            filtered_df = filtered_df.sort_values(
                by=["Invoice Number", "Posting Date"])

            grouped = filtered_df.groupby('Reimbursement ID')
            similar_invoices = set()

            for name, group in grouped:
                if len(group) > 1:  # Apply logic only if group size is greater than one
                    filtered_df = filtered_df[~filtered_df.duplicated(subset='Invoice Number', keep=False)]
                    # Create a list of tuples with original and cleaned invoice numbers
                    invoice_pairs = [(invoice, re.sub(r'[^A-Za-z0-9]+', '', str(invoice))) for invoice in
                                     group['Invoice Number']]
                    # Find all unique pairs where invoices are similar within the group
                    for i, (inv1, clean_inv1) in enumerate(invoice_pairs):
                        for inv2, clean_inv2 in invoice_pairs[i + 1:]:
                            if is_similar(clean_inv1, clean_inv2):
                                similar_invoices.add(inv1)
                                similar_invoices.add(inv2)
            checked_columnsspec = [
                (
                    "Posting Date"
                    if col == "Holiday Transactions"
                    else (
                        "Invoice Number"
                        if col in ("Inv-Special Character", "80 % Same Invoice")
                        else col
                    )
                )
                for col in checked_columnsspec
            ]
        elif (
            len(checked_columnsspec) == 2 and "80 % Same Invoice" in checked_columnsspec
        ):
            filtered_df["Document Date"] = pd.to_datetime(
                filtered_df["Document Date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df["Posting Date"] = pd.to_datetime(
                filtered_df["Posting Date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df["Verified on"] = pd.to_datetime(
                filtered_df["Verified on"], format="%Y/%m/%d", errors="coerce"
            )
            dfholiday["date"] = pd.to_datetime(
                dfholiday["date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df = filtered_df[
                filtered_df["Posting Date"].isin(dfholiday["date"])
            ]
            filtered_df["Invoice Number"] = filtered_df["Invoice Number"].astype(
                str)
            filtered_df = optimized_similar_invoices(filtered_df)
            checked_columnsspec = [
                (
                    "Posting Date"
                    if col == "Holiday Transactions"
                    else (
                        "Invoice Number"
                        if col in ("Inv-Special Character", "80 % Same Invoice")
                        else col
                    )
                )
                for col in checked_columnsspec
            ]
        else:
            filtered_df["Document Date"] = pd.to_datetime(
                filtered_df["Document Date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df["Posting Date"] = pd.to_datetime(
                filtered_df["Posting Date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df["Verified on"] = pd.to_datetime(
                filtered_df["Verified on"], format="%Y/%m/%d", errors="coerce"
            )
            dfholiday["date"] = pd.to_datetime(
                dfholiday["date"], format="%Y/%m/%d", errors="coerce"
            )
            filtered_df = filtered_df[
                filtered_df["Posting Date"].isin(dfholiday["date"])
            ]
            filtered_df = filtered_df.sort_values(
                by=["Invoice Number", "Posting Date"])

            # Find similar invoices without grouping
            # Find similar invoices without grouping
            grouped = filtered_df.groupby('Reimbursement ID')
            similar_invoices = set()

            for name, group in grouped:
                if len(group) > 1:  # Apply logic only if group size is greater than one

                    filtered_df = filtered_df[~filtered_df.duplicated(subset='Invoice Number', keep=False)]
                    # Create a list of tuples with original and cleaned invoice numbers
                    invoice_pairs = [(invoice, re.sub(r'[^A-Za-z0-9]+', '', str(invoice))) for invoice in
                                     group['Invoice Number']]
                    # Find all unique pairs where invoices are similar within the group
                    for i, (inv1, clean_inv1) in enumerate(invoice_pairs):
                        for inv2, clean_inv2 in invoice_pairs[i + 1:]:
                            if is_similar(clean_inv1, clean_inv2):
                                similar_invoices.add(inv1)
                                similar_invoices.add(inv2)
            filtered_df["Invoice Number"] = filtered_df["Invoice Number"].astype(
                str)
            filtered_df = optimized_similar_invoices(filtered_df)
    elif "80 % Same Invoice" in checked_columnsspec:
        if len(checked_columnsspec) == 1:
            filtered_df["Invoice Number"] = filtered_df["Invoice Number"].astype(
                str)
            filtered_df = optimized_similar_invoices(filtered_df)
        else:
            filtered_df = filtered_df.sort_values(
                by=["Invoice Number", "Posting Date"])

            # Find similar invoices without grouping
            # Find similar invoices without grouping
            filtered_df = identify_and_return_all_duplicates(filtered_df)
            filtered_df = optimized_similar_invoices(filtered_df)
    elif "Inv-Special Character" in checked_columnsspec:
        if len(checked_columnsspec) == 1:
            grouped = filtered_df.groupby('Reimbursement ID')
            similar_invoices = set()

            for name, group in grouped:
                if len(group) > 1:  # Apply logic only if group size is greater than one
                    filtered_df = filtered_df[~filtered_df.duplicated(subset='Invoice Number', keep=False)]
                    # Create a list of tuples with original and cleaned invoice numbers
                    invoice_pairs = [(invoice, re.sub(r'[^A-Za-z0-9]+', '', str(invoice))) for invoice in
                                     group['Invoice Number']]
                    # Find all unique pairs where invoices are similar within the group
                    for i, (inv1, clean_inv1) in enumerate(invoice_pairs):
                        for inv2, clean_inv2 in invoice_pairs[i + 1:]:
                            if is_similar(clean_inv1, clean_inv2):
                                similar_invoices.add(inv1)
                                similar_invoices.add(inv2)
    checked_columnsspec = [
        (
            "Posting Date"
            if col == "Holiday Transactions"
            else (
                "Invoice Number"
                if col in ("Inv-Special Character", "80 % Same Invoice")
                else col
            )
        )
        for col in checked_columnsspec
    ]
    sort_columns = checked_columnsspec.copy()
    # if (
    #     "Reimbursement ID" not in checked_columnsspec
    #     and "Cost Center" not in checked_columnsspec
    # ):
    #     sort_columns += ["Reimbursement ID", "Cost Center"]
    # elif (
    #     "Reimbursement ID" in checked_columnsspec
    #     and "Cost Center" not in checked_columnsspec
    # ):
    #     sort_columns += ["Cost Center"]
    # elif (
    #     "Cost Center" in checked_columnsspec
    #     and "Reimbursement ID" not in checked_columnsspec
    # ):
    #     sort_columns += ["Reimbursement ID"]
    filtered_df = filtered_df.sort_values(by=sort_columns)
    filtered_df.reset_index(drop=True, inplace=True)
    filtered_df.index += 1
    try:
        return filtered_df, checked_columnsspec, dfholiday, filename

    except pd.errors.EmptyDataError as e:
        st.error(f"An error occurred: {e}")
        return None, None, None, None
