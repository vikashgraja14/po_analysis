import pandas as pd
import streamlit as st
import os
import re
import numpy as np
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
    df['Clearing doc no.'].replace('', np.nan, inplace=True)
    #
    # # Drop rows with NaN in 'height' column
    df.dropna(subset=['Clearing doc no.'], inplace=True)
    #
    # # df = df[df['Clearing doc no.'] != '']
    # df = df[df['Clearing doc no.'].apply(lambda x: len(str(x)) > 1)]

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
    df['Cummulative_transactions'] = len(df)
    df['Cummulative_transactions/category'] = df.groupby('category')['category'].transform('count')
    df['overall_transactions/year'] = df.groupby('year')['year'].transform('count')
    df['overall_transactions/category/year'] = df.groupby(['category', 'year'])['category'].transform('count')
    df['cumulative_Transations/Vendor'] = df.groupby('Vendor')['Vendor'].transform('count')
    df['Transations/year/Vendor'] = df.groupby(['Vendor', 'year', 'category'])['Vendor'].transform('count')
    df['Cumulative_percentransations_made'] = (df['cumulative_Transations/Vendor'] / df[
        'Cummulative_transactions']) * 100
    df['Cumulative_percentransations_made/category'] = (df['cumulative_Transations/Vendor'] / df[
        'Cummulative_transactions/category']) * 100
    df['Yearly_percentransations_made/category'] = (df['overall_transactions/category/year'] / df[
        'overall_transactions/year']) * 100
    df['percentransations_made/category/year'] = (df['Transations/year/Vendor'] / df[
        'overall_transactions/category/year']) * 100
    df['percentransations_made/year'] = (df['Transations/year/Vendor'] / df['overall_transactions/year']) * 100
    df['cumulative_Alloted_Amount'] = df['Amount'].sum()
    df['cumulative_Alloted_Amount\Category'] = df.groupby('category')['Amount'].transform('sum')
    yearly_total = df.groupby('year')['Amount'].sum().reset_index()
    yearly_total.rename(columns={'Amount': 'Total_Alloted_Amount/year'}, inplace=True)
    df = pd.merge(df, yearly_total, on='year', how='left')
    df['Yearly_Alloted_Amount\Category'] = df.groupby(['category', 'year'])['Amount'].transform('sum')
    df['Cumulative_Amount_used'] = df.groupby(['Vendor', 'category'])['Amount'].transform('sum')
    df['Amount_used/Year'] = df.groupby(['Vendor', 'year', 'category'])['Amount'].transform('sum')
    df['Cumulative_percentageamount_used'] = (df['Cumulative_Amount_used'] / df['cumulative_Alloted_Amount']) * 100
    df['total_percentage_of_amount/category_used'] = (df['Cumulative_Amount_used'] / df[
        'cumulative_Alloted_Amount\Category']) * 100
    df['percentage_amount_used_per_year'] = (df['Amount_used/Year'] / df['Total_Alloted_Amount/year']) * 100
    df['percentage_of_amount/category_used/year'] = (df['Amount_used/Year'] / df[
        'Yearly_Alloted_Amount\Category']) * 100
    df['percentage_Yearly_Alloted_Amount\Category'] = (df['Yearly_Alloted_Amount\Category'] / df[
        'Total_Alloted_Amount/year']) * 100
    df['Cumulative_transactions/Cost Ctr'] = df.groupby(['Cost Ctr'])['Cost Ctr'].transform('count')
    df['Cumulative_transactions/Cost Ctr/Year'] = df.groupby(['Cost Ctr','year'])['Cost Ctr'].transform('count')
    df['Cumulative_Alloted/Cost Ctr'] = df.groupby(['Cost Ctr'])['Amount'].transform('sum')
    df['Cumulative_Alloted/Cost Ctr/Year'] = df.groupby(['Cost Ctr','year'])['Amount'].transform('sum')
    df['Percentage_Cumulative_Alloted/Cost Ctr'] = (df['Cumulative_Alloted/Cost Ctr'] / df['cumulative_Alloted_Amount']) * 100
    df['Percentage_Cumulative_Alloted/Cost Ctr/Year'] = (df['Cumulative_Alloted/Cost Ctr/Year'] / df['Total_Alloted_Amount/year']) * 100
    # df['Cumulative_transactions/G/L'] = df.groupby(['G/L'])['G/L'].transform('count')
    # df['Cumulative_transactions/G/L/Year'] = df.groupby(['G/L', 'G/L'])['G/L'].transform('count')
    # df['Cumulative_Alloted/G/L'] = df.groupby(['G/L'])['Amount'].transform('sum')
    # df['Cumulative_Alloted/G/L/Year'] = df.groupby(['G/L', 'year'])['Amount'].transform('sum')
    # df['Percentage_Cumulative_Alloted/G/L'] = (df['Cumulative_Alloted/G/L'] / df[
    #     'cumulative_Alloted_Amount']) * 100
    # df['Percentage_Cumulative_Alloted/G/L/Year'] = (df['Cumulative_Alloted/G/L/Year'] / df[
    #     'Total_Alloted_Amount/year']) * 100
    df['used_amount_crores'] = (df['Yearly_Alloted_Amount\Category']/10000000)
    df['percentage Transcation/costctr/year'] =df['Cumulative_transactions/Cost Ctr/Year'] /df['Cummulative_transactions']*100
    # df['percentage Transcation/G/L/year'] = df['Cumulative_transactions/G/L/Year'] / df[
    #     'Cummulative_transactions'] * 100
    df.reset_index(drop=True, inplace=True)
    df['Vendor'] = df['Vendor'].astype(str)
    df['G/L'] = df['G/L'].astype(str)
    df['G/L'] = df['G/L'].apply(lambda x: str(x) if isinstance(x, str) else '')
    df['G/L'] = df['G/L'].apply(lambda x: re.sub(r'\..*', '', x))
    df = df.sort_values(by='percentage_of_amount/category_used/year', ascending=False)
    df.reset_index(drop=True, inplace=True)
    return df
