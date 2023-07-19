# This code is a Streamlit web application that provides two functionalities:
# 1. Peer Funds Table Generator: It allows users to upload four Excel files, merge, format, and complete a peer funds table.
# 2. Missing Companies Sorter: It allows users to upload three Excel files and extract missing companies from the data.

import streamlit as st
import pandas as pd
import os
from datetime import date
import re

# Function to merge the tables and remove duplicates
def merge_tables(file1, file2):
    """
    Read two Excel files and merge the tables while removing duplicates.

    Args:
        file1 (str): Path to the first Excel file.
        file2 (str): Path to the second Excel file.

    Returns:
        pandas.DataFrame: Merged table without duplicate rows.
    """
    # Read Table 1 from the first Excel file
    table1 = pd.read_excel(file1)

    # Extract data from Table 1
    header_row_index = 1
    empty_row_index = table1.iloc[:, 0].isna().idxmax()
    table1_data = table1.iloc[header_row_index:empty_row_index]

    # Read Table 2 from the second Excel file
    table2 = pd.read_excel(file2, header=None, skiprows=2)

    # Extract data from Table 2
    empty_row_index = table2.iloc[:, 0].isna().idxmax()
    table2_data = table2.iloc[:empty_row_index]

    # Rename headers in Table 1
    table1_data.columns = ["序号", "公司名", "简介", "烯牛行业（一级）", "成立时间", "地区", "最新融资时间", "最新融资轮次", "最新融资金额", "投资方", "融资历程（多条）", "工商名称", "联系电话"]

    # Reset the index of Table 2
    table2_data.reset_index(drop=True, inplace=True)

    # Calculate the number of columns to shift Table 2
    num_columns_to_shift = table1_data.shape[1]

    # Shift Table 2 to start in the same column as Table 1
    table2_data.columns = table1_data.columns
    table2_data = table2_data.shift(0, axis=1)

    # Merge Table 1 and Table 2
    merged_table = pd.concat([table1_data, table2_data])

    # Find duplicate rows based on the '公司名' column
    duplicate_rows = merged_table[merged_table.duplicated(subset='公司名', keep='first')]

    # Keep only the first occurrence of each duplicate company
    merged_table.drop_duplicates(subset='公司名', keep='first', inplace=True)

    # Reset the index of the merged table
    merged_table.reset_index(drop=True, inplace=True)

    # Update the values in the "序号" column
    merged_table['序号'] = merged_table.index + 1

    # Remove the word "市" and the words after it under the "地区" column
    merged_table['地区'] = merged_table['地区'].str.replace('市.+', '', regex=True)

    # Return the merged table
    return merged_table

# Function to format the merged table
def format_table(merged_table):
    """
    Format the merged table by renaming columns and cleaning up values.

    Args:
        merged_table (pandas.DataFrame): Merged table to be formatted.

    Returns:
        pandas.DataFrame: Formatted table.
    """
    # Rename the columns for easier access
    merged_table.columns = ['序号', '公司名', '简介', '烯牛行业（一级）', '成立时间', '地区', '最新融资时间', '最新融资轮次', '最新融资金额',
                            '投资方', '融资历程（多条）', '工商名称', '联系电话']

    # Remove the phrase "金额" from the "融资历程（多条)" column
    merged_table["融资历程（多条）"] = merged_table["融资历程（多条）"].str.replace("金额：", "")

    # Replace "、" with "/"
    merged_table["融资历程（多条）"] = merged_table["融资历程（多条）"].str.replace("、", "/")

    # Replace "，" with "/"
    merged_table["融资历程（多条）"] = merged_table["融资历程（多条）"].str.replace("，", "/")

    # Replace delimiter "," with " – "
    merged_table["融资历程（多条）"] = merged_table["融资历程（多条）"].str.replace(",", " – ")

    # Replace "未披露" with "N/A"
    merged_table["融资历程（多条）"] = merged_table["融资历程（多条）"].str.replace("未披露", "N/A")

    # Return the formatted table
    return merged_table

# Function to create the Peer Funds Table
def create_pf_table(formatted_table):
    """
    Create the Peer Funds Table based on the formatted table.

    Args:
        formatted_table (pandas.DataFrame): Formatted table.

    Returns:
        pandas.DataFrame: Peer Funds Table.
    """
    # Create an empty Peer Funds Table DataFrame
    peer_funds_table = pd.DataFrame(columns=['Category', 'Updated', 'Company', 'Business', 'Peer Fund', 'Round', 'Amount',
                                             '城市', '是否值得跟进', '跟进人 & Deallog', '跟进记录', '是否值得考虑一下轮', 'Funding History',
                                             'Notes', 'Due Date', '工商名称', '投资方'])

    # Iterate over each row in the formatted table
    for _, row in formatted_table.iterrows():
        # Extract data from the formatted table and populate the Peer Funds Table
        data = {
            'Category': row['烯牛行业（一级）'],
            'Company': row['公司名'],
            'Business': row['简介'],
            'Round': row['最新融资轮次'],
            'Amount': row['最新融资金额'],
            '城市': row['地区'],
            'Funding History': row['融资历程（多条）'],
            '工商名称': row['工商名称'],
            '投资方': row['投资方']
        }
        peer_funds_table = pd.concat([peer_funds_table, pd.DataFrame(data, index=[0])], ignore_index=True)

    # Fill the "Updated" column with today's date
    peer_funds_table['Updated'] = date.today()
    
    # In each cell of the "Category" column, delete content after ","
    peer_funds_table['Category'] = peer_funds_table['Category'].str.split('，').str[0]

    # In each cell of the "城市" column, delete "市" and the content after "市"
    peer_funds_table['城市'] = peer_funds_table['城市'].str.replace('市.+', '', regex=True)

    # Replace "未披露" with "N/A" in the "Amount" column
    peer_funds_table['Amount'] = peer_funds_table['Amount'].replace("未披露", "N/A")

    # Return the Peer Funds Table
    return peer_funds_table

# New function to complete Peer Funds Table with company descriptions from Table 3
def complete_pf_notes(peer_funds_table_file, table3_file):
    """
    Complete the Peer Funds Table with company descriptions from Table 3.

    Args:
        peer_funds_table_file (str): Path to the Peer Funds Table Excel file.
        table3_file (str): Path to the Table 3 Excel file.

    Returns:
        str: Path to the output Excel file with the completed Peer Funds Table.
    """
    # Read the Peer Funds Table and Table 3 Excel files into pandas DataFrames
    pf_table = pd.read_excel(peer_funds_table_file)
    table3 = pd.read_excel(table3_file, skiprows=1)

    # Find the matching companies between Peer Funds Table and Table 3
    matching_companies = pf_table['工商名称'].isin(table3['公司名称'])

    # Iterate over matching companies and copy the company descriptions
    for idx, row in pf_table[matching_companies].iterrows():
        company_name = row['工商名称']
        company_description = table3.loc[table3['公司名称'] == company_name, '简介'].values
        if len(company_description) > 0:
            pf_table.loc[idx, 'Notes'] = company_description[0]

    # Fill the "Updated" column with the current date
    pf_table['Updated'] = date.today().strftime('%Y/%m/%d')

    # Remove the "工商名称" column
    pf_table.drop('工商名称', axis=1, inplace=True)

    # Override the current Excel table with the updated DataFrame
    output_file = 'peer_fund_table_with_notes.xlsx'
    pf_table.to_excel(output_file, index=False)
    return output_file


# New function to complete Peer Funds Table with investor names from Table 4
def complete_pf_investors(pf_table_file, table4_file):
    """
    Complete the Peer Funds Table with investor names from Table 4.

    Args:
        pf_table_file (str): Path to the Peer Funds Table Excel file.
        table4_file (str): Path to the Table 4 Excel file.

    Returns:
        str: Path to the output Excel file with the completed Peer Funds Table.
    """
    # Read the Peer Funds Table and Table 4 Excel files into pandas DataFrames
    pf_table = pd.read_excel(pf_table_file)
    table4 = pd.read_excel(table4_file, header=None)

    # Remove "领投" and "跟投" from the "投资方" column in pf_table
    pf_table['投资方'] = pf_table['投资方'].str.replace('领投', '').str.replace('跟投', '')

    # Define a regular expression pattern to split the fund names using delimiters ",", "、", "，"
    delimiter_pattern = r'[，,、]'

    # Iterate through each cell in the "投资方" column
    for index, row in pf_table.iterrows():
        investment_firms = re.split(delimiter_pattern, str(row['投资方'])) if isinstance(row['投资方'], str) else []
        tracked_firms = [firm for firm in investment_firms if any(firm in tracked for tracked in table4.values.flatten())]
        pf_table.at[index, 'Peer Fund'] = '/'.join(tracked_firms)

    # Remove the "投资方" column
    pf_table.drop('投资方', axis=1, inplace=True)

    # Save the modified table to a new Excel file
    output_file = 'completed_peer_fund_table.xlsx'
    pf_table.to_excel(output_file, index=False)
    return output_file

# First page: Peer Funds Table Generator
def peer_funds_table_generator():
    """
    Streamlit page for the Peer Funds Table Generator functionality.
    """
    # Set the title and description of the app
    st.title("Peer Funds Table Generator")
    st.subheader("Merge, format, and complete peer funds table")

    # Allow the user to upload the four Excel files
    file1 = st.file_uploader("请上载烯牛数据导出PF1-多条合并", type=["xlsx"])
    file2 = st.file_uploader("请上载烯牛数据导出PF2-多条合并", type=["xlsx"])
    file3 = st.file_uploader("请上载企名片导出", type=["xlsx"])
    file4 = st.file_uploader("请上载PF Tracked List", type=["xlsx"])

    # Check if all four files are uploaded
    if file1 and file2 and file3 and file4:
        # Execute your Python code with the uploaded files
        merged_table = merge_tables(file1, file2)
        formatted_table = format_table(merged_table)
        peer_funds_table = create_pf_table(formatted_table)
        peer_funds_table_file = "peer_fund_table.xlsx"
        peer_funds_table.to_excel(peer_funds_table_file, index=False)
        completed_table_file = complete_pf_notes(peer_funds_table_file, file3)
        output_file = complete_pf_investors(completed_table_file, file4)

        # Provide a download link for the completed table file
        with open(output_file, "rb") as f:
            file_bytes = f.read()
        st.subheader("下载完整的 Peer Funds Table")
        st.download_button("点击这里", file_bytes, file_name="completed_peer_fund_table.xlsx", mime="application/octet-stream")

# New function to extract missing companies from Table 3
def extract_missing_companies(table1, table2, table3):
    """
    Extract missing companies from Table 3 that are not present in Table 1 and Table 2.

    Args:
        table1 (pandas.DataFrame): Table 1 data.
        table2 (pandas.DataFrame): Table 2 data.
        table3 (pandas.DataFrame): Table 3 data.

    Returns:
        str: Path to the output Excel file with the missing companies.
    """
    # Save DataFrame objects to temporary Excel files
    table1_file = "table1_temp.xlsx"
    table1.to_excel(table1_file, index=False)
    table2_file = "table2_temp.xlsx"
    table2.to_excel(table2_file, index=False)
    table3_file = "table3_temp.xlsx"
    table3.to_excel(table3_file, index=False)

    # Merge Table1 and Table2
    merged_table = merge_tables(table1_file, table2_file)

    # Get the company names from the merged table
    merged_companies = set(merged_table['工商名称'].tolist())

    # Filter Table3 to find missing companies
    missing_companies = table3[~table3['公司名称'].isin(merged_companies)]

    # Write the missing companies to a new Excel file
    missing_companies_file = "missing_companies.xlsx"
    missing_companies.to_excel(missing_companies_file, index=False)

    # Remove the temporary Excel files
    os.remove(table1_file)
    os.remove(table2_file)
    os.remove(table3_file)

    return missing_companies_file

# Second page: Missing Companies Sorter
def missing_companies_sorter():
    """
    Streamlit page for the Missing Companies Sorter functionality.
    """
    # Set the title and description of the page
    st.title("Missing Companies Sorter")
    st.subheader("Extract missing companies from 犀牛数据")

    # Allow the user to upload the three Excel files
    table1_file = st.file_uploader("请上载 烯牛数据导出PF1-多条合并", type=["xlsx"])
    table2_file = st.file_uploader("请上载 烯牛数据导出PF2-多条合并", type=["xlsx"])
    table3_file = st.file_uploader("请上载 企名片导出", type=["xlsx"])

    # Check if all three files are uploaded
    if table1_file and table2_file and table3_file:
        # Read the uploaded files as pandas DataFrames
        table1 = pd.read_excel(table1_file)
        table2 = pd.read_excel(table2_file, skiprows=1)
        table3 = pd.read_excel(table3_file, skiprows=1)

        # Execute your Python code with the uploaded files
        missing_companies_file = extract_missing_companies(table1, table2, table3)

        # Provide a download link for the missing companies file
        with open(missing_companies_file, "rb") as f:
            file_bytes = f.read()
        st.subheader("下载表格")
        st.download_button("点击这里", file_bytes, file_name="missing_companies.xlsx", mime="application/octet-stream")

# Main function to run the app
def main():
    """
    Main function to run the Streamlit app.
    """
    # Set the title of the app
    st.title("Web App")

    # Create a navigation system between the pages
    page = st.radio("Select a page:", ("Peer Funds Table Generator", "Missing Companies Sorter"))

    # Display the selected page
    if page == "Peer Funds Table Generator":
        peer_funds_table_generator()
    elif page == "Missing Companies Sorter":
        missing_companies_sorter()

# Run the Streamlit app
if __name__ == "__main__":
    main()
