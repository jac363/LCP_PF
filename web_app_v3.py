# Execute Merge_Tables & Format_Table & Create_PF_Table correctly

import streamlit as st
import pandas as pd
import os
from datetime import date

# Function to merge the tables and remove duplicates
def merge_tables(file1, file2):
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
    # Create an empty Peer Funds Table DataFrame
    peer_funds_table = pd.DataFrame(columns=['Category', 'Updated', 'Company', 'Business', 'Peer Fund', 'Round', 'Amount',
                                             '城市', '是否值得跟进', '跟进人 & Deallog', '跟进记录', '是否值得考虑一下轮', 'Funding History',
                                             'Notes', 'Due Date', '工商名称'])

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
            '工商名称': row['工商名称']
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


# Modify the Streamlit section
def main():
    # Set the title and description of the app
    st.title("Peer Funds Table Generator")
    st.subheader("Upload three Excel files to merge, format, and complete peer funds table")

    # Allow the user to upload the three Excel files
    file1 = st.file_uploader("请上载烯牛数据导出PF1-多条合并 ", type=["xlsx"])
    file2 = st.file_uploader("请上载烯牛数据导出PF2-多条合并", type=["xlsx"])
    file3 = st.file_uploader("请上载企名片导出", type=["xlsx"])

    # Check if all three files are uploaded
    if file1 and file2 and file3:
        # Execute your Python code with the uploaded files
        merged_table = merge_tables(file1, file2)
        formatted_table = format_table(merged_table)
        peer_funds_table = create_pf_table(formatted_table)

        # Save the Peer Funds Table as an Excel file
        peer_funds_table_file = "peer_fund_table.xlsx"
        peer_funds_table.to_excel(peer_funds_table_file, index=False)

        # Complete Peer Funds Table with notes from Table 3
        completed_table_file = complete_pf_notes(peer_funds_table_file, file3)

        # Provide a download link for the completed table file
        with open(completed_table_file, "rb") as f:
            file_bytes = f.read()
        st.subheader("下载完整的Peer Funds Table")
        st.download_button("点击这里", file_bytes, file_name="peer_fund_table_with_notes.xlsx", mime="application/octet-stream")


# Run the Streamlit app
if __name__ == "__main__":
    main()