import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

def compare_and_extract_companies(pf_table_file, table_2_file, output_file):
    # Load the "New Investments" sheet from the PF Table Excel file
    pf_table = pd.read_excel(pf_table_file, sheet_name="New Investments", header=0)

    # Load Table 2
    table_2 = pd.read_excel(table_2_file, header=1)

    # Extract company names from PF Table and Table 2
    pf_company_names = set(pf_table["Company"].dropna())
    table_2_company_names = set(table_2["被投公司"].dropna())

    # Identify new company names in Table 2 not in PF Table
    new_company_names = table_2_company_names - pf_company_names

    # Convert "发布时间" column to Pandas Timestamp
    table_2["发布时间"] = pd.to_datetime(table_2["发布时间"])

    # Compare dates and filter matching companies based on date difference
    matching_companies = []
    for company_name in pf_company_names.intersection(table_2_company_names):
        pf_date = pf_table.loc[pf_table["Company"] == company_name, "Updated"].iloc[0]
        table_3_date = table_2.loc[table_2["被投公司"] == company_name, "发布时间"].iloc[0]

        date_difference = relativedelta(table_3_date, pf_date)
        
        if abs(date_difference.months) > 3 or abs(date_difference.years) > 0:
            matching_companies.append(company_name)

    # Filter rows from Table 2 based on companies in Table 3
    table_3_new_companies = table_2[table_2["被投公司"].isin(new_company_names)]
    table_3_matching_companies = table_2[table_2["被投公司"].isin(matching_companies)]

    # Concatenate the two filtered tables
    table_3 = pd.concat([table_3_new_companies, table_3_matching_companies])

    # Save Table 3 to a new Excel file
    table_3.to_excel(output_file, index=False)

# Usage
compare_and_extract_companies("PF Table.xlsx", "Table 2.xlsx", "Table 3.xlsx")
