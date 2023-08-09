"""
Title: Company Comparison and Extraction Script

Description:
This script compares company names between two Excel files, "PF Table.xlsx" and "Table 2.xlsx". It loads the "New Investments" sheet from the "PF Table" Excel file, extracts company names from both tables, identifies new companies that exist in "Table 2" but not in "PF Table", and generates a new Excel file "Table 3.xlsx" containing information about these new companies.

Dependencies:
- pandas: Data manipulation library for Python. Install using 'pip install pandas'

"""

import pandas as pd


# Load the "New Investments" sheet from the PF Table Excel file
pf_table = pd.read_excel("PF Table.xlsx", sheet_name="New Investments", header=0)

# Load Table 2
table_2 = pd.read_excel("Table 2.xlsx", header=1)

# Extract company names from PF Table and Table 2
pf_company_names = set(pf_table["Company"].dropna())
table_2_company_names = set(table_2["机构"].dropna())

# Identify new company names
new_company_names = table_2_company_names - pf_company_names

# Filter rows from Table 2 based on new company names
table_3 = table_2[table_2["机构"].isin(new_company_names)]

# Save Table 3 to a new Excel file
table_3.to_excel("Table 3.xlsx", index=False)
