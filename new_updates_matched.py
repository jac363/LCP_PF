import pandas as pd

# Load PF Table
pf_table = pd.read_excel("PF Table.xlsx", sheet_name="New Investments", header=0)

# Load Table 2
table2 = pd.read_excel("Table 2.xlsx", header=1)

# Extract company names from PF Table and Table 2
pf_companies = pf_table["Company"].tolist()
table2_companies = table2["被投公司"].tolist()

# Find matching companies
matching_companies = list(set(pf_companies) & set(table2_companies))

# Create Table 4 with matching rows
table4_rows = table2[table2["被投公司"].isin(matching_companies)]

# Save Table 4 to a new Excel file
table4_rows.to_excel("Table 4.xlsx", index=False, engine="openpyxl")
