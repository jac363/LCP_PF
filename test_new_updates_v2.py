import unittest
import pandas as pd
import os  # Add this line to import the os module
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from new_updates_v2 import compare_and_extract_companies

class TestCompareAndExtractCompanies(unittest.TestCase):
    def setUp(self):
        # Create sample data for testing
        self.pf_table_data = {
            "Company": ["Company A", "Company B", "Company C"],
            "Updated": ["01/01/2022", "05/05/2023", "10/10/2023"]
        }
        self.table_2_data = {
            "机构": ["Company A", "Company B", "Company D"],
            "发布时间": ["2022-01-01", "2023-06-15", "2023-05-01"]
        }
        self.output_file = "Test_Table_3.xlsx"

        # Create sample DataFrames
        self.pf_table = pd.DataFrame(self.pf_table_data)
        self.table_2 = pd.DataFrame(self.table_2_data)

    def test_compare_and_extract_companies(self):
        # Save sample DataFrames to Excel files
        self.pf_table.to_excel("Test_PF_Table.xlsx", index=False)
        self.table_2.to_excel("Test_Table_2.xlsx", index=False)

        # Call the function to be tested
        compare_and_extract_companies("Test_PF_Table.xlsx", "Test_Table_2.xlsx", self.output_file)

        # Load the resulting Table 3 for testing
        table_3 = pd.read_excel(self.output_file)

        # Expected new companies in Table 3
        expected_new_companies = ["Company D"]

        # Check if new companies are in Table 3
        new_companies_in_table_3 = set(table_3["机构"]).intersection(expected_new_companies)
        self.assertEqual(new_companies_in_table_3, set(expected_new_companies))

        # Expected matching companies in Table 3
        expected_matching_companies = ["Company A"]

        # Check if matching companies are in Table 3
        matching_companies_in_table_3 = set(table_3["机构"]).intersection(expected_matching_companies)
        self.assertEqual(matching_companies_in_table_3, set(expected_matching_companies))

    def tearDown(self):
        # Clean up the generated Excel files
        for filename in ["Test_PF_Table.xlsx", "Test_Table_2.xlsx", self.output_file]:
            try:
                os.remove(filename)
            except FileNotFoundError:
                pass

if __name__ == "__main__":
    unittest.main()
