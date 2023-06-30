import pandas as pd
import matplotlib.pyplot as plt

# Replace the URL with the direct link to your Google Sheets file
url = "https://docs.google.com/spreadsheets/d/1Jlsk55vlCu2vaweiUCOHxQfVG_YZdtqg_do0xAP4qNQ/export?format=xlsx"

# Read the Excel file with multiple sheets
xls = pd.ExcelFile(url)

# Create an empty dictionary to store DataFrames for each sheet
dfs = {}

# Loop through each sheet and load the data into DataFrames
for sheet_name in xls.sheet_names:
    dfs[sheet_name] = pd.read_excel(xls, sheet_name)

# Select the column you want to compare
column_to_compare = "P"

# Create a DataFrame to store the comparison results
comparison_results = pd.DataFrame()

# Loop through the sheets and compare the selected column
for sheet_name, df in dfs.items():
    if column_to_compare in df.columns:
        comparison_results[sheet_name] = df[column_to_compare]

# Print the comparison results
print(comparison_results)
comparison_results.plot()
plt.show()