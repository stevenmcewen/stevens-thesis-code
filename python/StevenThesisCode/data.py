import pandas as pd


# Load all the sheets of the Excel file into a dictionary of dataframes
dfs = pd.read_excel("file.xlsx", sheet_name=None)

# Iterate over the sheets in the dictionary
for sheet_name, df in dfs.items():
  # Add a new column to the dataframe with the sheet name
  df["Sheet Name"] = sheet_name

# Print the dataframe for the "Sheet1" sheet
print(dfs["Sheet1"])