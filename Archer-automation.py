from Utilities import *

excel_name = 'Archer Dummy Data.xlsx'
excel_sheet_name = 'Archer Search Report'
column_name = 'Finding Name'

# Step 1 : create a function to read an excel file based on dynamic inputs
excel_data = read_excel(excel_name, excel_sheet_name)

# Step 2 : create a function to read a certain column in an excel sheet
excel_column_data = read_column(column_name, excel_data)

# Step 3 : create a function to obtain unique values in a dataset
unique_datasets = get_unique_values(excel_column_data)

# Step 4 : create a function to create excel sheets based on unique issues found
create_unique_excels(unique_datasets, column_name, excel_name, excel_sheet_name)
