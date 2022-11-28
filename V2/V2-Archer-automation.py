from Utilities import *

# excel workbook names
excel_name = 'Automation Data v0.1.xlsx'
remed_name = 'Automation_Remediation_worksheet.xlsx'

# excel sheet names
main_excel_sheet_name = 'Automation Data'
keyword_excel_sheet_name = 'Keywords'
remed_sheet_name = 'Sheet1'

# excel column names
main_column_name = 'NUM-Name'

# check if the specific excel file exists
excel_exists = if_file_exists(excel_name)
remed_exists = if_file_exists(remed_name)

# get the target folder name
folder_name = folder_name()
remed_folder = str(folder_name) + "/Remediation-plans"

# check if the main findings excel file exists
if excel_exists != True:

    func_bar(100, "Searching for the main excel file...")
    print(" The specified Findings-Excel file does not exist.")
    print(" Program exiting now.")
    sys.exit()

# check if the remediation excel file exists
elif remed_exists != True:

    func_bar(100, "Searching for the remediation excel file...")
    print(" The specified Remediation Excel file does not exist.")
    print(" Program exiting now.")
    sys.exit()

# if both excel files exist, then initate the sequence
else :

    print("The specified file is found.")
    func_bar(100, "Data Analysis is initiated...")

    # get the whole data from the full excel files
    main_excel_data = pd.read_excel(excel_name, sheet_name=main_excel_sheet_name)
    keyword_excel_data = pd.read_excel(excel_name, sheet_name=keyword_excel_sheet_name)
    remed_excel_data = pd.read_excel(remed_name, sheet_name=remed_sheet_name)

    # get the column headers for traversing values
    keywords_headers = get_column_headers(excel_name, keyword_excel_sheet_name)

    # create the target directory
    os.makedirs(folder_name)
    os.makedirs(remed_folder)

    i = 0

    while i < len(keywords_headers):

        # invoke the function to create keyword-based excel sheets
        create_unique_excels(keywords_headers, keyword_excel_data, main_excel_data, main_column_name, i, folder_name, remed_folder, remed_excel_data)

        # increment i for while-loop ending
        i = i + 1

