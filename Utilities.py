import pandas as pd

def read_excel(file_name, sheet_name):
    """
    A function to read and return values of an excel sheet and return the values.
    :param file_name:
    :param sheet_name:
    :return:
    """
    datasets = pd.read_excel(file_name, sheet_name = sheet_name)
    return datasets

def read_column(column_name, input_datasets):
    """
    A function to read a particular column values and return the parameters as list.
    :param column_name:
    :return:
    """
    column_values = input_datasets[column_name]
    return column_values

def get_unique_values(column_values):
    """
    A function to find and return only unique elements from a dataset
    :param column_values:
    :return:
    """
    # How ? Convert a list into a set [removes duplicates automatically] and then again into a list
    unique_data = list(set(column_values))
    return unique_data

def create_unique_excels(dataset, column_name, excel_name, excel_sheet_name):
    """
    A function to create unique issue bound excel sheets based on the number of unique issues found
    :param dataset:
    :param column_name:
    :return:
    """

    i = 0
    range_value = len(dataset)

    while i < range_value :

        get_issue_name = dataset[i]
        create_excel_file = get_issue_name + ".xlsx"

        excel_data = read_excel(excel_name, excel_sheet_name)

        issue_found = excel_data[excel_data[column_name] == get_issue_name]

        issue_found.to_excel(create_excel_file)

        i = i + 1

    return
