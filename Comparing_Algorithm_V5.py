# COMPARING ALGORITHM - MAIN FILE_V2 - ARCHITECTONIC - </> by ShivaReddy.

import pandas as pd

######
# PART 1 - (START) - GET THE DATA, & CONVERT IT INTO DICTIONARIES
# PART 1.1 - CHOOSE THE TWO FILES
file_1 = pd.ExcelFile(r'C:\Users\sreddy\Desktop\testing\TestV1.xlsm')
file_2 = pd.ExcelFile(r'C:\Users\sreddy\Desktop\testing\TestV2.xlsm')
file1_image = r''
file2_image = r''


# PART 1.2 - GET ALL THE SHEET NAMES FROM THE EXCEL FILES
def sheet_names(location_1, location_2):
    pd_file_1 = pd.ExcelFile(location_1)
    pd_file_2 = pd.ExcelFile(location_2)
    names_of_sheets = pd_file_1.sheet_names
    # names_of_sheet2 = pd_file_2.sheet_names
    names_of_sheets.remove(names_of_sheets[0])
    return names_of_sheets


def create_data_frames(input_file_1, input_file_2, sheet_name):
    df1_1 = pd.read_excel(input_file_1, sheet_name, header=None)
    df1_1 = df1_1.drop([0, 1, 2, ((df1_1.shape[0]) - 1)])
    df1_1 = df1_1.where(pd.notnull(df1_1), "(none)")
    f1_s1 = df1_1.to_dict(orient='records')

    df2_1 = pd.read_excel(input_file_2, sheet_name, header=None)
    df2_1 = df2_1.drop([0, 1, 2, ((df2_1.shape[0]) - 1)])
    df2_1 = df2_1.where(pd.notnull(df2_1), "(none)")
    f2_s1 = df2_1.to_dict(orient='records')

    return f1_s1, f2_s1


# excel_sheet_number = 7


def all_sheets_data_frames(input_file_1, input_file_2, excel_workbook_number):
    i = excel_workbook_number
    list_of_all_sheet_names = sheet_names(input_file_1, input_file_2)
    file1_data_frames = []
    file2_data_frames = []
    data_frames_for_2versions_single_sheet = create_data_frames(input_file_1, input_file_2, list_of_all_sheet_names[i])
    file1_data_frames.append(data_frames_for_2versions_single_sheet[0])
    file2_data_frames.append(data_frames_for_2versions_single_sheet[1])
    return file1_data_frames, file2_data_frames

# PART 1 - (END) - GET THE DATA, & CONVERT IT INTO DICTIONARIES
######


#####
# PART 2 - (START) - COMPARING ALGORITHM
# PART 2.0 - CREATE INPUTS AND OUTPUTS FOR PART 2
new_data_2 = []


# PART 2.1 - FUNCTIONS FOR FINDING SAME ITEMS IN TWO LISTS - LIST_A & LIST_B
# PART 2.1.1 - EVERY ITEM IN LIST_A AGAINST ONE ITEM IN LIST_B
def compare_single_item_to_a_list(list_a, item_of_b, return_same_items_list, ):
    for current_item in list_a:
        if item_of_b == current_item:
            return_same_items_list.append(current_item)
            list_a[list_a.index(current_item)] = "same_removed"
            new_data_2.append(item_of_b)
            item_of_b = "same_dict_removed"
        else:
            pass


# PART 2.1.2 - REPEATING 2.1.1 FOR EVERY ITEM IN LIST_B
def compare_lists_for_same_items(first_list, second_list, return_same_items_list):
    for current_item in second_list:
        compare_single_item_to_a_list(first_list, current_item, return_same_items_list)


# PART 2.1.3 - REMOVING REPLACED ITEMS FROM LIST
def remove_items_from_list(input_list):
    for i in input_list[:]:
        if i == "same_removed":
            input_list.remove(i)
        else:
            pass


# PART 2.1.4 - REMOVING REPLACED ITEMS FROM LIST
def remove_items_from_list_type2(input_list):
    for i in input_list[:]:
        if i == "same_dict_removed":
            input_list.remove(i)
        else:
            pass


# PART 2.1.5 - RUNNING LIST COMPARISON FOR TWO LISTS, DELETING REPLACED ITEMS IN BOTH LISTS
def compare_same_items(list_a, list_b, return_same_items_list):
    compare_lists_for_same_items(list_a, list_b, return_same_items_list)
    remove_items_from_list(list_a)
    remove_items_from_list_type2(list_b)


# PART 2.1.6 - COMPARING THE DATA SETS FOR SAME ITEMS
# Function to remove item_b from list_b, after it is used to compare
def remove_same_items_from_data2(input_old_data2, input_new_data2):
    temp_data2 = input_old_data2

    for x in input_new_data2[:]:
        if x in temp_data2:
            temp_data2.remove(x)

    return temp_data2


# PART 2.2 - MODIFIED ITEMS
# PART 2.2.0 - FUNCTION TO COMPARE VALUES OF TWO DICTIONARIES, NOTE INDICES, GROUP INTO SUB-LISTS
def indices_of_modified(dict_a, dict_b, out_modified_indices, the_total_columns):
    for x in range(0, the_total_columns):
        if (list(dict_a.values())[x]) == (list(dict_b.values())[x]):
            out_modified_indices.append(True)
            pass
        else:
            out_modified_indices.append(False)
    return out_modified_indices


# PART 2.2.1 - EVERY DICTIONARY ITEM IN LIST_A VS. ONE DICTIONARY ITEM IN LIST_B
def compare_single_dict_to_a_list(list_a, item_of_b, return_modified_items_list, out_modified_indices, the_total_columns):
    for current_item in list_a:
        try:
            if ((list(current_item.values())[0]) == (list(item_of_b.values())[0])) & (item_of_b != current_item):
                indices_of_modified(current_item, item_of_b, out_modified_indices, the_total_columns)
                return_modified_items_list.append(item_of_b)
                list_a[list_a.index(current_item)] = "removed the dict"
                new_data_2.append(item_of_b)
                item_of_b = "removed_item_in_dict_b"
            else:
                pass
        except AttributeError:
            pass


# PART 2.2.2 - REPEATING 2.2.1 FOR EVERY DICTIONARY ITEM IN LIST_B
def compare_lists_for_modified_items(first_list, second_list, return_modified_items_list, out_modified_items, the_total_columns):
    for current_item in second_list:
        compare_single_dict_to_a_list(first_list, current_item, return_modified_items_list, out_modified_items, the_total_columns)


# PART 2.2.3 - REMOVE REPLACEMENTS IN LIST_A
def remove_dicts_from_list(input_list):
    for i in input_list[:]:
        if i == "removed the dict":
            input_list.remove(i)


# PART 2.2.4 - REMOVE REPLACEMENTS IN LIST_B
def remove_dicts_from_list_type2(input_list):
    for i in input_list[:]:
        if i == "removed_item_in_dict_b":
            input_list.remove(i)


# PART 2.2.5 - RUNNING THE COMPARISON FUNCTIONS FOR TWO DICTIONARY LISTS &  DELETING THE REPLACED ITEMS
def compare_modified_items(list_a, list_b, return_modified_items_list, out_modified_items, the_total_items):
    # tot_columns = len(list(original_data1[0].keys()))
    compare_lists_for_modified_items(list_a, list_b, return_modified_items_list, out_modified_items, the_total_items)
    remove_dicts_from_list(list_a)
    remove_dicts_from_list_type2(list_b)


# PART 2.3 - NEW ITEMS
def remove_same_and_modified_items(input_same_items, input_modified_items, original_data_2_list):
    temp_combined_list = input_same_items + input_modified_items
    temp_original_list = original_data_2_list

    for x in temp_combined_list[:]:
        if x in temp_original_list:
            temp_original_list.remove(x)

    return temp_original_list


# PART 2.4 - RUNNING THE ALGORITHM FOR ALL THE SHEETS
total_number_of_sheets = len(file_1.sheet_names)


def comparing_algorithm(excel_file_1, excel_file_2):
    combined_output_same = []
    combined_output_modified = []
    combined_output_modified_indices = []
    combined_output_new = []
    combined_output_deleted = []


    for x in range(0, total_number_of_sheets-1):
        data1 = all_sheets_data_frames(excel_file_1, excel_file_2, x)[0][0]
        data2 = all_sheets_data_frames(excel_file_1, excel_file_2, x)[1][0]
        original_data1 = all_sheets_data_frames(excel_file_1, excel_file_2, x)[0][0]
        original_data2 = all_sheets_data_frames(excel_file_1, excel_file_2, x)[1][0]
        number_of_columns = len(list(original_data1[0].keys()))

        same_items = []
        all_sheets_data_frames(excel_file_1, excel_file_2, x)
        compare_same_items(data1, data2, same_items)
        data2 = remove_same_items_from_data2(data2, new_data_2)

        modified_items = []
        total_columns = len(list(original_data1[0].keys()))
        modified_indices = []

        compare_modified_items(data1, data2, modified_items, modified_indices, total_columns)

        new_items = remove_same_and_modified_items(same_items, modified_items, original_data2)

        final_output_same = pd.DataFrame(same_items)
        final_output_modified = pd.DataFrame(modified_items)
        final_output_indices = modified_indices
        new_modified_indices = [modified_indices[i:i+total_columns]for i in range(0, len(modified_indices), total_columns)]
        final_output_new = pd.DataFrame(new_items)
        final_output_deleted = pd.DataFrame(data1)

        combined_output_same.append(final_output_same)
        combined_output_modified.append(final_output_modified)
        combined_output_modified_indices.append(new_modified_indices)
        combined_output_new.append(final_output_new)
        combined_output_deleted.append(final_output_deleted)

    return combined_output_same, combined_output_modified, combined_output_modified_indices, combined_output_new, combined_output_deleted


# comparing_algorithm(file_1, file_2)
print(comparing_algorithm(file_1, file_2))

#####
# PART 2 - (END) - COMPARING ALGORITHM

#####
# PART 3 - (START) - SENDING DATA TO EXCEL


#####
# PART 3 - (END) - SENDING DATA TO EXCEL
