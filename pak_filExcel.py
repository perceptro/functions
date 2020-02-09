import openpyxl
import time

path_template_excel = r"C:\Users\RostPK\Desktop\spec\template.xlsx"
path_save_excel = r"C:\Users\RostPK\Desktop\spec\result"
number_table_for_write = 1  # Start serial number for recording in excel file
number_row_for_write = 2  # Start position for write row

# define columns for write in excel. Variable = list_dict_atr[{}]
dict_default_columns_excel = {
    "number_table_for_write": "A",
    "table_name": "B",
    "number_row": "C",
    "name_atr": "D",
    "comment": "E",
    "type": "F",
    "scale": "G",
    "precision": "H"
}


# excel_index
#


def constructorFillExcel(file_error, file_name_struct, date_for_name_file, number_SS):
    # print(file_name_error)
    # print(file_name_struct)
    # print(date_for_name_file)

    # book = openpyxl.Workbook()
    global number_table_for_write
    global path_save_excel
    global path_template_excel
    global number_row_for_write
    global dict_default_columns_excel

    try:
        book = openpyxl.load_workbook(path_template_excel)
    except Exception as error:
        print("Could not open file template.xlsx", error)
        file_error.write("Could not open file template.xlsx:" + str(error) + "\n")
        return

    arr_file_struct = []
    try:
        with open(file_name_struct) as file_struct:
            for line in file_struct.readlines():
                arr_file_struct.append(line.strip())
    except Exception as error:
        print("Error opening file: ", error)

    # print(arr_file_struct[0]) # row_from_table_list|!|2|!!|table_name|!|ekhd.tables1first|!!|number_row|!|number_row1|!!|name_atr|!|atr1|!!|type|!|int|!!|scale|!|12345|!!|precision|!|0|!!|comment|!|comments  for atr1|!!|

    list_dict_atr = []
    for row in arr_file_struct:
        temp = row.split(
            "|!!|")  # ['row_from_table_list|!|2', 'table_name|!|ekhd.tables1first', 'number_row|!|number_row1', 'name_atr|!|atr1', 'type|!|int', 'scale|!|12345', 'precision|!|0', 'comment|!|comments  for atr1']
        if len(temp[len(temp) - 1]) == 0:
            del temp[len(temp) - 1]

        flag_cortege = True
        dict_cortege_row = {}
        for temp2 in temp:
            cortege = temp2.split("|!|")  # row_from_table_list|!|2

            try:
                key = cortege[0].strip()  # row_from_table_list
                value = cortege[1].strip()  # 2
            except IndexError:
                print("not find value for cortege = " + temp2 + ". Row = " + str(temp))
                file_error.write("not find value for cortege = " + temp2 + ". Row = " + str(temp) + "\n")
                continue
            dict_cortege_row[key] = value
        list_dict_atr.append(dict_cortege_row)

    # print(list_dict_atr[0])

    ### add number tabels for write in excel
    last_table_name = list_dict_atr[0]["table_name"]
    # print(last_table_name)
    for dict_atr in list_dict_atr:
        if dict_atr["table_name"] == last_table_name:
            dict_atr["number_table_for_write"] = number_table_for_write
        else:
            number_table_for_write = number_table_for_write + 1
            last_table_name = dict_atr["table_name"]
            dict_atr["number_table_for_write"] = number_table_for_write
    ###
    print(list_dict_atr[0])

    try:
        sheet = book.active
    except Exception as err:
        print("Error! failed to open sheet: " + str(err))
        file_error.write("Error! failed to open sheet: " + str(err) + "\n")
        return

    index_write = ""
    for dict_atr in list_dict_atr:
        for default_columns in dict_default_columns_excel:
            # print(default_columns) #  comment
            # print(dict_default_columns_excel[default_columns]) #  E

            if default_columns in dict_atr:
                index_write = dict_default_columns_excel[default_columns] + str(number_row_for_write)
                try:
                    sheet[index_write] = dict_atr[default_columns]
                except Exception as error:
                    print("Failed write value for atr: " + default_columns + ". Dict = " + str(
                        dict_atr) + ". " + "Error: " + str(
                                    error))
                    file_error.write("Failed write value for atr: " + default_columns + ". Dict = " + str(
                        dict_atr) + ". " + "Error: " + str(
                                    error) + "\n")
        number_row_for_write = number_row_for_write + 1


    name_excel = path_save_excel + "\\" + date_for_name_file + "  SS_" + str(number_SS) + " struct.xlsx"
    try:
        book.save(name_excel)
    except Exception as err:
        print("Error! failed to save file: " + str(err))
        file_error.write("Error! failed to save file: " + str(err) + "\n")
