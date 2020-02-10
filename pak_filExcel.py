import openpyxl
import time

path_template_excel = r"C:\Users\RostPK\Desktop\spec\template.xlsx"
path_save_excel = r"C:\Users\RostPK\Desktop\spec\result"
number_table_for_write = 1  # Start serial number for recording in excel file
number_row_for_write = 2  # Start position for write row

# defaine columns for write list2. Variable = list_dict_atr[{}]
dict_default_columns_excel = {
    "COLUMN_ORDER": "A",
    "SRC_COLUMN_NAME": "B",
    "TABLE_ID": "C",
    "COLUMN_NAME": "D",
    "COLUMN_COMMENT": "E",
    "DATA_TYPE_ID": "F",
    "SCALE": "G",
    "PRECISION": "H",
    "NOT_NULL": "I",
    "IS_PRIMARY_KEY": "J",
    "IS_DISTRIB_COL": "N"
}


# excel_index
#

def getTypePrecisonScale(data_type, p_precision, p_scale, file_error):
    result = {}
    type = data_type
    precision = p_precision
    scale = p_scale
    flag_find_type = False
    if len(data_type) > 2:
        if data_type == "NUMBER":
            flag_find_type = True
            try:
                if len(str(precision)) > 0:
                    if int(precision) > 38:
                        precision = 38
                    if int(precision) < 0:
                        precision = 38
                else:
                    precision = 38
            except Exception:
                print("Error in data_type NUMBER|precision for type:" + type + "; precision: " + str(
                    precision) + "; scale: " + str(
                    scale))
                file_error.write("Error in data_type NUMBER|precision for type:" + type + "; precision: " + str(
                    precision) + "; scale: " + str(
                    scale) + "\n")
                precision = 38
            try:
                if len(str(scale)) > 0:
                    if int(scale) > 5:
                        scale = 5
                    if int(scale) < 0:
                        scale = 0
                else:
                    scale = 0
            except Exception:
                print("Error in data_type NUMBER|scale for type:" + type + "; precision: " + str(
                    precision) + "; scale: " + str(
                    scale))
                file_error.write("Error in data_type NUMBER|scale for type:" + type + "; precision: " + str(
                    precision) + "; scale: " + str(
                    scale) + "\n")
                scale = 0
        if data_type == "TIMESTAMP":
            flag_find_type = True
            precision = ""
            scale = ""
        if data_type == "VARCHAR":
            flag_find_type = True
            try:
                if len(str(precision)) > 0:
                    if int(precision) > 4000:
                        precision = 4000
                    if int(precision) < 0:
                        precision = 100
                else:
                    precision = 4000
            except Exception:
                print("Error in data_type VARCHAR for type:" + type + "; precision: " + str(
                    precision) + "; scale: " + str(
                    scale))
                file_error.write("Error in data_type VARCHAR for type:" + type + "; precision: " + str(
                    precision) + "; scale: " + str(
                    scale) + "\n")
                precision = 4000

            scale = ""
    if flag_find_type:
        result = {
            "type": type,
            "precision": precision,
            "scale": scale
        }
    else:
        result = {
            "type": type,
            "precision": precision,
            "scale": scale
        }
        print("Don't find type in type_list for type:" + type + "; precision: " + str(precision) + "; scale: " + str(
            scale))
        file_error.write(
            "Don't find type in type_list for type:" + type + "; precision: " + str(precision) + "; scale: " + str(
                scale) + "\n")
    return result


def constructorFillExcel(file_error, file_name_struct, date_for_name_file, number_SS, file_warning):
    # print(file_name_error)
    # print(file_name_struct)
    # print(date_for_name_file)

    # book = openpyxl.Workbook()
    global number_table_for_write
    global path_save_excel
    global path_template_excel
    global number_row_for_write
    global dict_default_columns_excel

    prefix_table = ""
    for i in range(6 - len(str(number_SS))):
        prefix_table = prefix_table + "0"
    prefix_table = prefix_table + str(number_SS) + "_"
    # print(prefix_table)

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

    ### Treatment dict atr before write
    last_table_name = list_dict_atr[0]["table_name"]
    # print(last_table_name)
    for dict_atr in list_dict_atr:
        type_precision_scale = {}

        if dict_atr["table_name"] == last_table_name:
            dict_atr["number_table_for_write"] = number_table_for_write
        else:
            number_table_for_write = number_table_for_write + 1
            last_table_name = dict_atr["table_name"]
            dict_atr["number_table_for_write"] = number_table_for_write

        if "COLUMN_ORDER" in dict_atr:
            dict_atr['COLUMN_ORDER'] = (dict_atr['COLUMN_ORDER']).upper()
        else:
            file_warning.write("don't find COLUMN_ORDER value in dict_row = " + str(dict_atr) + "\n")

        if "SRC_COLUMN_NAME" in dict_atr:
            dict_atr['SRC_COLUMN_NAME'] = (dict_atr['SRC_COLUMN_NAME']).upper()

        if "table_name" in dict_atr:
            dict_atr["TABLE_ID"] = (prefix_table + dict_atr["table_name"]).upper()

        if "COLUMN_NAME" in dict_atr:
            dict_atr['COLUMN_NAME'] = (dict_atr['COLUMN_NAME']).upper()

        if "COLUMN_COMMENT" in dict_atr:
            dict_atr['COLUMN_COMMENT'] = (dict_atr['COLUMN_COMMENT']).upper()

        if "DATA_TYPE_ID" in dict_atr:
            if "PRECISION" in dict_atr:
                if "SCALE" in dict_atr:
                    type_precision_scale = getTypePrecisonScale(dict_atr["DATA_TYPE_ID"], dict_atr["PRECISION"],
                                                                dict_atr["SCALE"], file_error)
                else:
                    type_precision_scale = getTypePrecisonScale(dict_atr["DATA_TYPE_ID"], dict_atr["PRECISION"], "",
                                                                file_error)
            else:
                type_precision_scale = getTypePrecisonScale(dict_atr["DATA_TYPE_ID"], "", "", file_error)
                file_warning.write("don't find PRECISION value in dict_row = " + str(dict_atr) + "\n")
            dict_atr["DATA_TYPE_ID"] = type_precision_scale["type"]
            dict_atr["PRECISION"] = type_precision_scale["precision"]
            dict_atr["SCALE"] = type_precision_scale["scale"]
        else:
            file_warning.write("don't find DATA_TYPE_ID value in dict_row = " + str(dict_atr) + "\n")

        if "NOT_NULL" in dict_atr:
            dict_atr['NOT_NULL'] = (dict_atr['NOT_NULL']).upper()

        if "IS_PRIMARY_KEY" in dict_atr:
            dict_atr['IS_PRIMARY_KEY'] = (dict_atr['IS_PRIMARY_KEY']).upper()

        if "IS_DISTRIB_COL" in dict_atr:
            dict_atr['IS_DISTRIB_COL'] = (dict_atr['IS_DISTRIB_COL']).upper()

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
