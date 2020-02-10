def my_print(variable):
    if isinstance(variable, (float, int, str)):
        print(variable)
        return
    i_ind = 0
    j_ind = 0

    print("Array(count = " + str(len(variable)) + ")(")
    for i in variable:
        j_ind = 0
        print ("[" + str(i_ind) + "] => ")
        print ("    (")
        if isinstance(i, (float, int, str)):
            print ("      " + str(i))
            print ("    )")
            i_ind = i_ind + 1
            continue
        try:
            for j in i:
                print("     [" + str(j_ind) + "] => '" + j + "'")
                j_ind = j_ind + 1
        except Exception:
            b = 0
        print ("    )")
        i_ind = i_ind + 1
    print(");")


def getStructureTable(table_name_in_file, cursor, file_error, id_table_from_file):
    temp = table_name_in_file.split(".")
    try:
        table_schema = temp[0]
        table_name = temp[1]
    except IndexError:
        print("not find schema in name = " + table_name_in_file + ", " + id_table_from_file)
        file_error.write("not find schema in name = " + table_name_in_file + ", " + id_table_from_file + "\n")
        return []
    ### deleted this block
    path_file_temp_struct = r'C:\Users\RostPK\Desktop\spec\atr_list.txt'

    temp_struct = []
    with open(path_file_temp_struct) as file:
        for line in file.readlines():
            temp_struct.append(
                id_table_from_file + "|!!|table_name|!|" + table_name + "|!!|table_schema|!|" + table_schema + "|!!|" + line.strip())
    ###

    ###
    cursor = "execute query to DB"
    ###
    # treatment result from database
    ###
    table_structure = temp_struct
    # print(table_structure) #'row_from_table_list|!|2|!!|table_name|!|ekhd.tables1first|!!|number_row|!|number_row1|!!|name_atr|!|atr1|!!|type|!|int|!!|scale|!|12345|!!|precision|!|0|!!|comment|!|comments  for atr1|!!|'
    return table_structure


import time

# date_now = time.strftime("%d.%m  %H_%M_%S", time.localtime())
date_now = time.strftime("%d.%m  ", time.localtime())
print("prefix name for file: " + date_now)

path_tables_list_file = r"C:\Users\RostPK\Desktop\spec\tables_list.txt"

path_tables_resource = r"C:\Users\RostPK\Desktop\spec\log"
path_log = r"C:\Users\RostPK\Desktop\spec\log"

number_SS = 0  # number source system default

list_tables = []
try:
    with open(path_tables_list_file) as file:
        for line in file.readlines():
            list_tables.append(line.strip())
except Exception as error:
    print("Error opening file: ", error)

# print(list_tables)
id_row_from_file = 1  # numbering row in file starts from 1
if list_tables[0].isdigit():
    number_SS = list_tables[0]
    del list_tables[0]  # deleted number SS from list_tables
    id_row_from_file = 2  # row #1 in file - SS

path_tables_resource = path_tables_resource + "\\" + date_now + "  SS_" + str(number_SS)

file_name_error = path_log + "\\" + date_now + "  SS_" + str(number_SS) + " error.txt"
file_error = open(file_name_error, "w")
file_name_warning = path_log + "\\" + date_now + "  SS_" + str(number_SS) + " warning.txt"
file_warning = open(file_name_warning, "w")
file_name_struct = path_tables_resource + " struct.txt"
file_struct = open(file_name_struct, "w")

#################################### Set connection
cursor = 1
#################################### Get the structure of each table from the list_tables
list_tables_structure = []
for table in list_tables:
    list_tables_structure.append(
        getStructureTable(table, cursor, file_error, "row_from_table_list|!|" + str(id_row_from_file)))
    id_row_from_file = id_row_from_file + 1
    # break
# my_print(list_tables_structure)
for table in list_tables_structure:
    for row in table:
        file_struct.write(row + "\n")

####################################

file_struct.close()

import pak_fillExcel

pak_fillExcel.constructorFillExcel(file_error, file_name_struct, date_now, number_SS, file_warning)

file_error.close()
