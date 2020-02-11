# -*- coding: utf-8 -*-

import time
import openpyxl

path_resource_spec = r"C:\Users\RostPK\Desktop\spec\Центр - АСРЗ Спецификация.xlsx"
path_resource_spec = path_resource_spec.decode("UTF-8")

path_template_excel = r"C:\Users\RostPK\Desktop\spec\FAT_template.xlsx"
path_save_excel = r"C:\Users\RostPK\Desktop\spec\result"
path_log = r"C:\Users\RostPK\Desktop\spec\log"


def getCountRow(sheet_spec, file_error):
    count = 0
    list_val = []
    value = ""
    try:
        for i in range(2, 10000):
            value = sheet_spec['B' + str(i)].value
            list_val.append(value)
            if i > 5:
                if (list_val[i - 2] is None) and (list_val[i - 3] is None) and (list_val[i - 4] is None):
                    count = i - 2
                    break
    except Exception as err:
        print("Error in getCountRow:" + str(err))
        file_error.write("Error in getCountRow:" + str(err) + "\n")
    return count


def getListValueSheetOne(list, file_error):
    list_write = []
    dict_row = {}

    try:
        for row in list:  # {'TABLE_ID': u'000078_MTK_TTMAIN_TROUBLE_STAGELIST', 'SRC_COLUMN_NAME': u'TSTAGE_IMAGE', 'IS_PRIMARY_KEY': u'N', 'NOT_NULL': u'N'}
            dict_row = {}
            table_name = row["TABLE_ID"]
            column_name = row["SRC_COLUMN_NAME"]
            not_null = row["NOT_NULL"]  # .encode("UTF-8")

            # print(not_null)

            if not_null == "y" or not_null == "Y":
                dict_row["table_name"] = table_name
                dict_row["column_name"] = column_name
                dict_row[
                    'code'] = "SELECT '" + table_name + "' as table_name, count(*) as count_null FROM edw_stg_ods.t_" + table_name + " WHERE " + column_name + " IS NULL union all"
                # print(dict_row['code'])
                list_write.append(dict_row)
            # break
        last_code_str = list_write[len(list_write) - 1]["code"]
        last_code_str = last_code_str.replace(" union all", ";")
        list_write[len(list_write) - 1]["code"] = last_code_str
    except Exception as err:
        print("Error in getListValueSheetOne:" + str(err))
        file_error.write("Error in getListValueSheetOne:" + str(err) + "\n")

    # print(list_write[len(list_write) - 1]["code"])
    return list_write


def getListValueSheetTwo(list, file_error):
    list_write = []
    dict_row = {}

    last_table_name = ""
    index = 0
    try:  # {'TABLE_ID': u'000078_MTK_TTMAIN_TROUBLE_STAGELIST', 'SRC_COLUMN_NAME': u'TSTAGE_IMAGE', 'IS_PRIMARY_KEY': u'N', 'NOT_NULL': u'N'}
        for row in list:
            dict_row = {}
            table_name = row["TABLE_ID"]
            column_name = row["SRC_COLUMN_NAME"]
            PK = row["IS_PRIMARY_KEY"]  # .encode("UTF-8")

            # print(not_null)

            if PK == "y" or PK == "Y":
                if last_table_name == table_name:
                    try:
                        list_write[index - 1]["pk"] = list_write[index - 1]["pk"] + column_name + "||"
                    except:
                        1 + 1
                else:
                    last_table_name = table_name

                    dict_row["table_name"] = table_name
                    dict_row["column_name"] = column_name
                    dict_row["pk"] = column_name + "||"

                    list_write.append(dict_row)
                    index = index + 1

            # break

        for row in list_write:
            temp = row["pk"].split("||")
            # select count(*) from edw_ods.t_1 group by key1, key2 having count(*) > 1
            new_pk = ""
            for el in range(0, len(temp) - 1):
                if el < len(temp) - 2:
                    new_pk = new_pk + temp[el] + ", "
                else:
                    new_pk = new_pk + temp[el]
            row["pk"] = new_pk
            row["code"] = "select '" + row["table_name"] + "' as table_name, '" + row[
                "pk"] + "' as p_key, count(*) from edw_stg_ods.t_" + row["table_name"] + " group by " + row[
                              "pk"] + " having count(*) > 1 union all"
            # print(row["code"])

        last_code_str = list_write[len(list_write) - 1]["code"]
        last_code_str = last_code_str.replace(" union all", ";")
        list_write[len(list_write) - 1]["code"] = last_code_str

    except Exception as err:
        print("Error in getListValueSheetTwo:" + str(err))
        file_error.write("Error in getListValueSheetTwo:" + str(err) + "\n")

    # print(list_write[len(list_write) - 1]["code"])
    return list_write


def getListValueSheetThree(list, file_error):
    list_write = []
    dict_row = {}

    last_table_name = ""
    index = 0
    try:  # {'FILTER': None, 'SRC_TABLE_NAME': u'MTK_GIS_ADDRESS2_VW', 'DWH_TABLE_NAME': u'T_000078_MTK_GIS_ADDRESS2_VW', 'DWH_PARTITION_TYPE': None, 'DWH_PARTITION_KEY': None}
        for row in list:
            dict_row = {}
            src_table_name = row["SRC_TABLE_NAME"]
            dwh_table_name = row["DWH_TABLE_NAME"]
            filter = row["FILTER"]  # .encode("UTF-8")
            type = row["DWH_PARTITION_TYPE"]
            part_key = row["DWH_PARTITION_KEY"]
            dict_row["SRC_TABLE_NAME"] = src_table_name

            if filter is None:
                dict_row[
                    "code_SS"] = "select '" + src_table_name + "' as table_name, count(*) as count from " + src_table_name + " union all"
                dict_row[
                    "code_GP"] = "select '" + dwh_table_name + "' as table_name, count(*) as count from edw_stg_ods." + dwh_table_name + " union all"
                dict_row["code_SS_date"] = ""
            else:
                if part_key is not None:
                    dict_row[
                        "code_SS_date"] = "select '" + src_table_name + "' as table_name, '" + part_key + "' as date_atr from " + src_table_name + ";"
                else:
                    dict_row["code_SS_date"] = ""
                if type == "NUMBER" or type == "number":
                    dict_row[
                        "code_SS"] = "select '" + src_table_name + "' as table_name, count(*) as count from " + src_table_name + " where " + filter + " union all"
                    dict_row[
                        "code_GP"] = "select '" + dwh_table_name + "' as table_name, count(*) as count from edw_stg_ods." + dwh_table_name + " union all"

                    dict_row["code_SS"] = dict_row["code_SS"].replace("$fromdt", "to_date('from_date', 'format')")
                    dict_row["code_SS"] = dict_row["code_SS"].replace("$actualdt", "to_date('act_date', 'format')")
                else:
                    dict_row[
                        "code_SS"] = "select '" + src_table_name + "' as table_name, count(*) as count from " + src_table_name + " where " + filter + " union all"
                    dict_row[
                        "code_GP"] = "select '" + dwh_table_name + "' as table_name, count(*) as count from edw_stg_ods." + dwh_table_name + " union all"

                    dict_row["code_SS"] = dict_row["code_SS"].replace("$fromdt", "'from_date', 'format'")
                    dict_row["code_SS"] = dict_row["code_SS"].replace("$actualdt", "'act_date', 'format'")
            # print(dict_row)
            list_write.append(dict_row)

            # break

        last_code_str = list_write[len(list_write) - 1]["code_SS"]
        last_code_str = last_code_str.replace(" union all", ";")
        list_write[len(list_write) - 1]["code_SS"] = last_code_str

        last_code_str = list_write[len(list_write) - 1]["code_GP"]
        last_code_str = last_code_str.replace(" union all", ";")
        list_write[len(list_write) - 1]["code_GP"] = last_code_str
    except Exception as err:
        print("Error in getListValueSheetThree:" + str(err))
        file_error.write("Error in getListValueSheetThree:" + str(err) + "\n")

    # print(list_write[len(list_write) - 1]["code"])
    return list_write


def writeDataOnListExcelOne(file_error, sheet_template, list_value_for_sheet_one):
    # {'code': u"SELECT...", 'table_name': u'000078_MTK_GIS_ADDRESS2_VW', 'column_name': u'ID'}

    row_numb = 3
    index_write = ""
    for row in list_value_for_sheet_one:
        try:
            index_write = "A" + str(row_numb)
            sheet_template[index_write] = row["table_name"]

            index_write = "B" + str(row_numb)
            sheet_template[index_write] = row["column_name"]

            index_write = "J" + str(row_numb)
            sheet_template[index_write] = row["code"]
        except Exception as err:
            print("Error in writeDataOnListExcel: row = " + str(
                row) + ". index_write = " + index_write + ". Error = " + str(err))
            file_error.write("Error in writeDataOnListExcel: row = " + str(
                row) + ". index_write = " + index_write + ". Error = " + str(err) + "\n")
        row_numb = row_numb + 1

    return 0


def writeDataOnListExcelTwo(file_error, sheet_template, list_value_for_sheet_one):
    # {'code': u"SELECT...", 'table_name': u'000078_MTK_GIS_ADDRESS2_VW', 'column_name': u'ID'}

    row_numb = 3
    index_write = ""
    for row in list_value_for_sheet_one:
        try:
            index_write = "A" + str(row_numb)
            sheet_template[index_write] = row["table_name"]

            index_write = "E" + str(row_numb)
            sheet_template[index_write] = row["column_name"]

            index_write = "H" + str(row_numb)
            sheet_template[index_write] = row["code"]
        except Exception as err:
            print("Error in writeDataOnListExcelTwo: row = " + str(
                row) + ". index_write = " + index_write + ". Error = " + str(err))
            file_error.write("Error in writeDataOnListExcelTwo: row = " + str(
                row) + ". index_write = " + index_write + ". Error = " + str(err) + "\n")
        row_numb = row_numb + 1

    return 0


def writeDataOnListExcelThree(file_error, sheet_template, list_value_for_sheet_one):
    # {'code_SS': u"select ", 'code_GP': u"select ", 'code_SS_date':'select', 'SRC_TABLE_NAME':'mtk_auth_auth_user'}.

    row_numb = 3
    index_write = ""
    for row in list_value_for_sheet_one:
        try:
            index_write = "A" + str(row_numb)
            sheet_template[index_write] = row["SRC_TABLE_NAME"]

            index_write = "G" + str(row_numb)
            sheet_template[index_write] = row["code_SS"]

            index_write = "H" + str(row_numb)
            sheet_template[index_write] = row["code_GP"]

            index_write = "I" + str(row_numb)
            sheet_template[index_write] = row["code_SS_date"]
        except Exception as err:
            print("Error in writeDataOnListExcelThree: row = " + str(
                row) + ". index_write = " + index_write + ". Error = " + str(err))
            file_error.write("Error in writeDataOnListExcelThree: row = " + str(
                row) + ". index_write = " + index_write + ". Error = " + str(err) + "\n")
        row_numb = row_numb + 1

    return 0


def setActiveSheet(book, index_sheet, file_error, book_name):
    try:
        book.active = index_sheet
        sheet_spec = book.active
    except Exception as err:
        print("Error! failed to open sheet " + book_name + ": " + str(err))
        file_error.write("Error! failed to open sheet " + book_name + ": " + str(err) + "\n")
        return
    return sheet_spec


def getDraftFAT():
    # date_now = time.strftime("%d.%m  %H_%M_%S", time.localtime())
    date_now = time.strftime("%d.%m  ", time.localtime())
    print("prefix name for file: " + date_now)

    file_name_error = path_log + "\\" + date_now + "  FAT error.txt"
    file_error = open(file_name_error, "w")

    try:
        book_spec = openpyxl.load_workbook(path_resource_spec)
    except Exception as error:
        print("Could not open file specificaion '" + path_resource_spec + "'", error)
        file_error.write("Could not open file template.xlsx:" + str(error) + "\n")
        return

    try:
        book_template = openpyxl.load_workbook(path_template_excel)
    except Exception as error:
        print("Could not open file template.xlsx", error)
        file_error.write("Could not open file template.xlsx:" + str(error) + "\n")
        return

    ################
    sheet_spec = setActiveSheet(book_spec, 1, file_error, "book_spec")
    ################

    ind_row = 1
    COL_SRC_COLUMN_NAME = "B"
    COL_TABLE_ID = "C"
    COL_NOT_NULL = "I"
    COL_IS_PRIMARY_KEY = "J"

    count_row = getCountRow(sheet_spec, file_error)  # max count row

    list_dict = []
    dict_spec_row = {}
    value = ""
    for ind_row in range(2, count_row):
        dict_spec_row = {}
        value = sheet_spec[COL_SRC_COLUMN_NAME + str(ind_row)].value
        dict_spec_row["SRC_COLUMN_NAME"] = value

        value = sheet_spec[COL_TABLE_ID + str(ind_row)].value
        dict_spec_row["TABLE_ID"] = value

        value = sheet_spec[COL_NOT_NULL + str(ind_row)].value
        dict_spec_row["NOT_NULL"] = value

        value = sheet_spec[COL_IS_PRIMARY_KEY + str(ind_row)].value
        dict_spec_row["IS_PRIMARY_KEY"] = value

        list_dict.append(dict_spec_row)

    # print(list_dict) #  {'TABLE_ID': u'000078_MTK_TTMAIN_TROUBLE_STAGELIST', 'SRC_COLUMN_NAME': u'TSTAGE_IMAGE', 'IS_PRIMARY_KEY': u'N', 'NOT_NULL': u'N'}

    list_value_for_sheet_one = getListValueSheetOne(list_dict, file_error)
    sheet_template = setActiveSheet(book_template, 2, file_error, "book_template")
    writeDataOnListExcelOne(file_error, sheet_template, list_value_for_sheet_one)

    list_value_for_sheet_two = getListValueSheetTwo(list_dict, file_error)
    sheet_template = setActiveSheet(book_template, 3, file_error, "book_template")
    writeDataOnListExcelTwo(file_error, sheet_template, list_value_for_sheet_two)

    ################################

    sheet_spec = setActiveSheet(book_spec, 0, file_error, "book_spec")
    count_row = getCountRow(sheet_spec, file_error)  # max count row

    SRC_TABLE_NAME = "D"
    FILTER = "F"
    DWH_TABLE_NAME = "I"
    COL_IS_PRIMARY_KEY = "J"
    DWH_PARTITION_TYPE = "M"
    DWH_PARTITION_KEY = "L"
    list_dict2 = []
    dict_row = {}

    for ind_row in range(2, count_row):
        dict_row = {}
        value = sheet_spec[SRC_TABLE_NAME + str(ind_row)].value
        dict_row["SRC_TABLE_NAME"] = value

        value = sheet_spec[FILTER + str(ind_row)].value
        dict_row["FILTER"] = value

        value = sheet_spec[DWH_TABLE_NAME + str(ind_row)].value
        dict_row["DWH_TABLE_NAME"] = value

        value = sheet_spec[DWH_PARTITION_TYPE + str(ind_row)].value
        dict_row["DWH_PARTITION_TYPE"] = value

        value = sheet_spec[DWH_PARTITION_KEY + str(ind_row)].value
        dict_row["DWH_PARTITION_KEY"] = value

        list_dict2.append(dict_row)

    list_value_for_sheet_three = getListValueSheetThree(list_dict2, file_error)
    sheet_template = setActiveSheet(book_template, 4, file_error, "book_template")
    writeDataOnListExcelThree(file_error, sheet_template, list_value_for_sheet_three)

    # print(list_value_for_sheet_three[0])

    ################

    name_excel = path_save_excel + "\\" + date_now + " FAT.xlsx"
    try:
        book_template.save(name_excel)
    except Exception as err:
        print("Error! failed to save file: " + str(err))
        file_error.write("Error! failed to save file: " + str(err) + "\n")
        return
    ################

    file_error.close()


getDraftFAT()
