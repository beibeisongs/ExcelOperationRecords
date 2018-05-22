# encoding=utf-8
# Date:2018-05-02


import xlrd

import os
import xlwt
from  xlutils.copy import copy


def get_tot_times(excel_path):

    f1 = xlrd.open_workbook(excel_path, 'r')
    _sheet = f1.sheet_by_index(0)
    tot_times = _sheet.cell(1, 5).value

    return tot_times


def openExcelFileToWrite(excel_filepath, lng, lat, date, gender, get_user_id, time):

    if os.path.exists(excel_filepath) == False:

        book = xlwt.Workbook()  # 创建一个Excel表对象
        sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)
        title = ['lngtitude', 'latitude', 'date', 'time', 'gender', 'Tot_Times', 'user_id']

        tot_times = 1

        for i in range(len(title)):
            sheet.write(0, i, title[i])

        sheet.write(1, 0, lng)
        sheet.write(1, 1, lat)
        sheet.write(1, 2, date)
        sheet.write(1, 3, time)
        sheet.write(1, 4, gender)
        sheet.write(1, 5, tot_times)
        sheet.write(1, 6, get_user_id)

        book.save(excel_filepath)
    else:

        old_excel = xlrd.open_workbook(excel_filepath, formatting_info=True)

        # 将操作文件对象拷贝，变成可写的excel对象
        new_excel = copy(old_excel)

        # 获得第一个sheer的对象
        ws = new_excel.get_sheet(0)

        tot_times = int(get_tot_times(excel_filepath)) + 1

        ws.write(tot_times, 0, lng)
        ws.write(tot_times, 1, lat)
        ws.write(tot_times, 2, date)
        ws.write(tot_times, 3, time)
        ws.write(tot_times, 4, gender)
        ws.write(1, 5, tot_times)
        ws.write(tot_times, 6, get_user_id)

        if tot_times > 1:
            print("!!!")
            print("!!!")
            print("!!!")

        new_excel.save(excel_filepath)


def goThrough_Excel(sheet1, nrows, ncols, document_path):

    for line in range(1, nrows):    # <Description>: go through the content of the Excel

        print("row : ", line)

        get_user_id = str(int(sheet1.cell(line, 3).value))
        print("The user id is : ", get_user_id)
        filepath = document_path + "\\" + get_user_id

        lng = sheet1.cell(line, 0).value
        lat = sheet1.cell(line, 1).value
        date = sheet1.cell(line, 2).value

        gender = sheet1.cell(line, 4).value
        time = sheet1.cell(line, 5).value

        if os.path.exists(filepath) == False:
            os.makedirs(filepath)
            print("The filepath is : ", filepath)
            excel_filepath = filepath + "\\" + get_user_id + ".xls"
            openExcelFileToWrite(excel_filepath, lng, lat, date, gender, get_user_id, time)
        else:
            print("The filepath is : ", filepath)
            excel_filepath = filepath + "\\" + get_user_id + ".xls"
            openExcelFileToWrite(excel_filepath, lng, lat, date, gender, get_user_id, time)


def openExcel(excel_path):

    xlfile = xlrd.open_workbook(excel_path, 'r')
    sheet1 = xlfile.sheet_by_index(0)
    nrows = sheet1.nrows
    ncols = sheet1.ncols

    return xlfile, sheet1, nrows, ncols

def createDocument(document_path):
    if os.path.exists(document_path) == False:
        os.makedirs(document_path)


if __name__ == "__main__":

    document_path = "C:\\TheLabProject2th"
    createDocument(document_path)

    excel_path = "big1.xlsx"
    xlfile, sheet1, nrows, ncols = openExcel(excel_path)

    goThrough_Excel(sheet1, nrows, ncols, document_path)