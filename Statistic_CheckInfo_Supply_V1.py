# encoding=utf-8
# Date: 2018-5-17


import xlrd
import xlwt


def openAccountExcel(account_path):

    account_xlfile = xlrd.open_workbook(account_path, 'r')
    account_sheet1 = account_xlfile.sheet_by_index(0)
    account_nrows = account_sheet1.nrows
    account_ncols = account_sheet1.ncols

    return account_xlfile, account_sheet1, account_nrows, account_ncols


def goThrough_Excel(sheet1, nrows, ncols, excel_path_toWrite, written_mark):

    book = xlwt.Workbook()  # 创建一个Excel表对象
    sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)

    account_path_pre = "C:\\TheLabProject2th"

    line = 1

    write_supply_cursor = 0

    while line <= nrows - 1:

        print("row : ", line)

        get_user_id = str(int(sheet1.cell(line, 6).value))
        get_tot_times = int(sheet1.cell(line, 5).value)

        for read_small_account_cursor in range(0, get_tot_times):

            account_lng = sheet1.cell(line + read_small_account_cursor, 0).value
            account_lat = sheet1.cell(line + read_small_account_cursor, 1).value
            account_date = sheet1.cell(line + read_small_account_cursor, 2).value
            account_time = sheet1.cell(line + read_small_account_cursor, 3).value
            account_gender = sheet1.cell(line + read_small_account_cursor, 4).value
            account_tot_times = sheet1.cell(line + read_small_account_cursor, 5).value
            account_userid = sheet1.cell(line + read_small_account_cursor, 6).value
            account_region = sheet1.cell(line + read_small_account_cursor, 7).value

            sheet.write(write_supply_cursor, 0, account_lng)
            sheet.write(write_supply_cursor, 1, account_lat)
            sheet.write(write_supply_cursor, 2, account_date)
            sheet.write(write_supply_cursor, 3, account_time)
            sheet.write(write_supply_cursor, 4, account_gender)
            sheet.write(write_supply_cursor, 5, account_tot_times)
            sheet.write(write_supply_cursor, 6, account_userid)
            sheet.write(write_supply_cursor, 7, account_region)

            write_supply_cursor += 1

            if write_supply_cursor == 60000:

                book.save(excel_path_toWrite)
                written_mark += 1
                excel_path_toWrite = "supply" + str(written_mark) + ".xls"

                book = xlwt.Workbook()  # 创建一个Excel表对象
                sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)

                write_supply_cursor = 0

        try:

            account_path = account_path_pre + "\\" + get_user_id + "\\" + get_user_id + ".xls"
            account_xlfile, account_sheet1, account_nrows, account_ncols = openAccountExcel(account_path)

            for read_account_cursor  in range(1, account_nrows):

                account_lng = account_sheet1.cell(read_account_cursor, 0).value
                account_lat = account_sheet1.cell(read_account_cursor, 1).value
                account_date = account_sheet1.cell(read_account_cursor, 2).value
                account_time = account_sheet1.cell(read_account_cursor, 3).value
                account_gender = account_sheet1.cell(read_account_cursor, 4).value
                account_tot_times = account_sheet1.cell(read_account_cursor, 5).value
                account_userid = account_sheet1.cell(read_account_cursor, 6).value

                sheet.write(write_supply_cursor, 0, account_lng)
                sheet.write(write_supply_cursor, 1, account_lat)
                sheet.write(write_supply_cursor, 2, account_date)
                sheet.write(write_supply_cursor, 3, account_time)
                sheet.write(write_supply_cursor, 4, account_gender)
                sheet.write(write_supply_cursor, 5, account_tot_times)
                sheet.write(write_supply_cursor, 6, account_userid)

                write_supply_cursor += 1

                if write_supply_cursor == 60000:

                    book.save(excel_path_toWrite)
                    written_mark += 1
                    excel_path_toWrite = "supply" + str(written_mark) + ".xls"

                    book = xlwt.Workbook()  # 创建一个Excel表对象
                    sheet = book.add_sheet('Sheet1', cell_overwrite_ok=True)

                    write_supply_cursor = 0

        except:
            print("skip ! ")

        line += get_tot_times

        book.save(excel_path_toWrite)


def openExcel(excel_path):

    xlfile = xlrd.open_workbook(excel_path, 'r')
    sheet1 = xlfile.sheet_by_index(0)
    nrows = sheet1.nrows
    ncols = sheet1.ncols

    return xlfile, sheet1, nrows, ncols


if __name__ == "__main__":

    excel_path = "small.xlsx"
    xlfile, sheet1, nrows, ncols = openExcel(excel_path)

    written_mark = 0
    excel_path_toWrite = "supply" + str(written_mark) + ".xls"
    goThrough_Excel(sheet1, nrows, ncols, excel_path_toWrite, written_mark)

