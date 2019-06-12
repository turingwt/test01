import xlrd
import datetime
start = datetime.datetime.now()


# def form_init():  # 表格读取及格式转换处理
table = xlrd.open_workbook('LDYECS569G0005607_20190609140114.xls')  # 获取Excel表格文件
# da = table.sheet_by_name('Sheet1')  # 获取‘Sheet1’sheet页
# rows = da.nrows  # 总行数
# columns = da.ncols  # 总列数
# all_data = [[da._cell_values[i][j] for j in range(columns)] for i in range(rows)]  # 读取全部数据
    # for j in range(1, rows):  # 处理日期格式，将浮点值转化为XXXX-XX-XX格式
    #     all_data[j][0] = '-'.join([str(xlrd.xldate.xldate_as_tuple(all_data[j][0], 0)[i]) for i in range(3)])
    # return all_data, rows, columns


# data, data_rows, data_columns = form_init()
print('{} 秒'.format((datetime.datetime.now() - start).total_seconds()))




