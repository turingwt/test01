import xlrd
import xlwt
import xlutils
from xlutils.copy import copy


def form_init() -> tuple:  # 表格读取及格式转换处理
    """
    复杂方法：按照原表格形式排列，后续需先找出各列，再横向排列组合，不再使用，用下面form_init1()替换
    读入表格，获取全部数据并作为二维字符串列表赋值给变量；
    "统计日期"列格式处理：浮点型数值转化为字符串形式  XXXX-XX-XX
    :return: 返回元组：二维列表，行数，列数
    """
    table = xlrd.open_workbook('Bus_May.xlsx')  # 获取Excel表格文件
    da = table.sheet_by_name('2017')  # 获取‘2017’sheet页
    rows = da.nrows  # 总行数
    columns = da.ncols  # 总列数
    all_data = [[da._cell_values[i][j] for j in range(columns)] for i in range(rows)]  # 读取全部数据
    for j in range(1, rows):  # 处理日期格式，将浮点值转化为XXXX-XX-XX格式
        all_data[j][0] = '-'.join([str(xlrd.xldate.xldate_as_tuple(all_data[j][0], 0)[i]) for i in range(3)])

    return all_data, rows, columns


def data_init() -> tuple:
    """
    复杂方法，调用form_init(),不使用
    找到各列并单独赋值，字符串格式
    :return: 返回各列
    """
    data, data_rows, data_columns = form_init()  # da1=全部表格数据

    time_index = data[0].index('统计日期')  # 找到统计日期所在列
    init_time = [data[i][time_index] for i in range(data_rows)]  # time = “统计日期”列
    vin_index = data[0].index('VIN码')  # 找到VIN码所在列
    init_vin = [data[i][vin_index] for i in range(data_rows)]  # vin = “VIN码”列
    cartype_index = data[0].index('车型代码')  # 找到车型代码所在列
    init_cartype = [data[i][cartype_index] for i in range(data_rows)]  # cartype = “车型代码”列
    city_index = data[0].index('运营城市')  # 找到运营城市所在列
    init_city = [data[i][city_index] for i in range(data_rows)]  # city = “运营城市”列
    gps_day_index = data[0].index('GPS日行驶里程(KM)')  # 找到Gps日行驶里程所在列
    init_gps_day = [data[i][gps_day_index] for i in range(data_rows)]  # Gps_day = “Gps日行驶里程”列
    mileage_index = data[0].index('仪表里程(KM)')  # 找到仪表里程所在列
    init_mileage = [data[i][mileage_index] for i in range(data_rows)]  # mileage = “仪表里程”列

    init_province = [init_city[i].split('--')[0] for i in range(data_rows)]
    result_city = [init_city[i].split('--')[1] for i in range(1, data_rows)]

    result_data = [init_time, init_vin, init_cartype, init_province, result_city, init_gps_day, init_mileage]
    return result_data


# table = xlrd.open_workbook('Bus.xls', formatting_info=True)
# wbook = copy(table)
# w_sheet = wbook.get_sheet(0)
# w_sheet.write(0, 10, 'A573aj')
# wbook.save('A.xls')

print(type('asd-asdf-assf'.rpartition('-')))


