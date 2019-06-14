import xlrd
import os
import datetime
import xlsxwriter
from xlutils.copy import copy


# 获得目标文件
def getxls():
    path = os.getcwd()
    files = os.listdir(path)
    targetfile = [i for i in files if '月度' in i]

    return targetfile


# 表格预处理
def form_init(file: str) -> tuple:  # 表格读取及格式转换处理
    """
    读入表格，获取全部数据并作为二维字符串列表赋值给变量；
    "统计日期"列格式处理：浮点型数值转化为字符串形式  XXXX-XX-XX
    :param file: 目标文件列表
    :return: 返回元组：二维列表，行数，列数
    """
    table = xlrd.open_workbook(file)  # 获取Excel表格文件
    da = table.sheet_by_name('Sheet1')  # 获取‘2017’sheet页
    rows = da.nrows  # 总行数
    columns = da.ncols  # 总列数
    all_data = [[da._cell_values[j][i] for j in range(rows)] for i in range(columns)]  # 读取全部数据
    for j in range(1, rows):  # 处理日期格式，将浮点值转化为XXXX-XX-XX格式
        all_data[0][j] = '-'.join([str(xlrd.xldate.xldate_as_tuple(all_data[0][j], 0)[i]) for i in range(3)])

    all_data.append([all_data[6][i].split('--')[0] for i in range(0, rows)])  # 省份及城市分割
    all_data.append([all_data[6][i].split('--')[1] for i in range(1, rows)])
    all_data[11].insert(0, '城市')
    all_data[10][0] = '省份'

    return all_data, rows, columns


# 省份信息处理（各省车辆数、月总里程数、各去重列表）
def province_(data_province: list, rows_province: int) -> tuple:
    """
    省份数据处理
    :param data_province: data,初始化的全部数据
    :param rows_province: 全部数据条数
    :return: car_province_4 省份数据字典（省份、各省车辆数目、各省月总里程）、
             [_vin, _type, _city, _days] 去重后的数据列表[vin, 车型, 运营城市, 车辆月度上线天数]
    """
    car_province = list(set(data_province[10]))
    province_index = car_province.index('省份')
    if province_index != 0:  # 省份总揽去重
        car_province[province_index], car_province[0] = car_province[0], car_province[province_index]

    len_province = len(car_province)  # 构造字典
    car_province_1 = [0 for _ in range(len_province)]
    car_province_2 = dict(zip(car_province, car_province_1))
    car_province_1 = [0.0 for _ in range(len_province)]
    car_province_3 = dict(zip(car_province, car_province_1))

    # vin去重索引号
    vin_deduplication = [i + 1 for i in range(rows_province - 1) if data_province[2][i] != data_province[2][i + 1]]

    # 全部车辆的车型
    _vin = ['VIN']
    [_vin.append(data_province[2][i]) for i in vin_deduplication]

    # 全部车辆的车型
    _type = ['车型']
    [_type.append(data_province[4][i]) for i in vin_deduplication]

    # 全部车辆的运营城市
    _city = ['运营城市']
    [_city.append(data_province[11][i]) for i in vin_deduplication]

    # 全部车辆的月度上线天数
    _days = ['上线天数']
    for i in range(len(vin_deduplication)):
        ds = 0
        k = vin_deduplication[i]
        for j in range(35):
            if data_province[2][k] == data_province[2][vin_deduplication[i]]:
                if data_province[7][k] > 0:
                    ds += 1
            else:
                break
            k += 1
            if k == rows_province:
                break
        _days.append(ds)

    for i in vin_deduplication:  # 计算各省车辆数目
        for j, k in enumerate(car_province_2.keys()):
            if data_province[10][i] == k:
                car_province_2[k] += 1
    car_province_2['省份'] = '车辆数目'
    for i in range(1, rows_province):  # 计算各省月总里程
        for j, k in enumerate(car_province_3.keys()):
            if data_province[10][i] == k:
                car_province_3[k] += data_province[7][i]
    for i in car_province_3.keys():  # 各省月总里程之和不保留小数点
        car_province_3[i] = round(car_province_3[i])
    car_province_3['省份'] = '各省月总里程(KM)'
    car_province_4 = {j: [car_province_2[j], car_province_3[j]] for i, j in enumerate(car_province_2.keys())}

    return car_province_4, [_vin, _type, _city, _days]


# 每日信息处理（月度每日上线车辆数、每日总里程）
def day_(_data: list, _data_rows: int) -> dict:
    """
    :param _data: 全部数据
    :param _data_rows: 行数
    :return: 每日数据字典（日期：[每天上线车辆数，每天总里程]）
    """
    time = ['日期']
    [time.append(_data[0][i]) for i in range(1, 35) if _data[0][i] not in time]
    everyday_mileage = [['每天上线车辆数', '每天总里程(KM)']]
    [everyday_mileage.append([0, 0.0]) for _ in range(len(time) - 1)]
    for i in range(1, _data_rows):
        for j in range(1, len(time)):
            if _data[0][i] == time[j]:
                if _data[7][i] > 0:  # 如果车辆在线
                    everyday_mileage[j][1] += _data[7][i]  # 日行驶里程累加
                    everyday_mileage[j][0] += 1  # 日上线车辆数目累加
    for i in range(1, len(everyday_mileage)):
        everyday_mileage[i][1] = round(everyday_mileage[i][1])
    day_msg = dict(zip(time, everyday_mileage))

    return day_msg


# 月度上线天数区间统计
def days_online(_days: list):
    """
    :param _days: list[每辆车的月度上线天数]
    :return _days_online: dict{上线天数: 上线天数}
    """
    _days_online = {'上线天数': '上线天数',
                    '0': 0,
                    '1-10': 0,
                    '11-20': 0,
                    '21-30': 0,
                    '31': 0
                    }

    for i in range(1, len(_days)):
        if _days[i] == 0:
            _days_online['0'] += 1
        elif 0 < _days[i] <= 10:
            _days_online['1-10'] += 1
        elif 10 < _days[i] <= 20:
            _days_online['11-20'] += 1
        elif 20 < _days[i] <= 30:
            _days_online['21-30'] += 1
        elif _days[i] == 31:
            _days_online['31'] += 1

    return _days_online


# 车型信息处理（车型列表、车型车辆数目）
def type_cars_count(_dedupliction: list) -> dict:
    """
    车型及各车型车辆数目计算
    :param _dedupliction: 去重后列表
    :return: 车型及各车型车辆数目字典
    """
    type_cars = list(set(_dedupliction[1]))  # 车型去重，构建列表
    type_cars_0 = type_cars.index('车型')  # 查找'车型'项所在位置
    type_cars[0], type_cars[type_cars_0] = type_cars[type_cars_0], type_cars[0]  # '车型'项移到首行
    type_cars_dict = {type_cars[i]: _dedupliction[1].count(type_cars[i]) for i in range(len(type_cars))}  # 构建字典
    type_cars_dict['车型'] = '车辆数目'

    return type_cars_dict


# 累计里程信息处理（截止该月最后一天每辆车的累计里程）
def mileage_(_data: list, vin_dedupliction: list):
    """
    :param _data: 全部数据
    :param vin_dedupliction: 去重后的VIN列表
    :return: 字典{'仪表里程(KM)': '车辆数目'}
    """
    len_data = len(_data[0])
    time = ['日期']
    [time.append(_data[0][i]) for i in range(1, 35) if _data[0][i] not in time]
    max_date = str(len(time) - 1)
    date = [_data[0][i].rpartition('-')[-1] for i in range(len_data)]
    mileage = [_data[9][i] for i in range(len(date)) if date[i] == max_date]
    mileage.insert(0, '仪表里程')

    mil = {'仪表里程(KM)': '车辆数目',
           '5000以下': 0,
           '5000-9999': 0,
           '10000-14999': 0,
           '15000-19999': 0,
           '20000-24999': 0,
           '25000-29999': 0,
           '30000以上': 0}

    for i in range(1, len(mileage)):
        if 0 <= mileage[i] < 5000:
            mil['5000以下'] += 1
        elif 5000 <= mileage[i] < 10000:
            mil['5000-9999'] += 1
        elif 10000 <= mileage[i] < 15000:
            mil['10000-14999'] += 1
        elif 15000 <= mileage[i] < 20000:
            mil['15000-19999'] += 1
        elif 20000 <= mileage[i] < 25000:
            mil['20000-24999'] += 1
        elif 25000 <= mileage[i] < 30000:
            mil['25000-29999'] += 1
        elif mileage[i] >= 30000:
            mil['30000以上'] += 1
        else:
            print("累计里程有负值！请检查！ {}".format(vin_dedupliction[i]))

    return mil


# 信息写入表格及表格另存(只适用.xls格式，不适用.xlsx格式，本程序不再使用该函数)
def write_result(_filename: str,
                 _province_result: dict, _day_result: dict,
                 _dedupliction: list, _type_cars_count: dict,
                 _mileage_result: dict, _days_online: dict) -> None:  # 接上一行
    """
    :param _filename: 文件名
    :param _province_result: 省份数据字典{省份: [各省车辆数目,各省月总里程]}
    :param _day_result: 每日数据字典{日期：[每天上线车辆数，每天总里程]}
    :param _dedupliction: 去重后的数据列表[vin, 车型, 运营城市, 车辆月度上线天数]
    :param _type_cars_count: 字典{车型：各车型车辆数目}
    :param _mileage_result: 字典{仪表里程(KM): 车辆数目}
    :param _days_online:字典{月度上线天数区间：车辆数目}
    :return: 无
    """
    table = xlrd.open_workbook(_filename, formatting_info=True)
    wbook = copy(table)  # 格式化复制原表格
    w_sheet = wbook.get_sheet(0)
    base = 10  # 起始写入列

    [w_sheet.write(i, base, j) for i, j in enumerate(_province_result.keys())]  # 写入省份数据
    [w_sheet.write(i, base + 1, _province_result[j][0]) for i, j in enumerate(_province_result.keys())]
    [w_sheet.write(i, base + 2, _province_result[j][1]) for i, j in enumerate(_province_result.keys())]

    [w_sheet.write(i, base + 3, j) for i, j in enumerate(_day_result.keys())]  # 写入每日数据
    [w_sheet.write(i, base + 4, _day_result[j][0]) for i, j in enumerate(_day_result.keys())]
    [w_sheet.write(i, base + 5, _day_result[j][1]) for i, j in enumerate(_day_result.keys())]

    [w_sheet.write(j, base + 6 + i, _dedupliction[i][j]) for i in range(len(_dedupliction))  # 写入去重列表（4列）
     for j in range(len(_dedupliction[i]))]  # 接上一行

    [w_sheet.write(i, base + 10, j) for i, j in enumerate(_type_cars_count.keys())]  # 写入车型车辆数目数据
    [w_sheet.write(i, base + 11, _type_cars_count[j]) for i, j in enumerate(_type_cars_count.keys())]

    [w_sheet.write(i, base + 12, j) for i, j in enumerate(_mileage_result.keys())]  # 写入累计里程区间统计数据
    [w_sheet.write(i, base + 13, _mileage_result[j]) for i, j in enumerate(_mileage_result.keys())]

    # 写入月度上线天数区间统计数据 该字典写在累计里程数据下方隔一行
    [w_sheet.write(i + len(_mileage_result) + 1, base + 12, j) for i, j in enumerate(_days_online.keys())]
    [w_sheet.write(i + len(_mileage_result) + 1, base + 13, _days_online[j]) for i, j in enumerate(_days_online.keys())]

    saveas_name = _filename.partition('.')[0] + '-统计结果.xls'
    wbook.save(saveas_name)  # 处理完的表格另存


# 信息写入表格及表格另存（.xlsx格式）
def write_result_xlsx(_filename: str, _data: list,
                      _province_result: dict, _day_result: dict,
                      _dedupliction: list, _type_cars_count: dict,
                      _mileage_result: dict, _days_online: dict) -> None:  # 接上一行
    """
    :param _filename: 文件名
    :param _data: 全部数据（不包括data[10]和data[11]）
    :param _province_result: 省份数据字典{省份: [各省车辆数目,各省月总里程]}
    :param _day_result: 每日数据字典{日期：[每天上线车辆数，每天总里程]}
    :param _dedupliction: 去重后的数据列表[vin, 车型, 运营城市, 车辆月度上线天数]
    :param _type_cars_count: 字典{车型：各车型车辆数目}
    :param _mileage_result: 字典{仪表里程(KM): 车辆数目}
    :param _days_online:字典{月度上线天数区间：车辆数目}
    :return: 无
    """
    table = xlsxwriter.Workbook(_filename.partition('.')[0] + '-统计结果.xlsx')
    w_sheet = table.add_worksheet('Sheet1')

    base = 10  # 统计结果起始写入列

    [w_sheet.write_column(0, i, _data[i]) for i in range(10)]

    [w_sheet.write(i, base, j) for i, j in enumerate(_province_result.keys())]  # 写入省份数据
    [w_sheet.write(i, base + 1, _province_result[j][0]) for i, j in enumerate(_province_result.keys())]
    [w_sheet.write(i, base + 2, _province_result[j][1]) for i, j in enumerate(_province_result.keys())]

    [w_sheet.write(i, base + 3, j) for i, j in enumerate(_day_result.keys())]  # 写入每日数据
    [w_sheet.write(i, base + 4, _day_result[j][0]) for i, j in enumerate(_day_result.keys())]
    [w_sheet.write(i, base + 5, _day_result[j][1]) for i, j in enumerate(_day_result.keys())]

    [w_sheet.write(j, base + 6 + i, _dedupliction[i][j]) for i in range(len(_dedupliction))  # 写入去重列表（4列）
     for j in range(len(_dedupliction[i]))]  # 接上一行

    [w_sheet.write(i, base + 10, j) for i, j in enumerate(_type_cars_count.keys())]  # 写入车型车辆数目数据
    [w_sheet.write(i, base + 11, _type_cars_count[j]) for i, j in enumerate(_type_cars_count.keys())]

    [w_sheet.write(i, base + 12, j) for i, j in enumerate(_mileage_result.keys())]  # 写入累计里程区间统计数据
    [w_sheet.write(i, base + 13, _mileage_result[j]) for i, j in enumerate(_mileage_result.keys())]

    # 写入月度上线天数区间统计数据 该字典写在累计里程数据下方隔一行
    [w_sheet.write(i + len(_mileage_result) + 1, base + 12, j) for i, j in enumerate(_days_online.keys())]
    [w_sheet.write(i + len(_mileage_result) + 1, base + 13, _days_online[j]) for i, j in enumerate(_days_online.keys())]

    # saveas_name = _filename.partition('.')[0] + '-统计结果.xls'
    # wbook.save(saveas_name)  # 处理完的表格另存
    table.close()


# 总执行函数
def main() -> None:
    start0 = datetime.datetime.now()
    print('执行中，请稍等......')
    targetfile = getxls()
    for file in targetfile:
        start = datetime.datetime.now()
        data, data_rows, data_columns = form_init(file)  # 调用初始化函数导入数据
        province_result, dedupliction = province_(data, data_rows)  # 各省车辆数及各省月总里程
        day_result = day_(data, data_rows)  # 每天上线车辆数及每天里程
        type_cars_count_ = type_cars_count(dedupliction)
        mile = mileage_(data, dedupliction[0])
        days_online_ = days_online(dedupliction[3])

        write_result_xlsx(file, data, province_result, day_result, dedupliction, type_cars_count_, mile, days_online_)  # 写入
        print('{} 用时 {} 秒'.format(file.partition('.')[0] + '-统计结果.xlsx', (datetime.datetime.now() - start).total_seconds()))

    print('成功！总用时 {} 秒'.format((datetime.datetime.now() - start0).total_seconds()))


main()

