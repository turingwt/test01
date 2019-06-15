import datetime
import os
import xlsxwriter


# 遍历目录文件及子目录
def path_() -> tuple:
    """
    :return: dirs -> list 子目录列表
    :return: path -> str 当前目录字符串
    """
    path = os.getcwd()
    dirs = []
    for root, dirr, files in os.walk(path):
        dirs.append(root)
    dirs.pop(0)

    return dirs, path


# 将子目录中所有符合条件的'.xls'文件转换为'.txt'文件
def txt_trans(paths: list) -> None:
    """
    :param paths: list -> 子目录列表
    :return:
    """
    for path in paths:
        files = os.listdir(path)
        targetfile = [i for i in files if (i[0: 3] == 'LDY' and i[-4:] == '.xls')]
        os.chdir(path)
        if targetfile:
            print('{:<10} {}转换中。。'.format(path.rpartition('\\')[-1], ' ' * (7 - len(path.rpartition('\\')[-1]))))
            [os.rename(targetfile[i], targetfile[i].rpartition('.')[0] + '.txt') for i in range(len(targetfile))]


# 读取'.txt'文件
def txt_read(path: str) -> list:
    """
    :param path: 当前路径
    :return: TXT文件列表
    """
    files = os.listdir(path)
    os.chdir(path)  # 更改当前目录，否则会报错
    target = [i for i in files if i[-4:] == '.txt']
    return target


# 原始数据初始化处理
def data_init(filename: str) -> dict:
    """
    :param filename: 文件名（.txt）
    :return: 字典{'充电状态': data_charge, '累计里程': data_mileage, 'SOC': data_soc}
    """
    file = open(filename)
    file_init = file.readlines()
    target = [file_init[i].split() for i in range(len(file_init))]

    charge_index = target[0].index('充电状态') + 1
    mileage_index = target[0].index('累计里程') + 1
    soc_index = target[0].index('SOC') + 1

    avaible_data_index = []
    for i in range(1, len(target)):
        if len(target[i]) > mileage_index:
            if 0 <= float(target[i][soc_index]) <= 100:
                avaible_data_index.append(i)
    data_soc = [target[i][soc_index] for i in avaible_data_index]
    data_charge = [target[i][charge_index] for i in avaible_data_index]
    data_mileage = [target[i][mileage_index] for i in avaible_data_index]

    data_result = {'充电状态': data_charge, '累计里程': data_mileage, 'SOC': data_soc}  # 创建字典

    file.close()

    return data_result


# 能耗计算主函数
def milesge_aver(_data: dict) -> tuple:
    """
    :param _data: 字典{'充电状态': data_charge, '累计里程': data_mileage, 'SOC': data_soc}
    :return: _mil_aver_soc -> float 每1%SOC所行驶的里程
    :return: data_selected 二维列表 data_selected[0] -> 间隔为1的SOC数据列表 data_selected[1] -> 对应SOC的累计里程列表
    :return: a -> list 累计里程差值列表
    """
    soc_float = [float(_data['SOC'][i]) for i in range(len(_data['SOC']))]
    charge_float = [int(_data['充电状态'][i]) for i in range(len(_data['充电状态']))]
    mileage_float = [float(_data['累计里程'][i]) for i in range(len(_data['累计里程']))]

    data_selected = [[soc_float[0]], [mileage_float[0]]] if charge_float[0] != 1 else [[], []]
    soc_new = soc_float[0]
    for i in range(1, len(soc_float)):
        if charge_float[i] == 0:
            if soc_new >= soc_float[i]:
                if soc_new - soc_float[i] == 1:
                    data_selected[0].append(soc_float[i])
                    data_selected[1].append(mileage_float[i])
                    soc_new = soc_float[i]
        elif charge_float[i] == 1:
            soc_new = soc_float[i]
            if _data['SOC'][i] == 100.0:
                data_selected[0].append(soc_float[i])
                data_selected[1].append(mileage_float[i])
        elif charge_float[i] == 2:
            if charge_float[i - 1] != 2:
                data_selected[0].append(soc_float[i])
                data_selected[1].append(mileage_float[i])
                soc_new = soc_float[i]
            else:
                continue
    a = [(data_selected[1][i]) - (data_selected[1][i - 1]) for i in range(1, len(data_selected[0]))
         if (data_selected[0][i - 1]) - (data_selected[0][i]) == 1]
    mil_sum = 0
    _mil_aver_soc = 0.0
    flag_len_a = 0
    for i in a:
        mil_sum += i
    if not len(a):  # SOC无连续减1
        flag_len_a = 1
    else:
        _mil_aver_soc = mil_sum / len(a)

    return _mil_aver_soc, [data_selected, a], flag_len_a


# 统计结果写入表格
def xlsxwrite(result: dict, path_root: str, write_allcity_: dict) -> None:
    """
    :param result: 最终结果：{地区：该车型在该地区平均每度电所能行驶的里程}
    :param path_root: 主目录（创建统计结果表格用）
    :param write_allcity_: 每辆车的结果：{VIN：车辆平均每度电所能行驶的里程}
    :return:
    """
    table = xlsxwriter.Workbook(path_root + '\\能耗分析结果.xlsx')
    w_sheet = table.add_worksheet('Sheet1')
    base = 0

    [w_sheet.write(i, base + 0, j) for i, j in enumerate(result.keys())]
    [w_sheet.write(i, base + 1, result[j]) for i, j in enumerate(result.keys())]

    k = 0
    for i, j in enumerate(write_allcity_.keys()):
        w_sheet.write(k, base + 2, j)
        w_sheet.write(k, base + 3, '能耗(KM)')
        for m, n in enumerate(write_allcity_[j].keys()):
            w_sheet.write(k + m + 1, base + 2, n)
            w_sheet.write(k + m + 1, base + 3, write_allcity_[j][n])
        k = k + len(write_allcity_[j]) + 1

    table.close()


# 计算结果异常处理
def error_data(path: str, filename: str, errordata: list) -> None:
    """
    :param path: 数据异常的文件所在目录
    :param filename: 数据异常的文件的文件名
    :param errordata: 二维列表 [0] -> 间隔为1的SOC数据列表 [1] -> 对应SOC的累计里程列表
    :return:
    """
    table = xlsxwriter.Workbook(path + '_' + filename.rpartition('.')[0][0:17] + '_error_data.xlsx')
    w_sheet = table.add_worksheet('Sheet1')

    base = 0
    w_sheet.write(0, 0, 'SOC')
    w_sheet.write(0, 1, '累计里程')
    w_sheet.write(0, 2, '累计里程差值')
    [w_sheet.write(i + 1, base + 0, errordata[0][0][i]) for i in range(len(errordata[0][0]))]
    [w_sheet.write(i + 1, base + 1, errordata[0][1][i]) for i in range(len(errordata[0][1]))]
    [w_sheet.write(i + 1, base + 2, errordata[1][i]) for i in range(len(errordata[1]))]

    table.close()


# 主处理函数
def main():
    """
    :return: result  最终结果：{地区：该车型在该地区平均每度电所能行驶的里程}
    :return: path_root 主目录（创建统计结果表格用）
    :return: write_all_city 每辆车的结果：{VIN：车辆平均每度电所能行驶的里程}
    """
    paths, path_root = path_()
    txt_trans(paths)
    a = []
    result = {'地区': '地区能耗平均值'}
    write_allcity = {}
    for path in paths:
        files = txt_read(path)
        if not files:
            continue
        city = path.rpartition('\\')[-1]
        res = {}  # 'VIN': '每度电_里程(KM)'
        print(city, ':')
        for file in files:
            start_ = datetime.datetime.now()
            data = data_init(file)  # 数据清洗
            mil_aver_soc, data_error_need, flag_len_a = milesge_aver(data)
            if flag_len_a:
                continue
            mil_aver_quantity = round(mil_aver_soc * (100 / 133.51), 2)
            res[file[0:17]] = mil_aver_quantity
            if mil_aver_quantity > 5:
                print('\033[0;30;41m{} 能耗 -> {:<4},用时 {} 秒 数据异常,未计入结果！请检查车辆数据!\033[0m'.format
                      (file[0:17], mil_aver_quantity, (datetime.datetime.now() - start_).total_seconds()))  # 接上一行
                error_data(path, file, data_error_need)
            else:
                print('{} 能耗 -> {:<4}, 用时 {} 秒'.format(file[0:17], mil_aver_quantity,
                                                      (datetime.datetime.now() - start_).total_seconds()))  # 接上一行
        mil__aver_quantity_sum = 0.0
        k = 0
        for i, j in res.items():
            if j < 5:
                mil__aver_quantity_sum += j
                k += 1
        mil__aver_quantity_all = round(mil__aver_quantity_sum / k, 2)
        print('\033[0;30;44m{} -> 平均能耗：{} KM/KW.h\033[0m'.format(city, mil__aver_quantity_all))

        result[city] = mil__aver_quantity_all
        a.append(res)
        write_allcity[city] = res

    return result, path_root, write_allcity


start = datetime.datetime.now()
final, path_root_, write_all_city = main()
xlsxwrite(final, path_root_, write_all_city)
print('总用时 {} 秒！'.format((datetime.datetime.now() - start).total_seconds()))




