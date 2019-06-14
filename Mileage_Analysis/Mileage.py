import datetime
import os
import xlsxwriter


# 遍历目录文件及子目录
def path_() -> tuple:
    """
    :return: dirs -> list 子目录列表
    :return: path -> str 当前目录字符串
    """
    path = os.getcwd()  # 得到当前目录
    dirs = []  # 创建子目录空列表
    for root, dirr, files in os.walk(path):  # 返回生成器，因此不能使用列表生成式
        dirs.append(root)
    dirs.pop(0)  # 去除主目录，只包含子目录

    return dirs, path


# 将子目录中所有符合条件的'.xls'文件转换为'.txt'文件
def txt_trans(paths: list) -> None:
    """
    :param paths: list -> 子目录列表
    :return:
    """
    for path in paths:
        files = os.listdir(path)  # 列出当前目录下所有文件
        targetfile = [i for i in files if (i[0: 3] == 'LDY' and i[-4:] == '.xls')]  # 文件名开头为'LDY'，后缀'.xis'即为目标文件
        os.chdir(path)  # 更改当前目录，否则会报错
        if targetfile:  # 若包含待转换的文件，则
            print('{:<10} {}转换中。。'.format(path.rpartition('\\')[-1], ' ' * (7 - len(path.rpartition('\\')[-1]))))
            [os.rename(targetfile[i], targetfile[i].rpartition('.')[0] + '.txt') for i in range(len(targetfile))]  # 重命名


# 读取'.txt'文件
def txt_read(path: str) -> list:
    """
    :param path: 当前路径
    :return: TXT文件列表
    """
    files = os.listdir(path)
    os.chdir(path)  # 更改当前目录，否则会报错
    target = [i for i in files if i[-4:] == '.txt']  # 找到所有txt文件
    return target


# 原始数据初始化处理
def data_init(filename: str) -> dict:
    """
    :param filename: 文件名（.txt）
    :return: 字典{'充电状态': data_charge, '累计里程': data_mileage, 'SOC': data_soc}
    """
    file = open(filename)  # 打开文件
    file_init = file.readlines()  # 逐行读取全部内容，得到字符串列表（列表中的每一个元素都是长字符串）
    target = [file_init[i].split() for i in range(len(file_init))]  # 将每一行按空格分隔字符串，拼成列表（target列表中的每一个元素都是列表，即嵌套列表）

    charge_index = target[0].index('充电状态') + 1  # 找到充电状态所在列，因原表中日期和时间是同一列，但在上一行操作中被分成两列，故索引加1
    mileage_index = target[0].index('累计里程') + 1  # 找到累计里程所在列，因原表中日期和时间是同一列，但在上一行操作中被分成两列，故索引加1
    soc_index = target[0].index('SOC') + 1  # 找到SOC所在列，因原表中日期和时间是同一列，但在上一行操作中被分成两列，故索引加1

    avaible_data_index = []  # 找到数据完整的行（原表中有些数据项中的数据由于传输或其他原因导致数据不存在，为空格，这些数据行不计入）
    for i in range(1, len(target)):
        if len(target[i]) > mileage_index:  # 如果长度小于于第一行则不计入
            if 0 <= float(target[i][soc_index]) <= 100:  # 如果SOC范围不是0-100则不计入
                avaible_data_index.append(i)
    data_soc = [target[i][soc_index] for i in avaible_data_index]  # 符合上述条件的SOC列表
    data_charge = [target[i][charge_index] for i in avaible_data_index]  # 符合上述条件的充电状态列表
    data_mileage = [target[i][mileage_index] for i in avaible_data_index]  # 符合上述条件的累计里程列表

    data_result = {'充电状态': data_charge, '累计里程': data_mileage, 'SOC': data_soc}  # 创建字典

    file.close()  # 保存并关闭文件

    return data_result


# 能耗计算主函数
def milesge_aver(_data: dict) -> tuple:
    """
    :param _data: 字典{'充电状态': data_charge, '累计里程': data_mileage, 'SOC': data_soc}
    :return: _mil_aver_soc -> float 每1%SOC所行驶的里程
    :return: data_selected 二维列表 data_selected[0] -> 间隔为1的SOC数据列表 data_selected[1] -> 对应SOC的累计里程列表
    :return: a -> list 累计里程差值列表
    """
    soc_float = [float(_data['SOC'][i]) for i in range(len(_data['SOC']))]  # 将传入字典中的字符串列表转为数值列表
    charge_float = [int(_data['充电状态'][i]) for i in range(len(_data['充电状态']))]
    mileage_float = [float(_data['累计里程'][i]) for i in range(len(_data['累计里程']))]

    # 创建包含间隔为1的SOC列表的二维列表，第一行定义，若第一行数据为充电数据则不取
    data_selected = [[soc_float[0]], [mileage_float[0]]] if charge_float[0] != 1 else [[], []]    # 二维列表
    soc_new = soc_float[0]  # 用于比较的上一次SOC值
    for i in range(1, len(soc_float)):
        if charge_float[i] == 0:  # 未充电状态
            if soc_new >= soc_float[i]:  # 下一行SOC小于等于上一行SOC，即车辆处于耗电状态
                if soc_new - soc_float[i] == 1:  # 找到比上一次标记的SOC值小1的数据行
                    data_selected[0].append(soc_float[i])  # 得到SOC及累计里程，放入二维列表存储
                    data_selected[1].append(mileage_float[i])
                    soc_new = soc_float[i]  # 将标记的SOC值更新为本次SOC值
        elif charge_float[i] == 1:  # 充电中
            soc_new = soc_float[i]  # 将标记的SOC值更新为本次SOC值（充电时SOC一直上升，所以标记值一直更新，直到充完或充满）
            if _data['SOC'][i] == 100.0:  # 如果充到100%SOC
                data_selected[0].append(soc_float[i])  # 将此时的SOC及累计里程值记录
                data_selected[1].append(mileage_float[i])
        elif charge_float[i] == 2:  # 充电完成（有些车辆可能会不发或丢失充电完成状态数据，所以用充电中数据作为双保险）
            if charge_float[i - 1] != 2:  # 判断上一次是否记录，防止重复
                data_selected[0].append(soc_float[i])  # 记录充电完成时的SOC及累计里程
                data_selected[1].append(mileage_float[i])
                soc_new = soc_float[i]  # 将标记的SOC值更新为充电完成时的状态
            else:
                continue
    a = [(data_selected[1][i]) - (data_selected[1][i - 1]) for i in range(1, len(data_selected[0]))
         if (data_selected[0][i - 1]) - (data_selected[0][i]) == 1]  # 接上一行  #将相邻两行SOC差值为1的累计里程相减放入列表
    mil_sum = 0
    for i in a:
        mil_sum += i
    _mil_aver_soc = mil_sum / len(a)  # 得到该车辆的平均每1%SOC所能行驶的里程值

    return _mil_aver_soc, [data_selected, a]


# 统计结果写入表格
def xlsxwrite(result: dict, path_root: str, write_allcity_: dict) -> None:
    """
    :param result: 最终结果：{地区：该车型在该地区平均每度电所能行驶的里程}
    :param path_root: 主目录（创建统计结果表格用）
    :param write_allcity_: 每辆车的结果：{VIN：车辆平均每度电所能行驶的里程}
    :return:
    """
    table = xlsxwriter.Workbook(path_root + '\\能耗分析结果.xlsx')  # 创建统计结果表格
    w_sheet = table.add_worksheet('Sheet1')  # 新建Sheet1页
    base = 0  # 初始写入列

    # 写入地区平均值
    [w_sheet.write(i, base + 0, j) for i, j in enumerate(result.keys())]
    [w_sheet.write(i, base + 1, result[j]) for i, j in enumerate(result.keys())]

    # 写入每辆车的结果
    k = 0
    for i, j in enumerate(write_allcity_.keys()):
        w_sheet.write(k, base + 2, j)
        w_sheet.write(k, base + 3, '能耗(KM)')
        for m, n in enumerate(write_allcity_[j].keys()):
            w_sheet.write(k + m + 1, base + 2, n)
            w_sheet.write(k + m + 1, base + 3, write_allcity_[j][n])
        k = k + len(write_allcity_[j]) + 1

    table.close()  # 关闭文件（）必须关闭才能创建成功


# 计算结果异常处理
def error_data(path: str, filename: str, errordata: list) -> None:
    """
    :param path: 数据异常的文件所在目录
    :param filename: 数据异常的文件的文件名
    :param errordata: 二维列表 [0] -> 间隔为1的SOC数据列表 [1] -> 对应SOC的累计里程列表
    :return:
    """
    # 创建表格存储异常数据，每辆车一个表格，格式为 地区_VIN_error_data.xlsx
    table = xlsxwriter.Workbook(path + '_' + filename.rpartition('.')[0][0:17] + '_error_data.xlsx')
    w_sheet = table.add_worksheet('Sheet1')

    # 表格中存储间隔为1的SOC数据列表、对应SOC的累计里程列表、累计里程差值
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
    paths, path_root = path_()  # 得到子目录、主目录
    txt_trans(paths)  # 文件重命名预处理
    a = []
    result = {'地区': '地区能耗平均值'}  # 创建最终结果字典
    write_allcity = {}  # 创建每辆车结果字典
    for path in paths:  # 子目录循环
        files = txt_read(path)  # 子目录中目标文件列表
        if not files:  # 若该子目录中没有目标文件，则跳过本次循环，去下个子目录
            continue
        city = path.rpartition('\\')[-1]  # 得到地区名
        res = {}  # 'VIN': '每度电_里程(KM)'  每个地区车辆结果字典
        print(city, ':')  # 终端显示开始计算该地区
        for file in files:  # 子目录内循环处理所有目标文件
            start_ = datetime.datetime.now()
            data = data_init(file)  # 数据清洗
            mil_aver_soc, data_error_need = milesge_aver(data)  # 车辆能耗计算结果
            mil_aver_quantity = round(mil_aver_soc * (100 / 133.51), 2)  # SOC-电量换算，每度电行驶里程（需要的结果）
            res[file[0:17]] = mil_aver_quantity  # 字典{VIN：每度电行驶里程}
            if mil_aver_quantity > 5:  # 异常数据判断：若结果大于5则为异常
                print('\033[0;30;41m{} 能耗 -> {:<4},用时 {} 秒 数据异常,未计入结果！请检查车辆数据!\033[0m'.format
                      (file[0:17], mil_aver_quantity, (datetime.datetime.now() - start_).total_seconds()))  # 接上一行
                error_data(path, file, data_error_need)  # 转入异常处理函数
            else:  # 正常结果终端显示
                print('{} 能耗 -> {:<4}, 用时 {} 秒'.format(file[0:17], mil_aver_quantity,
                                                      (datetime.datetime.now() - start_).total_seconds()))  # 接上一行
        mil__aver_quantity_sum = 0.0  # 地区平均值计算
        k = 0
        for i, j in res.items():
            if j < 5:  # 异常结果不计入地区平均值计算
                mil__aver_quantity_sum += j
                k += 1
        mil__aver_quantity_all = round(mil__aver_quantity_sum / k, 2)  # 保留两位小数
        print('\033[0;30;44m{} -> 平均能耗：{} KM/KW.h\033[0m'.format(city, mil__aver_quantity_all))
        # 显示方式: 0（默认\）1（高亮）、22（非粗体）、4（下划线）、24（非下划线）、 5（闪烁）、25（非闪烁）、7（反显）、27（非反显）
        # 前景色:   30（黑色）、31（红色）、32（绿色）、 33（黄色）、34（蓝色）、35（洋 红）、36（青色）、37（白色）
        # 背景色:   40（黑色）、41（红色）、42（绿色）、 43（黄色）、44（蓝色）、45（洋 红）、46（青色）、47（白色）
        result[city] = mil__aver_quantity_all
        a.append(res)
        write_allcity[city] = res  # 嵌套字典{地区：{车辆：计算结果}。。}

    return result, path_root, write_allcity


start = datetime.datetime.now()
final, path_root_, write_all_city = main()
xlsxwrite(final, path_root_, write_all_city)
print('总用时 {} 秒！'.format((datetime.datetime.now() - start).total_seconds()))




