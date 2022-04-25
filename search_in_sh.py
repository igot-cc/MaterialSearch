import re
import main


# from search import wuliao_script_len, wuliao_script, wuliao_id


# ------------------------------------------------------------------
# 函数功能：在A002中查找电阻
# 参数： value_searched  要查找的元件阻值
#      footprint_searched  要查找的元件封装
# 返回值：('wuliao_id', 'wuliao_script')，如('6090101125', '贴片电阻/0R/0603【彩源】')
#       0：没有查找到数据
# ------------------------------------------------------------------


def search_res_in_a002(value_searched, footprint_searched):
    # 3、到A002中的wuliao_res匹配阻值和封装
    # 循环每一个描述的字符串进行匹配
    list_temp_all = []
    pattern_value = r'[^0-9\.]' + value_searched + r'\D'  # '\D': 匹配任意非数字
    pattern_footprint = r'\D' + footprint_searched
    for index in range(len(main.wuliao_res)):
        # 先匹配阻值
        # value_searched_boolean  = re.search(value_searched,wuliao_script[index])
        # print(pattern_value)
        if re.search(pattern_value, main.wuliao_res[index], re.I):
            # 再匹配封装
            if re.search(pattern_footprint, main.wuliao_res[index], re.I):
                list_temp = [main.wuliao_res_id[index], main.wuliao_res[index]]
                list_temp_all.append(list_temp)

    # 依次挑选含1%(F)或5%(J)的电阻
    for i in range(len(list_temp_all)):
        if ('1%' in list_temp_all[i][1]) | ('F' in list_temp_all[i][1]):
            # print(list_temp_all[i])
            return list_temp_all[i]
    for i in range(len(list_temp_all)):
        if ('5%' in list_temp_all[i][1]) | ('J' in list_temp_all[i][1]):
            return list_temp_all[i]
        else:
            return list_temp_all[i]

    return 0

# ------------------------------------------------------------------
# 函数功能：在A002中查找贴片电容
# 参数： value_searched  要查找的元件容值
#      footprint_searched  要查找的元件封装
# 返回值：('wuliao_id', 'wuliao_script')
#       0：没有查找到数据
# ------------------------------------------------------------------

def search_cap_in_a002(value_searched, footprint_searched):
    # 3、到A002中的wuliao_script匹配阻值和封装
    # 循环每一个描述的字符串进行匹配
    list_temp_all = []
    pattern_value = r'[^0-9\.]' + value_searched  # + '\D'  # '\D': 匹配任意非数字
    pattern_footprint = r'\D' + footprint_searched
    for index in range(len(main.wuliao_cap)):
        # 先匹配容值
        # value_searched_boolean  = re.search(value_searched,wuliao_script[index])
        # print(pattern_value)
        if re.search(pattern_value, main.wuliao_cap[index], re.I):
            # 再匹配封装
            if re.search(pattern_footprint, main.wuliao_cap[index], re.I):
                list_temp = [main.wuliao_cap_id[index], main.wuliao_cap[index]]
                list_temp_all.append(list_temp)
    # print(list_temp_all)

    # 依次挑选含50V或35V或25V或16V的电容
    for i in range(len(list_temp_all)):
        if '50V' in list_temp_all[i][1]:
            # print(list_temp_all[i])
            return list_temp_all[i]
    for i in range(len(list_temp_all)):
        if '35V' in list_temp_all[i][1]:
            return list_temp_all[i]
    for i in range(len(list_temp_all)):
        if '25V' in list_temp_all[i][1]:
            return list_temp_all[i]
        else:
            return list_temp_all[i]
                # return wuliao_id[index], wuliao_script[index]
    return 0

# ------------------------------------------------------------------
# 函数功能：在A002中查找其他元器件
# 参数： value_searched  要查找的元件阻值
#      footprint_searched  要查找的元件封装
# 返回值：('wuliao_id', 'wuliao_script')，如('6090101125', '贴片电阻/0R/0603【彩源】')
#       0：没有查找到数据
# ------------------------------------------------------------------
def search_other_in_a002(value):
    # 3、到A002中的wuliao_script匹配物料描述
    # 循环每一个描述的字符串进行匹配
    list_temp_all = []
    if '电容' in value:
        for index in range(len(main.wuliao_cap)):
            has_wuliao = re.search(value, main.wuliao_cap[index], re.I)
            # if value in main.wuliao_script[index]:
            if has_wuliao:
                list_temp = [main.wuliao_cap_id[index], main.wuliao_cap[index]]
                list_temp_all.append(list_temp)
    elif '电阻' in value:
        for index in range(len(main.wuliao_res)):
            has_wuliao = re.search(value, main.wuliao_res[index], re.I)
            # if value in main.wuliao_script[index]:
            if has_wuliao:
                list_temp = [main.wuliao_res_id[index], main.wuliao_res[index]]
                list_temp_all.append(list_temp)
    else:
        for index in range(main.wuliao_script_len):
            has_wuliao = re.search(value, main.wuliao_script[index], re.I)
            # if value in main.wuliao_script[index]:
            if has_wuliao:
                list_temp = [main.wuliao_id[index], main.wuliao_script[index]]
                list_temp_all.append(list_temp)
    return list_temp_all

