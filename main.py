# -*- coding: utf8 -*-
import xlwings as xw
import pandas as pd
from tkinter import *
import gui, search_in_sh as sc
import string
import re
# import os

def get_column_name(columnindex):
    ret = ''
    ci = columnindex - 1
    index = ci // 26
    if index > 0:
        ret += get_column_name(index)
    ret += string.ascii_uppercase[ci % 26]

    return ret

# 定义BOM表名称
# BOM = "WG_NILMV02.xlsx"
# BOM_sheet = "WG_NILMV02"
BOM = ''
BOM_sheet = ''
export_address = ''

comment = []
comment_len = 0
footprint = []
# footprint_len = 0
def read_bom():
    global comment
    global comment_len
    global footprint
    # global footprint_len
    # 读取BOM表里的数据
    # app.display_alerts = False    # 关闭一些提示信息，可以加快运行速度。 默认为 True。
    # app.screen_updating = True    # 更新显示工作表的内容。默认为 True。关闭它也可以提升运行速度。
    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(BOM)
    # 打开sheet
    sht = wb.sheets[0]
    # 拿取sheet的comment数据
    nrows = sht.used_range.last_cell.row    # 获取总行数

    # 获取comment和footprint列标
    first_line = sht.range('A1:J1').value
    comment_index = first_line.index('Comment') + 1
    comment_col = get_column_name(comment_index)
    footprint_index = first_line.index('Footprint') + 1
    footprint_col = get_column_name(footprint_index)

    # if sht.range('A1').value == '':
    #     comment_col = 'C'
    #     footprint_col = 'G'
    # else:
    #     comment_col = 'A'
    #     footprint_col = 'E'
    comment = sht.range(comment_col + '2:' + comment_col + str(nrows)).value    # 从A2开始拿数据，跳过表头
    comment_len = len(comment)
    footprint = sht.range(footprint_col + '2:' + footprint_col + str(nrows)).value   # 从E2开始拿数据，跳过表头
    # footprint_len = len(footprint)

    # print(len(comment))
    # print(type(comment_len))
    # print(comment)

    # 关闭工作薄
    wb.save()
    wb.close()
    app.quit()

# --------------------------------------------------
wuliao_script = []
wuliao_script_len = 0
wuliao_id = []
# wuliao_id_len = 0
# A002 = ''
# 读取仓库里的物料描述（D列）和物料号（C列）
def read_warehouse(adress):
    global wuliao_script
    global wuliao_script_len
    global wuliao_id
    # global wuliao_id_len
    # global A002
    # 定义A002仓库的名称
    # A002 = "A002采购数据11.10.xlsx"
    # A002_sheet = "Sheet1"

    app = xw.App(visible=False, add_book=False)
    wb = app.books.open(adress)
    # 打开sheet
    sht = wb.sheets[0]
    # 拿取sheet的物料描述和物料编号
    nrows = sht.used_range.last_cell.row   # 获取总行数
    wuliao_script = sht.range('D2:' + 'D' + str(nrows)).value
    wuliao_script_len = len(wuliao_script)
    wuliao_id = sht.range('C2:' + 'C' + str(nrows)).value
    # wuliao_id_len = len(wuliao_id)
    deal_wuliao_script()

    # print(type(wuliao_script_len))
    # print(wuliao_script_len)
    # print(wuliao_script[0])
    # print(wuliao_script[3182])

    # 关闭工作薄
    wb.save()
    wb.close()
    app.quit()

# 处理物料表数据,挑出所有电容和电阻，替换μF、R、
wuliao_cap = []
wuliao_cap_id = []
wuliao_res = []
wuliao_res_id = []

def replace_uf(old_script):
    if ('μF' in old_script):
        new_script = re.sub('μ', 'u', old_script, re.I)
        return new_script
    else:
        return old_script

def replace_R(old_script):
    # 匹配类似2R或5R1类型
    value_num = re.findall('(\d+)', old_script, re.I) # value[0]
    if len(value_num) == 1:
        new_script = value_num[0] + 'Ω'
        return new_script
    elif len(value_num) == 2:
        if value_num[1] == '0':
            new_script = value_num[0] + 'Ω'
            return new_script
        else:
            new_script = value_num[0] + '.' + value_num[1] + 'Ω'
            return new_script

def deal_wuliao_script():
    # global wuliao_script_len
    wuliao_cap.clear()
    wuliao_cap_id.clear()
    wuliao_res.clear()
    wuliao_res_id.clear()
    for i in range(wuliao_script_len):
        if ('电容' in wuliao_script[i]):
            wuliao_script[i] = replace_uf(wuliao_script[i])
            wuliao_cap.append(wuliao_script[i])
            wuliao_cap_id.append(wuliao_id[i])

        if('电阻' in wuliao_script[i]):
            value = re.findall('\d+\.?\d?R\d?\W{1}', wuliao_script[i], re.I) # \d+\.?\d?R\d?
            if value:
                new_value = replace_R(value[0]) + value[0][-1]
                wuliao_script[i] = wuliao_script[i].replace(value[0], new_value)
            wuliao_res.append(wuliao_script[i])
            wuliao_res_id.append(wuliao_id[i])

# --------------------------------------------------
# 查找BOM表里的comment项和footprint项对应仓库的wuliao_script项
# dict = {}   # 创建一个字典保存查找到的元件信息
dict_A002 = {}     # 用来存储第一次匹配到的物料信息
dict_C016 = {}     # 用来存储第二次匹配到的物料信息
has_searched_index = []     # 已经找到的元件索引
res_or_cap_index = []     # 搜索过的电容或电阻元件索引

# 挑出空值的元件
def search_null():
    for i in range(comment_len):
        if not comment[i]:
            has_searched_index.append(i)
        else:
            continue

# 匹配电阻
def search_res():
    global has_searched_index
    dict_res = {}
    for i in range(comment_len):
        # 1、匹配电阻阻值
        # 先将comment的数值处理一下，只得到阻值和单位如“0R”
        # pattern = re.compile(r'\d+[RΩkM]')
        if i not in has_searched_index:
            value_searched = re.match(r'[0-9]+[.]?[0-9]?[RΩkMH]+\d?', str(comment[i]), re.I)  # re.I 不区分大小写,\d+ 1个或多个数字，[RΩ] 匹配R或Ω    [0-9]+[RΩkM]+\d?    [0-9]+([.]{1}[0-9]+)?[RΩkM]+\d?
            if value_searched:
                value_searched = value_searched.group()
                if 'H' in value_searched or 'h' in value_searched:
                    continue
                elif 'R' in value_searched:
                    value_searched = replace_R(value_searched)
                if '.' in value_searched:
                    value_searched = value_searched.replace('.', r'\.')
                if 'kΩ' in value_searched or 'MΩ' in value_searched or 'KΩ' in value_searched or 'mΩ' in value_searched:
                    value_searched = value_searched.replace('Ω', '')

                # 2、匹配电阻封装
                footprint_searched = re.search(r'[\d]{4}', footprint[i], re.I)
                if footprint_searched:
                    res_or_cap_index.append(i)
                    footprint_searched = footprint_searched.group()

                    # 3、到search_component_in_A002包中查找电阻
                    component_list = sc.search_res_in_a002(value_searched, footprint_searched)
                    if component_list:
                        has_searched_index.append(i)
                        dict_res[i + 2] = component_list

            else:
                continue

    return dict_res

# -----------------------------------------------------------------------------
# 匹配电容
def num_to_cap(value):
    cap_dict = {'102': '1nF',
                '103': '10nF',
                '104': '100nF',
                '105': '1uF',
                '106': '10uF',
                '223': '22nF',
                '224': '220nF',
                }
    if value in cap_dict.keys():
        return cap_dict[value]
    else:
        return 0
def search_cap():
    global has_searched_index
    dict_cap = {}
    for i in range(comment_len):
        # 先将comment的数值处理一下，只得到容值和单位如“1nF”
        if i not in has_searched_index:
            value_searched = re.search(r'[0-9]+([.]{1}[0-9]+){0,1}(nF|pF|uF|μF|F)', str(comment[i]), re.I) # [0-9]+([.]{1}[0-9]+){0,1}  匹配整数或小数
            footprint_searched = re.search(r'[\d]{4}', str(footprint[i]), re.I)  # [0-9]+([.]{1}[0-9]+){0,1}表示一个整数或小数
            # 匹配贴片电容***********************************************
            # 1、匹配电容容值
            if value_searched:
                # res_or_cap_index.append(i)
                value_searched = replace_uf(value_searched.group())
                if '.' in value_searched:
                    value_searched = value_searched.replace('.', '\.')
            elif (len(str(comment[i])) == 3) & (str(comment[i]).isdigit()):
                value = num_to_cap(comment[i])
                res_or_cap_index.append(i)
                if value:
                    value_searched = value
                else:
                    continue
            else:
                continue

            # 2、匹配贴片电容封装
            if footprint_searched:
                footprint_tan = re.search(r'[^a-zA-Z]+([ABCDE])[^a-zA-Z]*', footprint[i], re.I)
                if not footprint_tan:
                    res_or_cap_index.append(i)
                # res_or_cap_index.append(i)
                footprint_searched = footprint_searched.group()
            else:
                continue

            # 4、到search_component_in_A002包中查找电容
            component_list = sc.search_cap_in_a002(value_searched, footprint_searched)
            if component_list:
                has_searched_index.append(i)
                dict_cap[i + 2] = component_list
            else:
                continue

    return dict_cap
    # print(dict)


# 查找其它未匹配的元器件，拿到对应bom表序号和关键字
unsearched_index = []   # 其他未匹配的元件对应bom表的序号
unsearched_key = {}     # 其他未匹配的元件的关键字，用来到物料库中查找
def get_other():
    global has_searched_index
    unsearched_index.clear()
    unsearched_key.clear()
    for i in range(comment_len):    # 根据comment值匹配所有可能的元件，手动选择
        value = str(comment[i])
        if i in has_searched_index or i in res_or_cap_index:
            continue
        elif str(comment[i])[0].isdigit():  # 如果第一个字符是数字
            # 匹配钽电容或电解电容************************************
            cap = re.search(r'[0-9]+([.]{1}[0-9]+){0,1}(nF|pF|uF|μF|F)+', str(comment[i]), re.I)  # [0-9]+([.]{1}[0-9]+){0,1}  匹配整数或小数
            if cap:
                cap = cap.group()
                if '.' in cap:
                    cap = cap.replace('.', '\.')
                # voltage = re.search(r'[0-9]+V', comment[i], re.I)
                footprint_tan = re.search(r'[^a-zA-Z]+([ABCDE])[^a-zA-Z]*', footprint[i], re.I)
                if footprint_tan:
                    key = '电容' + '.*' + cap + '.*' + footprint_tan.group(1)
                    unsearched_key[i] = key
                    unsearched_index.append(i)
                    continue

                # 3、匹配电解电容封装
                footprint_dianjie = re.search(r'([0-9]+([.]{1}[0-9]+){0,1}[*×xX][0-9]+([.]{1}[0-9]+){0,1})', footprint[i], re.I)  # [0-9]+([.]{1}[0-9]+){0,1}表示一个整数或小数
                if footprint_dianjie:
                    key = '电解电容' + '.*' + cap
                    unsearched_key[i] = key
                    unsearched_index.append(i)
                    continue

            # 匹配压敏电阻************************************
            yamin_res = re.search(r'(\d{2}[KD]+\d{3})|(\d{3}[KD]+\d{2})', str(comment[i]), re.I)
            if yamin_res:
                yamin = yamin_res.group()
                if yamin:
                    key = '压敏电阻' + '.*' + re.search('\d{3}', yamin, re.I).group()
                    unsearched_key[i] = key
                    unsearched_index.append(i)
                    continue

            # value = str(comment[i])
            # 匹配晶振
            if 'Hz' in value or 'hz' in value:
                key = re.search(r'[0-9]+([.]{1}[0-9]+){0,1}', value, re.I)
                if key:
                    key = '晶振.*' + key.group()
                    unsearched_index.append(i)
                    unsearched_key[i] = key
                    continue
            # 匹配电感
            elif 'mH' in value or 'uH' in value or 'H' in value or 'μH' in value:
                key = re.search(r'[0-9]+([.]{1}[0-9]+){0,1}.*H', value, re.I)
                if key:
                    key = '电感.*' + key.group()
                    unsearched_index.append(i)
                    unsearched_key[i] = key
                    continue

            # 匹配没有匹配到的电阻
            key_res = re.match(r'[0-9]+[.]?[0-9]?[RΩkM]+\d?', value, re.I)
            if key_res:
                key_res = key_res.group()
                # if 'H' in key_res or 'h' in key_res:
                #     continue
                key_res = replace_R(key_res)
                if '.' in key_res:
                    key_res = key_res.replace('.', r'\.')
                key_res = '电阻.*' + key_res
                unsearched_index.append(i)
                unsearched_key[i] = key_res
                continue

            # 匹配长度小于4，不在电阻或电容内的
            elif i not in res_or_cap_index:
                if len(value) <= 4:
                    unsearched_index.append(i)
                    unsearched_key[i] = value
                    continue
                else:
                    unsearched_index.append(i)
                    unsearched_key[i] = value[0:4]
                    continue

        else:   # 第一个字符不是数字
            # 不匹配查找过的电阻电容
            if i in res_or_cap_index:
                continue
            key = re.search(r'^[a-zA-Z]+[\d]+', value, re.I)  # 匹配以字母开头，以数字结尾

            if key:
                key = key.group()
                if (len(key) < 4) & (key.isdigit()):
                    continue
                else:
                    unsearched_index.append(i)
                    unsearched_key[i] = key
                    continue
            # 匹配长度小于4
            if (len(value) <= 4):
                unsearched_index.append(i)
                unsearched_key[i] = value
                continue
            # 匹配长度大于4
            if (len(value) > 4):
                unsearched_index.append(i)
                unsearched_key[i] = value[0:3]
                continue

# pandas读出BOM表
def read_BOM_dataframe():
    df_BOM = pd.read_excel(BOM, engine= 'openpyxl')     # sheet_name=BOM_sheet
    return df_BOM

# 字典值写入到dataframe
# def dict_dataframe():
#     df_dict = pd.DataFrame(index= range(0, comment_len+1), columns= ['物料号', '物料描述'])  # 创建一个dataframe
#     for i in dict.keys():   # 将电阻和电容信息添加到dataframe
#         df_dict.loc[i, '物料号'] = dict[i][0]
#         df_dict.loc[i, '物料描述'] = dict[i][1]
#     return df_dict

# 拼合DataFrame写出到excel
# def write_to_excel():
#     bom_dataframe = read_BOM_dataframe()
#     # 字典值写入到dataframe
#     df_dict = pd.DataFrame(index=range(0, comment_len + 1), columns=['物料号', '物料描述'])  # 创建一个dataframe
#     for i in dict.keys():  # 将dict{}信息添加到dataframe
#         df_dict.loc[i-2, '物料号'] = dict[i][0]
#         df_dict.loc[i-2, '物料描述'] = dict[i][1]
#     # BOM表拼合到字典值后面
#     df_contact = pd.concat([df_dict, bom_dataframe], axis=1, ignore_index=False)
#
#     df_contact.to_excel(export_address, sheet_name=BOM_sheet, index=False)

# # 定义第一次匹配和第二次匹配的物料仓库的名字
# first_warehouse = ''
# second_warehouse = ''

def write_to_excel(adress):
    # 建立表头
    df_bom = read_BOM_dataframe()
    bom_heard = list(df_bom)
    column = ['物料号', '物料描述'] + bom_heard
    df_result = pd.DataFrame(columns=column)

    if dict_A002:
        # 将第一次A002仓库匹配到的dict写入到df_result
        # 添加仓库名字分割行
        df_result.loc[df_result.index.size] = {'物料号': 'A002仓库'}
        first_searched = []
        for i in dict_A002.keys():
            dict_list = dict_A002[i]   # [物料号，物料描述]
            bom_list = list(df_bom.iloc[i-2])  # bom表的行转为列表
            component_list = dict_list + bom_list   # 查找到的物料信息＋原bom的行
            first_searched.append(component_list)
        df_first_searched = pd.DataFrame(first_searched, columns=column)  # 创建一个dataframe
        # 将df_first_searched和df_all添加到df_result
        df_result = pd.concat([df_result, df_first_searched], ignore_index=True)

    # ---------------------------------------------------------
    if dict_C016:
        # 添加仓库名字分割行
        df_result.loc[df_result.index.size] = {'物料号': 'C016仓库'}
        # 将第二次C016仓库匹配到的添加到df_result后面
        second_searched = []
        for i in dict_C016.keys():
            dict_list = dict_C016[i]  # [物料号，物料描述]
            bom_list = list(df_bom.iloc[i - 2])  # bom表的行转为字典
            component_list = dict_list + bom_list  # 查找到的物料信息＋原bom的行
            second_searched.append(component_list)
        df_second_searched = pd.DataFrame(second_searched, columns=column)  # 创建一个dataframe
        # 将df_first_searched和df_all添加到df_result
        df_result = pd.concat([df_result, df_second_searched], ignore_index=True)

    # ---------------------------------------------------------
    # 将未匹配到的添加到df_result后面
    unsearched = []
    df_result.loc[df_result.index.size] = {'物料号': '购买项'}
    for i in range(comment_len):
        if ((i+2) in dict_A002.keys()) or ((i+2) in dict_C016.keys()):
            continue
        else:
            bom_unsearched_list = list(df_bom.iloc[i])  # bom表的行转为列表
            unsearched.append(bom_unsearched_list)
            continue

    df_unsearched = pd.DataFrame(unsearched, columns=bom_heard)  # 创建一个dataframe

    # 将df_A002_unsearched和df_all添加到df_result
    df_result = pd.concat([df_result, df_unsearched], ignore_index=True)

    # 写入到excel
    df_result.to_excel(adress, sheet_name=BOM_sheet, index=False)

    # # 设置excel格式
    # app = xw.App(visible=False, add_book=False)
    # wb = app.books.open(export_address)
    # sht = wb.sheets[BOM_sheet]  # 打开sheet
    # rows = sht.used_range.last_cell.row  # 获取总行数
    # columns = sht.used_range.last_cell.column  # 获取总列数
    #
    # range_all = sht.range((1, 1), (rows, columns))
    # range_all.row_height = 20
    # range_all.api.VerticalAlignment = -4108  # -4108 垂直居中（默认）。 -4160 靠上，-4107 靠下， -4130 自动换行对齐。
    # range_all.api.Font.Name = 'Arial'
    #
    # # 第一列
    # range_1 = sht.range((1, 1), (rows, 1))
    # range_1.column_width = 13
    # # 第二列
    # range_2 = sht.range((1, 2), (rows, 2))
    # range_2.column_width = 37
    # # 第三列
    # range_3 = sht.range((1, 3), (rows, 3))
    # range_3.column_width = 12
    # # 后面列
    # range_l = sht.range((1, 4), (rows, columns))
    # range_l.column_width = 20
    #
    # range_all.api.WrapText = True
    # range_all.api.EntireRow.AutoFit()
    # # sht.autofit()
    #
    # # 设置边框
    # # Borders(11) 垂直边框，LineStyle = 1 直线。
    # range_all.api.Borders(11).LineStyle = 1
    # # range_all.api.Borders(9).Weight = 3                # 设置边框粗细。
    #
    # # Borders(12) 水平框，LineStyle = 2 虚线。
    # range_all.api.Borders(12).LineStyle = 1
    # # range_all.api.Borders(7).Weight = 3
    #
    # for i in range(1, rows + 1):
    #     if sht.range(i, 1).value == 'A002仓库' or sht.range(i, 1).value == 'C016仓库' or sht.range(i, 1).value == '购买项':
    #         sht.range(i, 1).api.Font.Size = 14
    #         sht.range(i, 1).row_height = 25
    #         # sht.range(i, 1).api.Font.Bold = True
    #         sht.range((i, 1), (i, columns)).merge()
    #
    # wb.save()
    # wb.close()
    # app.kill()

if __name__ == '__main__':
    root = Tk()
    root.geometry('750x450+200+100')
    # root.minsize(100,50)
    root.resizable(0, 0)
    root.title('物料查找')
    # root.iconbitmap('D:\search\图标.ico')
    app = gui.My_gui(master=root)
    root.mainloop()
