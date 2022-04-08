'''
获取和保存json文件的excel初始地址

'''
import json
filename = "RecordFile.json"
#保存文件
def save(data_json):
    with open(filename, mode="w", encoding='UTF-8') as fn:
        json.dump(data_json, fn, indent=4, ensure_ascii=False)
        fn.close()
#读取文件
def read():
    with open(filename, mode="r", encoding='UTF-8') as fn:
        data_json = json.load(fn)
        fn.close()
        return data_json


#获取需要的地址
def get_address(excel_name):
    data_json = read()
    if excel_name == 'A002':
        return data_json['A002']
    elif excel_name == 'C016':
        return data_json['C016']
    elif excel_name == '研发仓':
        return data_json['研发仓']
    elif excel_name == 'BOM':
        return data_json['BOM']

#修改并存储地址
def modify(excel_name, new_address):
    data_json = read()
    if excel_name == 'A002':
        data_json['A002'] = new_address
    elif excel_name == 'C016':
        data_json['C016'] = new_address
    elif excel_name == '研发仓':
        data_json['研发仓'] = new_address
    elif excel_name == 'BOM':
        data_json['BOM'] = new_address
    save(data_json)

#修改并存储地址字典
def modify_by_dict(new_address_dict):
    data_json = read()
    for key,value in new_address_dict.items():
        if key == 'A002':
            data_json['A002'] = value
        elif key == 'C016':
            data_json['C016'] = value
        elif key == '研发仓':
            data_json['研发仓'] = value
        elif key == 'BOM':
            data_json['BOM'] = value
    save(data_json)

if __name__ == '__main__':
    data_json = {'A002': 'D:/A002采购数据.xlsx', 'C016': 'D:/C016采购数据.xlsx', '研发仓': 'D:/研发仓采购数据.xlsx', 'BOM': '123'}
    save(data_json)
    modify(['A002'], 22222)
