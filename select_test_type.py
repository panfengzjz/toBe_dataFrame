# coding: utf-8
import numpy as np
import pandas as pd
from pandas import DataFrame

def get_name_set(data):
    return set(data["姓名"])

#筛选某一列中的特定字符，并保存到新表格中
def split_origin_data(data, saveFileName):
    test_item = set(data['项目代号'])
    writer = pd.ExcelWriter(saveFileName)

    for i in test_item:
        df = data[data['项目代号']==i]
        df.to_excel(writer, i)
    writer.save()
    
def search_data_and_give_result(fileName, nameSet):
    writer = pd.ExcelWriter("new_"+fileName)
    data = pd.DataFrame()
    data = DataFrame({"姓名":list(nameSet)})
    data.set_index("姓名", inplace=True)

    f = pd.ExcelFile(fileName)
    for i in f.sheet_names:
        data[i] = np.nan
        if (i == "ACTH") or (i == "CORT"):
            #暂时先不操作，存在比较大的问题
            #data[i+"_备注1"] = np.nan
            #data[i+"_备注"] = np.nan
            continue
        d = pd.read_excel(fileName, sheetname=i)
        d.set_index("姓名", inplace=True)
        for name in d.index:
            cell = d.loc[name, "数字结果"]
            try:
                data.loc[name, i] = cell.values[-1]
            except AttributeError:
                data.loc[name, i] = cell

    data.to_excel(writer)
    writer.save()

if __name__ == "__main__":
    fileName = "肝穿与DM激素.xlsx"

    sheetName = "糖尿病"
    data = pd.read_excel(fileName, sheet_name=sheetName)
    nameSet = get_name_set(data)
    #split_origin_data(data, sheetName+".xlsx")
    search_data_and_give_result(sheetName+".xlsx", nameSet)
    
    #sheetName = "肝穿"
    #data = pd.read_excel(fileName, sheet_name=sheetName)
    #nameSet = get_name_set(data)
    #split_origin_data(data, sheetName+".xlsx")
    #search_data_and_give_result(sheetName+".xlsx", nameSet)
