# coding: utf-8
import time
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
    print("split_origin_data finish")
    
def search_data_and_give_result(fileName, nameSet):
    writer = pd.ExcelWriter("new_"+fileName)
    data = pd.DataFrame()
    data = DataFrame({"姓名":list(nameSet)})
    data.set_index("姓名", inplace=True)

    f = pd.ExcelFile(fileName)
    d = pd.read_excel(fileName, sheet_name=f.sheet_names)
    for i in f.sheet_names:
        if (i == "ACTH") or (i == "CORT"):
            #暂时先不操作，存在比较大的问题
            #data[i+"_备注1"] = np.nan
            #data[i+"_备注"] = np.nan
            continue
        tmp_d = d.get(i)
        tmp_d.set_index("姓名", inplace=True)
        for name in tmp_d.index:
            cell = tmp_d.loc[name, "数字结果"]
            try:
                data.loc[name, i] = cell.values[-1]
            except AttributeError:
                data.loc[name, i] = cell

    data.to_excel(writer)
    writer.save()
    print("search_data_and_give_result finish")

if __name__ == "__main__":
    start_time = time.time()

    fileName = "内分泌科化验结果.xlsx"
    sheetName = "Sheet1"
    data = pd.read_excel(fileName, sheet_name=sheetName)
    print("load data complete")
    print("time cost: ", time.time() - start_time)
    nameSet = get_name_set(data)
    #split_origin_data(data, sheetName+".xlsx")
    #print("time cost: ", time.time() - start_time)
    search_data_and_give_result(sheetName+".xlsx", nameSet)
    print("time cost: ", time.time() - start_time)
    
    #sheetName = "肝穿"
    #data = pd.read_excel(fileName, sheet_name=sheetName)
    #nameSet = get_name_set(data)
    #split_origin_data(data, sheetName+".xlsx")
    #search_data_and_give_result(sheetName+".xlsx", nameSet)
