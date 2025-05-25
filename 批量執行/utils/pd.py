import pandas as pd
from itertools import islice
import numpy as np

def read_excel_and_process(file_path):
    data_list = []
    df = pd.read_excel(file_path, sheet_name="創建參數")
    df = df.fillna("")
    df.columns = ["參數名稱", "顯示名稱", "資料型態", "計量值", "取樣數量", "下限MIN", "期望值", "上限MAX", "單位","狀態"]


    for _, row in df.iterrows():
        data_row = {
            "param_name": row["參數名稱"],
            "param_name2": row["顯示名稱"],
            "data_type": row["資料型態"],
            "value": row["計量值"],
            "get_quan": row["取樣數量"],
            "limit_low": row["下限MIN"],
            "fit": row["期望值"],
            "limit_high": row["上限MAX"],
            "unit": row["單位"],
            "state": row["狀態"]
        }
        data_list.append(data_row)  
    return data_list

def read_excel_and_group_by_hierarchy(file_path):
    data_list = []
    df = pd.read_excel(file_path, sheet_name="工作站設定")
    df = df.fillna("")
    df.columns = ['件號', '工作站', '順序', '參數名稱', '重要性']
    df.set_index(['件號', '工作站'], inplace=True)
    grouped_data = {}
    for product, product_group in df.groupby('件號'):
        grouped_data[product] = {}
        for station, station_group in product_group.groupby('工作站'):
            grouped_data[product][station] = station_group.to_dict(orient="records")
    return grouped_data

def read_excel_and_auto_work_order(file_path):
    data_list = []
    df = pd.read_excel(file_path,sheet_name="自動工單")
    df = df.fillna("")
    df.colums = ['工單號碼']
    for index, row in df.iterrows():
        data_list.append(row['工單號碼'])    
    return data_list