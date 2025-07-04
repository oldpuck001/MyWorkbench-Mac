# data_cleaning.py

# 数据清洗：对获取的数据进行清洗，确保其格式统一、无重复项。

import os
import shutil
import pandas as pd
import sqlite3

def data_cleaning_select(request):

    file_path = request.get("data", {}).get("file_path", "")

    sheet_file = pd.ExcelFile(file_path)                    # 使用pandas讀取Excel文件
    sheetnames = sheet_file.sheet_names                     # 獲取所有工作表名稱

    return ['data_cleaning_select', sheetnames]


def data_cleaning_index(request):

    file_path = request.get("data", {}).get("file_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")

    file_extension = os.path.splitext(file_path)[1].lower()

    if file_extension == '.xlsx':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    elif file_extension == '.xls':
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

    columns = df.columns.tolist()                           # 获取工作表的列名

    return ['data_cleaning_index', columns]


def data_cleaning_clean(request):
    
    file_path = request.get("data", {}).get("file_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")
    column_name = request.get("data", {}).get("column_name", "")
    cleaning_mode = request.get("data", {}).get("cleaning_mode", "")

    folder_path = os.path.dirname(file_path)
    target_path = os.path.join(folder_path, 'data_cleaning.xlsx')

    if os.path.exists(target_path):

        df = pd.read_excel(target_path, engine='openpyxl')

    else:

        file_extension = os.path.splitext(file_path)[1].lower()

        if file_extension == '.xlsx':
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        elif file_extension == '.xls':
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')

        df.to_excel(target_path, index=False)

    # 数据清理
    if cleaning_mode == 'remove_duplicates':
        df = df.drop_duplicates(subset=[column_name])               # 删除指定列的重复行

    elif cleaning_mode == 'fill_missing_zero':
        df[column_name] = df[column_name].fillna(0)                 # 填充缺失值为 0

    elif cleaning_mode == 'fill_missing_blank':
        df[column_name] = df[column_name].fillna('<空白>')           # 填充缺失值为 '<空白>'

    elif cleaning_mode == 'fill_missing_previous':
        df[column_name] = df[column_name].ffill()                   # 充缺失值为 重复上一行

    elif cleaning_mode == 'standardize_text':
        df[column_name] = df[column_name].str.strip().str.lower()   # 标准化文本（去除首尾空格及转换为小写英文字母）

    elif cleaning_mode == 'convert_data_str':
        df[column_name] = df[column_name].astype(str)               # 将数据类型转换为字符型

    elif cleaning_mode == 'convert_data_int':
        df[column_name] = df[column_name].astype(str).str.replace(',', '')
        df[column_name] = pd.to_numeric(df[column_name], errors='coerce').fillna(0).astype(int) # 转换为整数，支持缺失值

    elif cleaning_mode == 'convert_data_float':
        df[column_name] = df[column_name].astype(str).str.replace(',', '')
        df[column_name] = pd.to_numeric(df[column_name], errors='coerce').astype(float)         # 转换为浮点数，支持缺失值

    elif cleaning_mode == 'convert_data_date':
        df[column_name] = pd.to_datetime(df[column_name], errors='coerce')                      # 将数据类型转换为时间日期类型

    elif cleaning_mode == 'drop_columns':
        df = df.drop(columns=[column_name])                         # 删除指定列

    backup_path = os.path.join(folder_path, 'data_cleaning_1.xlsx')

    counter = 2
    while os.path.exists(backup_path):
        new_filename = f'data_cleaning_{counter}.xlsx'
        backup_path = os.path.join(folder_path, new_filename)
        counter += 1

    shutil.copy(target_path, backup_path)

    df.to_excel(target_path, index=False)

    result_text = f'\n数据清洗成功，清洗前数据备份文件路径：{backup_path}\n'

    return ['data_cleaning_clean', [result_text]]


def data_cleaning_export(request):

    file_path = request.get("data", {}).get("file_path", "")
    sheet_name = request.get("data", {}).get("sheet_name", "")

    folder_path = os.path.dirname(file_path)
    course_path = os.path.join(folder_path, 'data_cleaning.xlsx')
    target_path = os.path.join(folder_path, 'sqlite_database.db')

    if os.path.exists(course_path):

        df = pd.read_excel(course_path, engine='openpyxl')

    else:

        file_extension = os.path.splitext(file_path)[1].lower()

        if file_extension == '.xlsx':
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        elif file_extension == '.xls':
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')
    
    # 將 DataFrame 輸出到 SQLite 資料庫
    try:

        conn = sqlite3.connect(target_path)                             # 建立或開啟資料庫檔案
        table_name = sheet_name                                         # 你可以自訂資料表名稱
        df.to_sql(table_name, conn, if_exists='replace', index=False)   # 寫入資料表
        conn.close()
        result_text = f'\n导出SQLite数据库成功，数据库文件路径：{target_path}\n'

    except Exception as e:

        result_text = f'\n导出失败，错误信息：{str(e)}\n'

    return ['data_cleaning_export', [result_text]]