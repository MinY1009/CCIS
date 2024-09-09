# -*- coding: utf-8 -*-
"""
Created on Sat Feb 17 09:59:15 2024

@author: ga0664
"""

import os
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

##code存檔位置
code_path = r'C:\Users\ga0664\Desktop\中華徵信所'
##設定當前工作目錄
os.chdir(code_path)

##獲取當前年月日
now = datetime.now()
nowYM = now.strftime("%Y%m")
minus_threedays = now - timedelta(days=3)
db_nowYMD = minus_threedays.strftime("%Y%m%d") # 資料集年月
#nowYMD = 20240322
#MarAgo = (nowYMD - relativedelta(months=3)).strftime("%Y%m%d")

##讀取資料位置[日盛台駿]=1
df1_path = rf'C:\Users\ga0664\Desktop\中華徵信所\中華拒往查詢\{nowYM}\{db_nowYMD}_52356_拒往退票檢核表.xlsx'
# 讀取資料原檔，跳過前 4 列
df1 = pd.read_excel(df1_path, sheet_name='退票檢核表', skiprows=4)
df1 = df1[['姓名', '身份證', '公司統編', '公司名稱', '查詢日期', '查覆截止日', \
           '存款不足已清償張數', '存款不足已清償金額', '存款不足未清償張數', '存款不足未清償金額']]

#篩出法人的資料(篩選公司統編不包含Na的資料 )
company1 = df1[df1['公司統編'].notna()]
#篩出自然人的資料(留下公司統編=NA的資料)
person1 = df1[df1['公司統編'].isna()]

#篩出存款不足已(ID_NPT)、未(ID_NOT)清償張數不為0的資料為
ID_NPT1 = company1[company1['存款不足已清償張數'] != 0]
ID_NOT1 = company1[company1['存款不足未清償張數'] != 0]
combined1 = pd.concat([ID_NPT1, ID_NOT1]).drop_duplicates()


# 讀取total_df檔案-先解鎖
# 設定 slove_lock.py 的路徑
slove_lock_path = r'C:\Users\ga0664\Desktop\中華徵信所\slove_lock.py'
with open(slove_lock_path, 'r', encoding='utf-8') as file:
     exec(file.read())   # 執行 slove_lock.py
     
unlock_path = r'C:\Users\ga0664\Desktop\中華徵信所\unlock\total_df.xlsx'
total_df = pd.read_excel(unlock_path, sheet_name='sheet1')

#比對total_df檔比對前三個月已有發查的資料設為，有資料的設為3
#和月底案件餘額比對，為0的設為T


