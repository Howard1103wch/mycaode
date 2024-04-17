import pandas as pd
import os

# 定義需要刪除的行數和列名，行為直的，行由我們自己定義，列為橫的，第一列為0、1、2
rows_to_drop = list(range(0, 147)) + list(range(249, 395)) #第三步:你要移除哪幾列，兩部分刪除用list，單一區域可用range就好。
columns_to_drop = ['A','D','E','C'] #第四步:你要移除哪幾行，根據下方我們自己定義的名稱。

# 定義CSV資料夾路徑和輸出的Excel檔名
csv_folder = 'D:\AAdatatrans' #你要讀取的資料夾
excel_file = 'D:\AAdatatrans\IDVG_output.xlsx' #你要輸出的資料夾與檔案名稱

# 取得所有CSV檔案路徑
csv_files = [os.path.join(csv_folder, f) for f in os.listdir(csv_folder) if f.endswith('.csv')]

# 建立一個空的DataFrame對象
df_all = pd.DataFrame()

# 逐一處理CSV文件
for csv_file in csv_files:

     file_name = os.path.splitext(os.path.basename(csv_file))[0]

     # 使用pandas讀取CSV文件
     df = pd.read_csv(csv_file, encoding='utf-8', skiprows=rows_to_drop)
     df.columns = ['A', 'B', 'C', 'D','E', file_name] #第一步:定義我們的行，取決於你保留資料那列最多有幾行。
     df[file_name] = pd.to_numeric(df[file_name], errors='coerce') #第二步:將你要的資料行轉為數值
     df['B'] = pd.to_numeric(df['B'], errors='coerce') #第二步:將你要的資料那行轉為數值
     df = df.drop(columns=columns_to_drop)

     # 將目前CSV檔案的資料加入總DataFrame物件中
     df_all = pd.concat([df_all, df], axis=1)

# 將總DataFrame寫入Excel檔案中
with pd.ExcelWriter(excel_file) as writer:
     df_all.to_excel(writer, index=False)