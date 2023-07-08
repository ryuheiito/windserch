import os
import pandas as pd

# フォルダパス
folder_path = "output"

# Excelファイルのパス
excel_path = os.path.join(folder_path, "抽出結果_修正.xlsx")

# Excelファイルを読み込む
df = pd.read_excel(excel_path)

# 列の値を取得
station_names = df["station_name"].tolist()

# station_nameと同じ名前のcsvファイルを削除
for filename in os.listdir(folder_path):
    if filename.endswith(".csv"):
        name_without_ext = os.path.splitext(filename)[0]
        if name_without_ext in station_names:
            file_path = os.path.join(folder_path, filename)
            os.remove(file_path)

# "_conv"が含まれるcsvファイルの名前を変更
for filename in os.listdir(folder_path):
    if filename.endswith(".csv") and "_conv" in filename:
        new_filename = filename.replace("_conv", "")
        new_file_path = os.path.join(folder_path, new_filename)
        os.rename(os.path.join(folder_path, filename), new_file_path)
