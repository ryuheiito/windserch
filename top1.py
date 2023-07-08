import os
import pandas as pd
import numpy as np

# フォルダパス
folder_path = "output"

# Excelファイルのパス
excel_path = os.path.join(folder_path, "抽出結果_修正.xlsx")

# Excelファイルを読み込む
df_excel = pd.read_excel(excel_path)

# 観測所ごとのデータを保持するリスト
data_list = []

# 観測所ごとに処理
for index, row in df_excel.iterrows():
    # Noとstation_nameを取得
    no = row["No"]
    station_name = row["station_name"]
    
    # 観測所のCSVファイルパス
    csv_path = os.path.join(folder_path, f"{station_name}.csv")
    
    # CSVファイルが存在するかチェック
    if os.path.isfile(csv_path):
        df_csv = pd.read_csv(csv_path, encoding='shift-jis')
        
        # 最大風速ランキングが1のデータを抽出
        df_filtered = df_csv[df_csv["最大風速ランキング"] == 1]
        
        # 最大風速ランキングの数値が重複した場合は年が大きいものを優先
        df_filtered = df_filtered.sort_values(by=["最大風速ランキング", "年"], ascending=[True, False])
        
        # 最大風速ランキングが1の最新のデータを取得
        if not df_filtered.empty:
            latest_data = df_filtered.iloc[0]
            
            # データをリストに追加
            data_list.append([no, station_name, latest_data["年"], latest_data["最大風速ランキング"], latest_data["最大風速"]])
    else:
        # CSVファイルが存在しない場合はNaN値を追加
        data_list.append([no, station_name, np.nan, np.nan, np.nan])

# 結果をDataFrameに変換
df_result = pd.DataFrame(data_list, columns=["No", "観測所", "年", "最大風速ランキング", "最大風速"])

# No列で昇順ソート
df_result = df_result.sort_values(by="No", ascending=True)

# 結果をExcelファイルとして出力
result_path = os.path.join(folder_path, "result.xlsx")
df_result.to_excel(result_path, index=False)
