import os
import pandas as pd

# 1. "抽出結果_修正.xlsx"から"station_name"と"station_id_length"を読み取る
xlsx_path = "output/抽出結果_修正.xlsx"
df = pd.read_excel(xlsx_path, usecols=["station_name", "station_id_length"])

# 2. station_nameの各行と一致する名前のCSVファイルを検索し、処理を行う
output_folder = "output"
station_names = df['station_name']

for file_name in os.listdir(output_folder):
    if file_name.endswith('.csv'):
        base_name = os.path.splitext(file_name)[0]
        if base_name in station_names.values:
            csv_path = os.path.join(output_folder, file_name)
            df_csv = pd.read_csv(csv_path, encoding='SHIFT_JIS')
            
            # 列の削除と空白列の削除
            station_id_length = df.loc[df['station_name'] == base_name, 'station_id_length'].values[0]
            if station_id_length == 5:
                df_csv = df_csv.iloc[:, [0] + list(range(14, 19))].dropna(axis=1, how='all')
            elif station_id_length == 4:
                df_csv = df_csv.iloc[:, [0] + list(range(12, 17))].dropna(axis=1, how='all')
            
            # 結果の保存
            output_file = os.path.join(output_folder, f"{base_name}_conv.csv")
            df_csv.to_csv(output_file, index=False, encoding='SHIFT_JIS')
            print(f"Processed file: {file_name} -> {output_file}")
