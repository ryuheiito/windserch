import os
import pandas as pd

# 1. 「_conv」が含まれるcsvファイルをoutputフォルダから検索して読み込む
output_folder = "output"
for file_name in os.listdir(output_folder):
    if file_name.endswith('.csv') and "_conv" in file_name:
        csv_path = os.path.join(output_folder, file_name)
        df_csv = pd.read_csv(csv_path, encoding='SHIFT_JIS')

        # 2. ヘッダーを書き換える
        df_csv.columns = ['年', '平均風速', '最大風速', '最大風速(風向)', '最大瞬間風速', '最大瞬間風速(風向)']

        # 3. CSVを上書きして保存する
        df_csv.to_csv(csv_path, index=False, encoding='SHIFT_JIS')

        print(f"Processed file: {file_name}")
