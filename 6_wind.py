import os
import csv
import pandas as pd
import numpy as np

# outputフォルダから「_conv」を含むcsvファイルを検索
folder_path = 'output'
files = [f for f in os.listdir(folder_path) if f.endswith('.csv') and '_conv' in f]

# 各ファイルの最大風速を計算し、順位を付ける
for file in files:
    file_path = os.path.join(folder_path, file)
    with open(file_path, 'r', encoding='shift_jis') as csv_file:
        reader = csv.reader(csv_file)
        data = list(reader)
        
        # 平均風速の列のインデックスを取得
        header = data[0]
        wind_speed_index = header.index('最大風速')
        
        # 最大風速を数値化し、順位を計算する
        speeds = []
        for row in data[1:]:
            try:
                speed = float(row[wind_speed_index])
                speeds.append(speed)
            except ValueError:
                speeds.append(np.nan)  # 数値以外の場合はNaNにする

        # 順位を計算して新しいヘッダーを追加する
        ranks = pd.Series(speeds).rank(ascending=False, method='min')
        header.append('最大風速ランキング')

        # 順位をデータに追加する
        for i in range(len(data[1:])):
            rank = ranks[i]
            if not np.isnan(rank):
                data[i+1].append(int(rank))
            else:
                data[i+1].append('')  # NaNの場合は空白文字列を追加する

    # 修正されたデータをファイルに書き込む（SHIFT_JIS）
    with open(file_path, 'w', encoding='shift_jis', newline='') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerows(data)
