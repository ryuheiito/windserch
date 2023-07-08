import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# データフォーマットをSHIFT_JISに変換するための設定
encoding = "SHIFT_JIS"

# ファイルパスの設定
input_folder = "input"
output_folder = "output"
obs_stations_path = os.path.join(input_folder, "obs_stations.xlsx")
extracted_data_path = os.path.join(input_folder, "抽出観測点.xlsx")
output_path = os.path.join(output_folder, "抽出結果.xlsx")
output_csv_path = os.path.join(output_folder, "抽出結果.csv")
webdriver_path = "chromedriver"

# obs_stations.xlsxからstation_id, amedas_id1, station_name, fuken_idを読み取る
df_stations = pd.read_excel(obs_stations_path)
df_stations = df_stations[["station_id", "amedas_id1", "station_name", "fuken_id"]]

print("obs_stations.xlsxから読み取ったデータ:")
print(df_stations)

# 抽出観測点.xlsxからNo, 地点, 地点コードを読み取る
df_extracted = pd.read_excel(extracted_data_path)
df_extracted = df_extracted[["No", "地点", "地点コード"]]
print("抽出観測点.xlsxから読み取ったデータ:")
print(df_extracted)

# amedas_id1と地点コードが一致するデータを抽出
merged_data = pd.merge(df_extracted, df_stations, left_on="地点コード", right_on="amedas_id1", how="left")
output_data = merged_data[["No", "station_id", "station_name", "fuken_id"]]

# 一致するデータがない場合はNo以外の項目を9999とする
output_data.loc[output_data["station_id"].isnull(), ["station_id", "station_name", "fuken_id"]] = 9999

# 抽出結果をExcelファイルとして保存
output_excel_path = os.path.join(output_folder, "抽出結果.xlsx")
output_data.to_excel(output_excel_path, index=False)
print("抽出結果.xlsxを作成しました。保存先:", output_excel_path)

# Seleniumの設定
chrome_options = Options()
chrome_options.add_argument("--headless")  # ヘッドレスモードで実行する場合に指定（画面表示なし）
driver = webdriver.Chrome(webdriver_path, options=chrome_options)

# 3. スクレイピングとファイル保存
visited_stations = set()  # 訪れたstation_idを保持するセット
no_data_count = 0  # 対象データがない観測所の数
for index, row in output_data.iterrows():
    fuken_id = row["fuken_id"]
    station_id = row["station_id"]
    station_name = row["station_name"]

    if station_id in visited_stations:
        print(f"station_id {station_id} は既に処理済みです。")
        continue

    visited_stations.add(station_id)

    url = f"https://www.data.jma.go.jp/obd/stats/etrn/index.php?prec_no={fuken_id}&block_no={station_id}&year=&month=&day=&view="
    driver.get(url)

    try:
        # 画面遷移を待つ
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="main"]/table[4]/tbody/tr/td[3]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/a'))
        )

        # リンクをクリック
        link_element = driver.find_element(By.XPATH, '//*[@id="main"]/table[4]/tbody/tr/td[3]/div/table/tbody/tr[1]/td[1]/table/tbody/tr[1]/td[2]/a')
        link_element.click()

        # 画面遷移を待つ
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="tablefix1"]/tbody/tr'))
        )

        # テーブルの要素を取得
        rows = driver.find_elements(By.XPATH, '//*[@id="tablefix1"]/tbody/tr')

        # ヘッダーを取得
        header_row = rows[0]
        header_elements = header_row.find_elements(By.TAG_NAME, 'th')
        header = [element.text for element in header_elements]

        # データ行を取得
        data_rows = rows[1:]
        data = []
        for data_row in data_rows:
            data_elements = data_row.find_elements(By.TAG_NAME, 'td')
            row_data = [element.text for element in data_elements]
            data.append(row_data)

        # DataFrameを作成
        table_data = [header] + data
        df_table = pd.DataFrame(table_data)

        # 保存するファイル名を変更
        excel_filename = f"{station_name}.xlsx"
        excel_path = os.path.join(output_folder, excel_filename)

        # CSVファイルとして保存
        csv_filename = f"{station_name}.csv"
        csv_path = os.path.join(output_folder, csv_filename)
        df_table.to_csv(csv_path, index=False, header=False, encoding="SHIFT_JIS")
    except:
        no_data_count += 1
        print("対象データがない観測所があります")

# ブラウザを閉じる
driver.quit()

print("対象データがない観測所の数:", no_data_count)


xlsx_path = "output/抽出結果.xlsx"
output_path = "output/抽出結果_修正.xlsx"

# データ読み取り
df = pd.read_excel(xlsx_path, dtype={"station_id": str})

# 桁数の算出
df["station_id_length"] = df["station_id"].str.len()

# 修正内容を新しいExcelファイルに書き込み
df.to_excel(output_path, index=False)

print("修正内容を書き込みました。")



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


# 1. 「_conv」が含まれるcsvファイルをoutputフォルダから検索して読み込む
output_folder = "output"
for file_name in os.listdir(output_folder):
    if file_name.endswith('.csv') and "_conv" in file_name:
        csv_path = os.path.join(output_folder, file_name)
        df_csv = pd.read_csv(csv_path, encoding='SHIFT_JIS')

        # 2. 空白行を削除する
        df_csv = df_csv.dropna()

        # 3. CSVを上書きして保存する
        df_csv.to_csv(csv_path, index=False, encoding='SHIFT_JIS')

        print(f"Processed file: {file_name}")




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

