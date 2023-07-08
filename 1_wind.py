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
