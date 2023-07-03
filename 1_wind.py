import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ファイルパスの設定
input_folder = "input"
output_folder = "output"
webdriver_path = "chromedriver"

# 1. obs_stations.xlsxからデータを読み取る
obs_stations_path = os.path.join(input_folder, "obs_stations.xlsx")
df_stations = pd.read_excel(obs_stations_path, encoding="SHIFT_JIS")

# 2. 抽出観測点.xlsxからデータを読み取り、一致するデータを抽出
extracted_data_path = os.path.join(input_folder, "抽出観測点.xlsx")
df_extracted = pd.read_excel(extracted_data_path, encoding="SHIFT_JIS")

# station_nameと一致するデータを抽出
merged_data = pd.merge(df_extracted, df_stations, left_on="地点", right_on="station_name", how="inner")
output_data = merged_data[["No", "station_id", "station_name", "fuken_id"]]

# 抽出結果を保存
output_path = os.path.join(output_folder, "抽出結果.xlsx")
output_data.to_excel(output_path, index=False)

# Seleniumの設定
chrome_options = Options()
chrome_options.add_argument("--headless")  # ヘッドレスモードで実行する場合に指定（画面表示なし）
driver = webdriver.Chrome(webdriver_path, options=chrome_options)

# 3. スクレイピングとファイル保存
for index, row in output_data.iterrows():
    fuken_id = row["fuken_id"]
    station_id = row["station_id"]
    station_name = row["station_name"]

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

        # CSVファイルとして保存
        csv_filename = f"{station_name}.csv"
        csv_path = os.path.join(output_folder, csv_filename)
        df_table.to_csv(csv_path, index=False, header=False, encoding="SHIFT_JIS")
    except:
        print("対象データがありません")

# ブラウザを閉じる
driver.quit()
