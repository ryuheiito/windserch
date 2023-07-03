import pandas as pd

xlsx_path = "output/抽出結果.xlsx"
output_path = "output/抽出結果_修正.xlsx"

# データ読み取り
df = pd.read_excel(xlsx_path, dtype={"station_id": str})

# 桁数の算出
df["station_id_length"] = df["station_id"].str.len()

# 修正内容を新しいExcelファイルに書き込み
df.to_excel(output_path, index=False)

print("修正内容を書き込みました。")

