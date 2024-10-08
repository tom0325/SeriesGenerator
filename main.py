import pandas as pd

# 讀取 Excel 文件
df = pd.read_excel("天數測試.xlsx")

# 獲取 `Series Code` 為 "SOR" 的行
test1 = df["Series Code"] == "SOR"

# 找到 `Line No.` 的最大值
max_line_no = df[test1]["Line No."].max()

# 根據 `Series Code` 為 "SOR" 和 `Line No.` 等於最大值過濾 DataFrame
result = df[test1 & (df["Line No."] == max_line_no)]
print(result)

series_code =result["Series Code"].values[0]

a=[]

for i in range(3):
    a.append(series_code)
print(a)


