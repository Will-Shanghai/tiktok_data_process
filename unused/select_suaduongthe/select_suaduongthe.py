import pandas as pd
import re

# 1. 设置文件名和关键列名
# ***** 请将 'test.xlsx' 替换为您文件的实际名称（确保后缀也是 .xlsx） *****
file_name = "test.xlsx"
column_name = '商品名称'

# 2. 读取 XLSX 文件 (使用 read_excel)
try:
    # ***** 使用 pd.read_excel 读取 XLSX 文件 *****
    df = pd.read_excel(file_name)
    # 如果您的数据在 Excel 的第二个工作表，可能需要添加 sheet_name=1
    # df = pd.read_excel(file_name, sheet_name='Sheet1')

except FileNotFoundError:
    print(f"错误：文件 '{file_name}' 未找到。请确保文件与 Python 脚本在同一目录下。")
    exit()

# 保持后续的越南语筛选逻辑不变...
if column_name not in df.columns:
    # ... (省略其余代码，保持不变)
    print(f"错误：数据中未找到名为 '{column_name}' 的列。请检查您的原始文件。")
    print("您的数据中可用的列名（前5个）：", list(df.columns)[:5])
    exit()

# 3. 定义常见的越南语关键词 (直接匹配更可靠)
vietnamese_keywords_regex = '(?i)sữa dưỡng thể|kem dưỡng thể|dưỡng da toàn thân|kem cơ thể'

# 4. 执行筛选
print("\n--- 使用越南语关键词进行匹配和筛选 ---")
filtered_df = df[df[column_name].str.contains(vietnamese_keywords_regex, case=False, na=False, regex=True)]

# 5. 导出筛选后的数据到新的 CSV 文件
output_file_name = "筛选后的_身体乳_身体霜_产品.csv"
filtered_df.to_csv(output_file_name, index=False, encoding='utf-8')

print("\n--- 筛选完成 ---")
print(f"总共找到 {len(filtered_df)} 条符合条件的产品数据。")
print(f"筛选后的数据已成功保存到文件：'{output_file_name}'")

print("\n--- 筛选结果预览 (前5行) ---")
print(filtered_df.head())