import pandas as pd
import numpy as np

# 1. 配置文件的路径。请确保路径指向您本地的 est.xlsx 文件
# 如果文件在当前目录下，可以写成 "est.xlsx"
file_path = "test.xlsx"

# --- 定义核心列名，增加代码可读性 ---
ITEM_COL = '商品'  # F列
VV_COL = '视频vv'  # G列
ORDER_COL = '视频成交订单数'  # P列 (用作加权权重)
CTR_COL = '点击率（视频）'  # U列
CCR_COL = '点击成交转化率（视频）'  # X列

# 2. & 3. 提取日期范围并读取数据
try:
    # --- 提取日期范围（Excel 文件第一个单元格 A1）---
    # 读取整个文件，不设表头 (header=None)，只读取第一个工作表 (sheet_name=0)
    df_temp = pd.read_excel(file_path, header=None, sheet_name=0)

    # 获取第一个单元格的内容 (行索引0，列索引0)
    first_line = str(df_temp.iloc[0, 0]).strip()

    # 清理日期字符串，去除可能存在的标签、引号或多余字符
    first_line = first_line.replace('[日期范围]:', '').strip().replace('"', '').replace(',', '')

    print("-------------------------------------------------")
    print(f"提取的日期范围: {first_line}")
    print("-------------------------------------------------")

    # --- 读取主数据 ---
    # 使用 skiprows=2 跳过日期和空行（即跳过原始文件的第 1、2 行）
    # 原始文件的第 3 行将自动作为表头
    df = pd.read_excel(file_path, skiprows=2)

    print("✅ Excel 数据文件读取成功！")

except FileNotFoundError:
    print(f"❌ 错误：文件 '{file_path}' 未找到。请检查文件名和路径。")
    exit()
except Exception as e:
    # 捕获其他读取错误，并提醒用户安装所需库
    print(f"❌ 错误：读取Excel文件时发生错误: {e}")
    print("💡 提示：请确保安装了 'openpyxl' 库：pip install openpyxl")
    exit()

# --- 数据清洗与预处理 ---

# 4. 关键改进：清理列名（去除首尾空格，防止 Excel 读取时的格式错误）
df.columns = df.columns.str.strip()

# 5. 清理百分比数据 (CTR 和 CCR)
cols_to_clean = [CTR_COL, CCR_COL]

for col in cols_to_clean:
    if col in df.columns:
        # 将数据转为字符串，替换掉 '%'
        df[col] = (
            df[col]
            .astype(str)
            .str.replace('%', '', regex=False)
            .replace('', '0')  # 将空字符串（可能由缺失值引起）视为 0
        )
        # 强制转换为数字（errors='coerce' 会将非数字转为 NaN），然后将 NaN 设为 0，并除以 100
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) / 100
    else:
        print(f"⚠️ 警告：数据中未找到列名 '{col}'。请检查表头是否正确。")

# 6. 确保涉及计算的列 (订单数和 VV) 是数值类型
numeric_cols = [ORDER_COL, VV_COL]
for col in numeric_cols:
    if col in df.columns:
        # 强制转为数字，并将非数字（如缺失值或乱码）设为 0
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    else:
        print(f"⚠️ 警告：数据中未找到列名 '{col}'。请检查表头是否正确。")


# 7. 定义映射函数 (F列: '商品')
def map_product(product_name):
    """根据商品名称映射到指定的产品分类"""
    if pd.isna(product_name):
        return '其他/未匹配'

    product_name_str = str(product_name)

    # 映射规则：包含特定关键词的字符串
    if 'Serum' in product_name_str:
        return '精华液'
    elif '500g' in product_name_str:
        return '500ml身体乳'
    elif 'SPF 50 +' in product_name_str:
        return '防晒霜'
    else:
        return '其他/未匹配'


# 应用映射函数，创建新的 '产品分类' 列
if ITEM_COL in df.columns:
    df['产品分类'] = df[ITEM_COL].apply(map_product)
else:
    print("❌ 致命错误：数据中未找到核心列 '商品' (F列)。无法进行映射和分组。")
    exit()

# --- 分组计算与统计 ---

# 8. 筛选出需要分析的三种产品
target_categories = ['精华液', '500ml身体乳', '防晒霜']
df_analysis = df[df['产品分类'].isin(target_categories)].copy()

# 9. 计算加权项：订单数 * 播放量
df_analysis['订单vv乘积'] = df_analysis[ORDER_COL] * df_analysis[VV_COL]

# 10. 按 '产品分类' 进行分组聚合计算
results_df = df_analysis.groupby('产品分类').agg(
    # 总视频数（行数）
    总视频数=(ITEM_COL, 'size'),

    # 加权平均播放量所需的基础累计值
    累计订单数=(ORDER_COL, 'sum'),
    累计订单vv乘积=('订单vv乘积', 'sum'),

    # 累计点击率 (U列的总和)
    累计点击率=(CTR_COL, 'sum'),

    # 平均点击率 (U列的平均值)
    平均点击率=(CTR_COL, 'mean'),

    # 累计点击成交率 (X列的总和)
    累计点击成交率=(CCR_COL, 'sum'),

    # 平均点击成交率 (X列的平均值)
    平均点击成交率=(CCR_COL, 'mean')
)


# 11. 计算加权平均播放量
def calculate_weighted_avg(row):
    """计算加权平均播放量 = (订单数*VV)的累计和 / 订单数的累计和"""
    cumulative_orders = row['累计订单数']
    cumulative_order_vv = row['累计订单vv乘积']

    if cumulative_orders > 0:
        # 当订单数大于 0 时，执行除法
        return cumulative_order_vv / cumulative_orders
    else:
        # 当订单数为 0 时，返回 NaN (代表该指标不适用/无法计算)
        return np.nan


results_df['加权平均播放量'] = results_df.apply(calculate_weighted_avg, axis=1)

# 12. 整理最终输出格式
final_results = results_df.drop(columns=['累计订单数', '累计订单vv乘积'])

# 格式化百分比和播放量
# 处理加权平均播放量：如果是 NaN 则显示 'N/A'，否则四舍五入取整
final_results['加权平均播放量'] = final_results['加权平均播放量'].apply(
    lambda x: int(round(x)) if pd.notna(x) else 'N/A'
)

# 格式化所有百分比指标，保留两位小数
final_results['累计点击率'] = final_results['累计点击率'].apply(lambda x: f"{x:.2%}")
final_results['平均点击率'] = final_results['平均点击率'].apply(lambda x: f"{x:.2%}")
final_results['累计点击成交率'] = final_results['累计点击成交率'].apply(lambda x: f"{x:.2%}")
final_results['平均点击成交率'] = final_results['平均点击成交率'].apply(lambda x: f"{x:.2%}")

# 13. 打印最终结果
print("\n--- 关键数据分析结果 ---")
print(final_results)