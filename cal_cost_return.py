import pandas as pd
import re

# 文件路径
file_path = './data/process_data.csv'

# 1️⃣ 读取 CSV 文件
df = pd.read_csv(file_path)

# 2️⃣ 打印前几行数据，确认数据加载正确
print("前10行数据预览：")
print(df.head(10))

# 3️⃣ 指定需要处理的关键列
status_col = 'Order Status'               # B列：订单状态
n_col = 'SKU Platform Discount'           # N列：SKU平台折扣
p_col = 'SKU Subtotal After Discount'    # P列：折扣后的价格
sku_col = 'Seller SKU'                   # G列：卖家SKU（产品）

# 4️⃣ 过滤掉状态为“已取消”的行，只保留“已处理”或其他状态的数据
df_filtered = df[df[status_col] != '已取消'].copy()  # 使用 .copy() 防止警告

# 5️⃣ 创建 SKU 映射字典，将 SKU 代码映射为具体的产品名称
sku_mapping = {
    'FCBL001': '身体乳单瓶',
    '2FCBL001': '身体乳双瓶',
    '3FCBL001': '身体乳三瓶',
    'FCSS001': '防晒霜单瓶',
    '2FCSS001': '防晒霜双瓶',
    '3FCSS001': '防晒霜三瓶',
    'FCB&S001': '身体乳和防晒霜组合套装',
    'FCBLGREENTEE002': '身体乳单瓶'
}

# 6️⃣ 定义函数：根据 SKU 代码映射至产品名称
def map_sku_to_product(sku):
    return sku_mapping.get(sku, sku)  # 如果映射字典中有对应项，则返回映射后的名称；否则，返回原 SKU

# 7️⃣ 定义函数：从 SKU 中提取数量
def extract_quantity(sku):
    # 使用正则表达式从 SKU 中提取数字部分（数量）
    match = re.match(r'(\d+)', sku)
    return int(match.group(1)) if match else 1  # 如果找到数量，返回数字部分，否则默认为 1

# 8️⃣ 对 G列 (Seller SKU) 映射为具体产品名称
df_filtered['Product Name'] = df_filtered[sku_col].apply(map_sku_to_product)

# 9️⃣ 提取每个 SKU 对应的数量，并在新的列中保存
df_filtered['Product Quantity'] = df_filtered[sku_col].apply(extract_quantity)

# 🔟 对 G列映射后的产品名称进行分组，计算每个 SKU 对应的产品数量
sku_summary = df_filtered.groupby('Product Name').agg({'Product Quantity': 'sum'}).reset_index()

# 1️⃣1️⃣ 自定义排序的产品名称列表（按照你希望的顺序）
order = ['身体乳单瓶', '身体乳双瓶', '身体乳三瓶', '防晒霜单瓶', '防晒霜双瓶', '防晒霜三瓶', '身体乳和防晒霜组合套装']

# 1️⃣2️⃣ 为 Product Name 添加一个排序字段
sku_summary['Sort Order'] = sku_summary['Product Name'].apply(lambda x: order.index(x) if x in order else len(order))

# 1️⃣3️⃣ 根据排序字段进行排序
sku_summary = sku_summary.sort_values(by='Sort Order').drop(columns='Sort Order').reset_index(drop=True)

# 1️⃣4️⃣ 输出每个 SKU 的具体产品数量汇总，并按照数量顺序输出
print("\nSeller SKU 分类汇总具体产品数量：")
print(sku_summary)

# 1️⃣5️⃣ 创建产品成本映射字典
cost_mapping = {
    '身体乳单瓶': 9.36,
    '身体乳双瓶': 18.72,
    '身体乳三瓶': 28.08,
    '防晒霜单瓶': 8.37,
    '防晒霜双瓶': 16.74,
    '防晒霜三瓶': 25.11
}

# 1️⃣6️⃣ 添加产品成本列，计算每个分类的总成本
sku_summary['Product Cost'] = sku_summary['Product Name'].apply(lambda x: cost_mapping.get(x, 0))  # 获取对应的成本，如果没有映射则为 0
sku_summary['Total Product Cost'] = sku_summary['Product Quantity'] * sku_summary['Product Cost']  # 计算每个分类的总成本

# 1️⃣7️⃣ 创建物流成本映射字典
logistics_cost_mapping = {
    '身体乳单瓶': 6.22,
    '身体乳双瓶': 9.53,
    '身体乳三瓶': 13.1,
    '防晒霜单瓶': 3.15,
    '防晒霜双瓶': 3.9,
    '防晒霜三瓶': 4.65
}

# 1️⃣8️⃣ 添加物流成本列，计算每个分类的物流成本
sku_summary['Logistics Cost'] = sku_summary['Product Name'].apply(lambda x: logistics_cost_mapping.get(x, 0))  # 获取对应的物流成本

# 1️⃣9️⃣ 计算每个分类的总成本（包括产品成本 + 物流成本）
sku_summary['Total Cost Including Logistics'] = sku_summary['Total Product Cost'] + (sku_summary['Product Quantity'] * sku_summary['Logistics Cost'])

# 2️⃣0️⃣ 输出每个分类的总成本（包括产品成本和物流成本）
print("\n每个分类的总成本（包括产品成本和物流成本）：")
print(sku_summary[['Product Name', 'Total Cost Including Logistics', 'Total Product Cost', 'Logistics Cost']])

# 2️⃣1️⃣ 输出每个分类的物流成本
print("\n每个分类的物流成本：")
print(sku_summary[['Product Name', 'Logistics Cost']])

# 2️⃣2️⃣ 拆分身体乳和防晒霜组合套装
# 获取组合套装的数量
combo_count = sku_summary.loc[sku_summary['Product Name'] == '身体乳和防晒霜组合套装', 'Product Quantity'].sum()

# 增加身体乳单瓶的数量
sku_summary.loc[sku_summary['Product Name'] == '身体乳单瓶', 'Product Quantity'] += combo_count
# 增加防晒霜单瓶的数量
sku_summary.loc[sku_summary['Product Name'] == '防晒霜单瓶', 'Product Quantity'] += combo_count

# 移除身体乳和防晒霜组合套装行
sku_summary = sku_summary[sku_summary['Product Name'] != '身体乳和防晒霜组合套装']

# 2️⃣3️⃣ 输出拆分后的身体乳和防晒霜数量
print("\nSeller SKU 分类汇总（防晒组合拆分）具体产品数量：")
print(sku_summary)

# 2️⃣4️⃣ 输出拆分后的产品成本
print("\n拆分后的产品成本：")
sku_summary['Total Product Cost'] = sku_summary['Product Quantity'] * sku_summary['Product Cost']
print(sku_summary[['Product Name', 'Product Quantity', 'Product Cost', 'Total Product Cost']])

# 2️⃣5️⃣ 输出拆分后的物流成本
print("\n拆分后的物流成本：")
sku_summary['Total Logistics Cost'] = sku_summary['Product Quantity'] * sku_summary['Logistics Cost']
print(sku_summary[['Product Name', 'Product Quantity', 'Logistics Cost', 'Total Logistics Cost']])

# 2️⃣6️⃣ 汇总整体 N列和 P列的总和（即数量求和）
n_sum = df_filtered[n_col].sum()  # 计算 N列的总和
p_sum = df_filtered[p_col].sum()  # 计算 P列的总和
total_sum = n_sum + p_sum       # N列和 P列的总和

# 2️⃣7️⃣ 输出 N列和 P列的总和
print(f"\nN列折扣总和: {n_sum}")
print(f"P列折后价总和: {p_sum}")
print(f"N+P总和: {total_sum}")
