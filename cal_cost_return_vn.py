import pandas as pd
import glob
import os

# -----------------------
# 文件夹路径
# -----------------------
folder_path = './data_VN'

# 获取文件夹下所有 CSV 文件
file_list = glob.glob(os.path.join(folder_path, '*.csv'))

# 关键列
status_col = 'Order Status'
quantity_col = 'Normal or Pre-order'
sku_col = 'Seller SKU'
n_col = 'SKU Platform Discount'
p_col = 'SKU Subtotal After Discount'
r_col = 'Original Shipping Fee'

# SKU 映射
sku_mapping = {
    'FCBL001': '身体乳单瓶',
    '2FCBL001': '身体乳双瓶',
    '3FCBL001': '身体乳三瓶',
    'FCSS001': '防晒霜单瓶',
    '2FCSS001': '防晒霜双瓶',
    '3FCSS001': '防晒霜三瓶',
    'FCSERUM001': '精华液单瓶',
    '2FCSERUM001': '精华液双瓶',
    '3FCSERUM001': '精华液三瓶',
    'FCB&S001': '身体乳和防晒霜组合套装',
    'FCBLGREENTEE002': '身体乳单瓶',
    '2FCBLGREENTEE002': '身体乳双瓶',
    '3FCBLGREENTEE002': '身体乳三瓶'
}

# -----------------------
# 循环处理每个文件
# -----------------------
for file_path in file_list:
    print(f"\n==================== 处理文件: {file_path} ====================\n")

    # 读取 CSV
    df = pd.read_csv(file_path)

    # -------------------------------------------
    # 功能一：排除已取消，保留 Normal 订单
    # -------------------------------------------
    df_normal = df[(df[status_col] != '已取消') & (df[quantity_col] == 'Normal')].copy()

    # 清洗 N、P、R 列，确保它们是数字类型
    df_normal[n_col] = pd.to_numeric(df_normal[n_col], errors='coerce')
    df_normal[p_col] = pd.to_numeric(df_normal[p_col], errors='coerce')
    df_normal[r_col] = pd.to_numeric(df_normal[r_col], errors='coerce')

    df_normal['Product Name'] = df_normal[sku_col].map(lambda x: sku_mapping.get(x, x))

    # 统计拆分前数量
    sku_summary = df_normal.groupby('Product Name').size().reset_index(name='Product Quantity_PreSplit')
    sku_summary['Product Quantity_PostSplit'] = sku_summary['Product Quantity_PreSplit']

    # 组合套装拆分
    combo_name = '身体乳和防晒霜组合套装'
    if combo_name in sku_summary['Product Name'].values:
        combo_qty = sku_summary.loc[sku_summary['Product Name'] == combo_name, 'Product Quantity_PostSplit'].values[0]
        for single in ['身体乳单瓶', '防晒霜单瓶']:
            if single in sku_summary['Product Name'].values:
                sku_summary.loc[sku_summary['Product Name'] == single, 'Product Quantity_PostSplit'] += combo_qty
            else:
                sku_summary = pd.concat([sku_summary,
                                         pd.DataFrame({'Product Name': [single],
                                                       'Product Quantity_PreSplit': [0],
                                                       'Product Quantity_PostSplit': [combo_qty]})],
                                        ignore_index=True)
        sku_summary.loc[sku_summary['Product Name'] == combo_name, 'Product Quantity_PostSplit'] = 0

    # 产品成本
    cost_mapping = {
        '身体乳单瓶': 9.36,
        '身体乳双瓶': 18.72,
        '身体乳三瓶': 28.08,
        '防晒霜单瓶': 8.37,
        '防晒霜双瓶': 16.74,
        '防晒霜三瓶': 25.11,
        '身体乳和防晒霜组合套装': 30.0,
        '精华液单瓶': 9.02,
        '精华液双瓶': 18.04,
        '精华液三瓶': 27.06
    }

    # 物流成本
    logistics_cost_mapping = {
        '身体乳单瓶': 6.22,
        '身体乳双瓶': 9.53,
        '身体乳三瓶': 13.1,
        '防晒霜单瓶': 3.15,
        '防晒霜双瓶': 3.9,
        '防晒霜三瓶': 4.65,
        '身体乳和防晒霜组合套装': 5.0,  # 已拆分
        '精华液单瓶': 3.21,
        '精华液双瓶': 4.02,
        '精华液三瓶': 4.83
    }

    # 产品成本及总成本
    sku_summary['Product Cost'] = sku_summary['Product Name'].map(cost_mapping).fillna(0)
    sku_summary['Total Product Cost'] = (sku_summary['Product Quantity_PostSplit'] * sku_summary['Product Cost']).round(2)

    # 物流成本及总成本
    sku_summary['Shipping Cost'] = sku_summary['Product Name'].map(logistics_cost_mapping).fillna(0)
    sku_summary['Total Shipping Cost'] = (sku_summary['Product Quantity_PostSplit'] * sku_summary['Shipping Cost']).round(2)

    # 总成本
    sku_summary['Total Cost Including Logistics'] = (sku_summary['Total Product Cost'] + sku_summary['Total Shipping Cost']).round(2)

    # 固定顺序输出
    order = ['身体乳单瓶', '身体乳双瓶', '身体乳三瓶',
             '防晒霜单瓶', '防晒霜双瓶', '防晒霜三瓶',
             '精华液单瓶', '精华液双瓶', '精华液三瓶', '身体乳和防晒霜组合套装']

    sku_summary['Sort Order'] = sku_summary['Product Name'].apply(lambda x: order.index(x) if x in order else len(order))
    sku_summary = sku_summary.sort_values(by='Sort Order').drop(columns='Sort Order').reset_index(drop=True)

    # 输出拆分前 vs 拆分后数量对比
    print("\n拆分前 vs 拆分后数量对比表：")
    print(sku_summary[['Product Name', 'Product Quantity_PreSplit', 'Product Quantity_PostSplit']])

    # 输出拆分后的产品成本明细及汇总
    print("\n拆分后的产品成本明细及汇总：")
    print(sku_summary[['Product Name', 'Product Quantity_PostSplit', 'Product Cost', 'Total Product Cost']])
    print(f"产品成本总和: {sku_summary['Total Product Cost'].sum():.2f}")

    # 输出拆分后的物流成本明细及汇总
    print("\n拆分后的物流成本明细及汇总：")
    print(sku_summary[['Product Name', 'Product Quantity_PostSplit', 'Shipping Cost', 'Total Shipping Cost']])
    print(f"物流成本总和: {sku_summary['Total Shipping Cost'].sum():.2f}")

    # N 列折扣总和、P 列折后价总和及 N+P 总和
    n_sum = df_normal[n_col].sum()
    p_sum = df_normal[p_col].sum()
    print(f"\nN列折扣总和: {n_sum:.2f}")
    print(f"P列折后价总和: {p_sum:.2f}")
    print(f"N+P总和: {n_sum + p_sum:.2f}")

    # -------------------------------------------
    # 功能二：未取消且E列为空的订单，拆分组合套装并计算成本
    # -------------------------------------------

    # 1️⃣ 筛选未取消订单且E列为空（NaN 或空字符串）
    df_not_cancelled_empty = df[(df[status_col] != '已取消') &
                                (df[quantity_col].isna() | (df[quantity_col] == ''))].copy()

    # 2️⃣ 映射产品名称
    df_not_cancelled_empty['Product Name'] = df_not_cancelled_empty[sku_col].map(lambda x: sku_mapping.get(x, x))

    # 3️⃣ 按产品统计拆分前数量
    empty_summary = df_not_cancelled_empty.groupby('Product Name').size().reset_index(name='Product Quantity_PreSplit')
    empty_summary['Product Quantity_PostSplit'] = empty_summary['Product Quantity_PreSplit']

    # 4️⃣ 拆分组合套装
    combo_name = '身体乳和防晒霜组合套装'
    if combo_name in empty_summary['Product Name'].values:
        combo_qty = empty_summary.loc[empty_summary['Product Name'] == combo_name, 'Product Quantity_PostSplit'].values[0]
        for single in ['身体乳单瓶', '防晒霜单瓶']:
            if single in empty_summary['Product Name'].values:
                empty_summary.loc[empty_summary['Product Name'] == single, 'Product Quantity_PostSplit'] += combo_qty
            else:
                empty_summary = pd.concat([empty_summary,
                                           pd.DataFrame({'Product Name': [single],
                                                         'Product Quantity_PreSplit': [0],
                                                         'Product Quantity_PostSplit': [combo_qty]})],
                                          ignore_index=True)
        empty_summary.loc[empty_summary['Product Name'] == combo_name, 'Product Quantity_PostSplit'] = 0

    # 5️⃣ 寄样成本映射
    cost_mapping_function2 = {
        '身体乳单瓶': 23.42,
        '身体乳双瓶': 32.96,
        '身体乳三瓶': 43.02,
        '防晒霜单瓶': 19.96,
        '防晒霜双瓶': 28.53,
        '防晒霜三瓶': 37.10,
        '精华液单瓶': 20.21,
        '精华液双瓶': 29.23,
        '精华液三瓶': 38.25
    }

    empty_summary['Product Cost'] = empty_summary['Product Name'].map(cost_mapping_function2).fillna(0)
    empty_summary['Total Product Cost'] = (empty_summary['Product Quantity_PostSplit'] * empty_summary['Product Cost']).round(2)

    # 6️⃣ 固定顺序输出
    order = ['身体乳单瓶', '身体乳双瓶', '身体乳三瓶',
             '防晒霜单瓶', '精华液单瓶', '精华液双瓶', '精华液三瓶', '身体乳和防晒霜组合套装']

    empty_summary['Sort Order'] = empty_summary['Product Name'].apply(lambda x: order.index(x) if x in order else len(order))
    empty_summary = empty_summary.sort_values(by='Sort Order').drop(columns='Sort Order').reset_index(drop=True)

    # 7️⃣ 输出拆分前 vs 拆分后数量对比
    print("\n拆分前 vs 拆分后数量对比（功能二）：")
    print(empty_summary[['Product Name', 'Product Quantity_PreSplit', 'Product Quantity_PostSplit']])

    # 8️⃣ 输出产品成本明细及汇总
    print("\n产品成本明细及汇总（功能二）：")
    print(empty_summary[['Product Name', 'Product Quantity_PostSplit', 'Product Cost', 'Total Product Cost']])
    print(f"产品成本总和: {empty_summary['Total Product Cost'].sum():.2f}")

