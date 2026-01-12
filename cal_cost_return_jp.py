import pandas as pd
import os
import glob

# -----------------------
# 文件夹路径
# -----------------------
folder_path = './data_JP'

# 获取文件夹下所有 CSV 文件
file_list = glob.glob(os.path.join(folder_path, '*.csv'))

# -----------------------
# 关键列
# -----------------------
status_col = 'Order Status'
quantity_col = 'Normal or Pre-order'
product_col = 'Product Name'  # B市场用 H列
n_col = 'SKU Platform Discount'
p_col = 'SKU Subtotal After Discount'
r_col = 'Original Shipping Fee'

# -----------------------
# SKU 中文映射函数
# -----------------------
# -----------------------
# SKU 中文映射函数 (修正版)
# -----------------------
def map_sku(name):
    if pd.isna(name):
        return name
    if '最新アップグレード版' in name:
        return '黑色睫毛夹'
    if '2025年最新版アップグレード 5D' in name:
        return '白色睫毛推'

    # 修正：使用 'or' 连接两个 'in name' 的检查
    elif ('正規品 虫除け リアル' in name) or ('Dragonfly' in name):
        return '蜻蜓两个装'

    elif '収納 ヘアアイロンポーチ 耐熱300度' in name:
        return '卷发棒隔热袋'
    elif '髪飾り 13 / 14点' in name:
        return '发饰水银13/14点'
    elif 'Bluetooth 5.4ヘッドフォン、LCD保護ケース、ANCノイズキャンセル、ハイファイ音質' in name:
        return '屏显耳机'
    elif '電働眉剃' in name:
        return '电动剃眉刀'
    elif '車のサンバイ' in name:
        return '车载手机支架'
    elif '電働1台多用耳毛刀眉毛刀' in name:
        return '电动鼻毛刀'
    return name

# -----------------------
# 数值清洗函数
# -----------------------
def parse_amount(value):
    if pd.isna(value):
        return 0.0
    s = str(value).replace(' ', '').replace('JPY', '')
    try:
        return float(s)
    except:
        return 0.0

# -----------------------
# 循环处理每个文件
# -----------------------
for file_path in file_list:
    print(f"\n==================== 处理文件: {file_path} ====================\n")

    # 读取 CSV
    df = pd.read_csv(file_path)

    # -----------------------
    # 读取第二行 Paid Time 列作为订单创建时间
    # -----------------------
    if 'Created Time' in df.columns and len(df) > 1:
        order_creation_time = df.at[1, 'Created Time']  # 第二行索引是1
        print(f"订单创建时间: {order_creation_time}")
    else:
        print("未找到 Created Time 列或文件行数不足")

    # -----------------------
    # SKU 映射
    # -----------------------
    df['Product Name Mapped'] = df[product_col].map(map_sku)

    # -----------------------
    # 功能一：Normal订单（已取消排除）
    # -----------------------
    df_normal = df[(df[status_col] != 'Canceled') & (df[status_col] != 'Unpaid') & (df[quantity_col] == 'Normal')].copy()

    print(f"过滤后的Normal订单行数: {len(df_normal)}")

    # 清洗 N 列、P 列 和 R 列
    df_normal[n_col] = df_normal[n_col].map(parse_amount)
    df_normal[p_col] = df_normal[p_col].map(parse_amount)
    df_normal[r_col] = df_normal[r_col].map(parse_amount)

    # -----------------------
    # 拆分前数量统计
    # -----------------------
    sku_summary_1 = df_normal.groupby('Product Name Mapped').size().reset_index(name='Product Quantity_PreSplit')
    sku_summary_1['Product Quantity_PostSplit'] = sku_summary_1['Product Quantity_PreSplit']

    # -----------------------
    # 产品成本及物流成本
    # -----------------------
    cost_mapping_1 = {'黑色睫毛夹': 8.86, '白色睫毛推': 8.9, '电动剃眉刀': 4.6, '车载手机支架': 4.6, '发饰水银13/14点': 16.3}
    logistics_cost_mapping_1 = {'黑色睫毛夹': 22.65, '白色睫毛推': 22.45, '电动剃眉刀': 12.5, '车载手机支架':17.5, '发饰水银13/14点': 22.2}

    sku_summary_1['Product Cost'] = sku_summary_1['Product Name Mapped'].map(cost_mapping_1).fillna(0)
    sku_summary_1['Total Product Cost'] = (
                sku_summary_1['Product Quantity_PostSplit'] * sku_summary_1['Product Cost']).round(2)
    sku_summary_1['Shipping Cost'] = sku_summary_1['Product Name Mapped'].map(logistics_cost_mapping_1).fillna(0)
    sku_summary_1['Total Shipping Cost'] = (
                sku_summary_1['Product Quantity_PostSplit'] * sku_summary_1['Shipping Cost']).round(2)
    sku_summary_1['Total Cost Including Logistics'] = (
                sku_summary_1['Total Product Cost'] + sku_summary_1['Total Shipping Cost']).round(2)

    print("\n功能一：Normal订单（已取消排除和未支付）")
    print("\n拆分前 vs 拆分后数量：")
    print(sku_summary_1[['Product Name Mapped', 'Product Quantity_PreSplit', 'Product Quantity_PostSplit']])
    print(sku_summary_1[['Product Name Mapped', 'Product Quantity_PostSplit', 'Product Cost', 'Total Product Cost']])
    print(f"产品成本总和: {sku_summary_1['Total Product Cost'].sum():.2f}")
    print(sku_summary_1[['Product Name Mapped', 'Product Quantity_PostSplit', 'Shipping Cost', 'Total Shipping Cost']])
    print(f"物流成本总和: {sku_summary_1['Total Shipping Cost'].sum():.2f}")

    n_sum_1 = df_normal[n_col].sum()
    p_sum_1 = df_normal[p_col].sum()
    r_sum_1 = df_normal[r_col].sum()
    print(f"\nN列折扣总和: {n_sum_1:.2f}, P列折后价总和: {p_sum_1:.2f}, R列产品运费总和: {r_sum_1:.2f}, N+P+R总和: {n_sum_1 + p_sum_1 + r_sum_1:.2f}")

    # -----------------------
    # 功能二：E列为空订单（非已取消）
    # -----------------------
    df_empty = df[(df[status_col] != 'Canceled') & (df[quantity_col].isna() | (df[quantity_col] == ''))].copy()
    df_empty['Product Name Mapped'] = df_empty[product_col].map(map_sku)

    sku_summary_2 = df_empty.groupby('Product Name Mapped').size().reset_index(name='Product Quantity_PreSplit')
    sku_summary_2['Product Quantity_PostSplit'] = sku_summary_2['Product Quantity_PreSplit']

    cost_mapping_2 = {
        '黑色睫毛夹': 31.51,
        '蜻蜓两个装': 26.18,
        '头饰': 37.2,
        '卷发棒隔热袋': 37.2,
        '屏显耳机': 62,
        '电动剃眉刀': 17.1,
        '电动鼻毛刀': 37,
        '车载手机支架': 22.1,
        '发饰水银13/14点': 38.5
    }

    sku_summary_2['Product Cost'] = sku_summary_2['Product Name Mapped'].map(cost_mapping_2).fillna(0)
    sku_summary_2['Total Product Cost'] = (
                sku_summary_2['Product Quantity_PostSplit'] * sku_summary_2['Product Cost']).round(2)

    print("\n功能二：E列为空订单（非已取消）")
    print("\n拆分前 vs 拆分后数量：")
    print(sku_summary_2[['Product Name Mapped', 'Product Quantity_PreSplit', 'Product Quantity_PostSplit']])
    print(sku_summary_2[['Product Name Mapped', 'Product Quantity_PostSplit', 'Product Cost', 'Total Product Cost']])
    print(f"产品成本总和: {sku_summary_2['Total Product Cost'].sum():.2f}")
