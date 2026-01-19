import pandas as pd
import os
import glob

# -----------------------
# 配置区
# -----------------------
folder_path = './data_JP'
output_filename = 'summary_report_JP.xlsx'
file_list = glob.glob(os.path.join(folder_path, '*.csv'))

# 关键列名
status_col = 'Order Status'
quantity_col = 'Normal or Pre-order'
product_col = 'Product Name'
n_col = 'SKU Platform Discount'
p_col = 'SKU Subtotal After Discount'
r_col = 'Original Shipping Fee'

# 功能一：Normal 订单成本配置
cost_mapping_1 = {'黑色睫毛夹': 8.86, '白色睫毛推': 8.9, '电动剃眉刀': 4.6, '车载手机支架': 4.6,
                  '发饰水银13/14点': 16.3}
logistics_cost_mapping_1 = {'黑色睫毛夹': 22.65, '白色睫毛推': 22.45, '电动剃眉刀': 12.5, '车载手机支架': 17.5,
                            '发饰水银13/14点': 22.2}

# 功能二：寄样订单成本配置 (产品+物流总和)
cost_mapping_2 = {
    '黑色睫毛夹': 31.51, '蜻蜓两个装': 26.18, '卷发棒隔热袋': 37.2,
    '屏显耳机': 62, '电动剃眉刀': 17.1, '电动鼻毛刀': 37,
    '车载手机支架': 22.1, '发饰水银13/14点': 38.5
}


# -----------------------
# 工具函数
# -----------------------
def map_sku(name):
    if pd.isna(name): return name
    if '最新アップグレード版' in name: return '黑色睫毛夹'
    if '2025年最新版アップグレード 5D' in name: return '白色睫毛推'
    if ('正規品 虫除け リアル' in name) or ('Dragonfly' in name): return '蜻蜓两个装'
    if '収納 ヘアアイロンポーチ 耐熱300度' in name: return '卷发棒隔热袋'
    if '髪飾り 13 / 14点' in name: return '发饰水银13/14点'
    if 'Bluetooth 5.4ヘッドフォン' in name: return '屏显耳机'
    if '電働眉剃' in name: return '电动剃眉刀'
    if '車のサンバイ' in name: return '车载手机支架'
    if '電働1台多用耳毛刀眉毛刀' in name: return '电动鼻毛刀'
    return name


def parse_amount(value):
    if pd.isna(value): return 0.0
    s = str(value).replace(' ', '').replace('JPY', '').replace(',', '')
    try:
        return float(s)
    except:
        return 0.0


# -----------------------
# 主处理逻辑
# -----------------------
summary_data = []
product_detail_data = {}

for file_path in file_list:
    file_name = os.path.basename(file_path)
    print(f"正在处理: {file_name}")

    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        print(f"读取失败 {file_name}: {e}")
        continue

    # 1. 日期格式化处理
    order_date_str = "N/A"
    if 'Created Time' in df.columns and len(df) > 0:
        raw_time = df.iloc[0]['Created Time']  # 取第一行即可
        try:
            # 尝试解析并标准化为 YYYY-MM-DD
            order_date_str = pd.to_datetime(raw_time).strftime('%Y-%m-%d')
        except:
            order_date_str = str(raw_time).split(' ')[0]

    # 映射产品名称
    df['Product Name Mapped'] = df[product_col].map(map_sku)

    # --- 功能一：Normal 订单计算 ---
    df_normal = df[
        (df[status_col] != 'Canceled') & (df[status_col] != 'Unpaid') & (df[quantity_col] == 'Normal')].copy()

    # 清洗金额
    for col in [n_col, p_col, r_col]:
        df_normal[col] = df_normal[col].map(parse_amount)

    # 统计数量 (Sheet 2 数据源)
    sku_counts = df_normal.groupby('Product Name Mapped').size().to_dict()
    product_detail_data[order_date_str] = sku_counts

    # 计算各项总和
    n_sum = df_normal[n_col].sum()
    p_sum = df_normal[p_col].sum()
    r_sum = df_normal[r_col].sum()
    sales_total = n_sum + p_sum + r_sum  # 日本站逻辑：N+P+R

    # 计算成本
    sku_summary_1 = df_normal.groupby('Product Name Mapped').size().reset_index(name='qty')
    sku_summary_1['p_cost'] = sku_summary_1['Product Name Mapped'].map(cost_mapping_1).fillna(0)
    sku_summary_1['l_cost'] = sku_summary_1['Product Name Mapped'].map(logistics_cost_mapping_1).fillna(0)

    total_p_cost = (sku_summary_1['qty'] * sku_summary_1['p_cost']).sum()
    total_l_cost = (sku_summary_1['qty'] * sku_summary_1['l_cost']).sum()

    # --- 功能二：寄样订单计算 ---
    df_empty = df[(df[status_col] != 'Canceled') & (df[quantity_col].isna() | (df[quantity_col] == ''))].copy()
    df_empty['Product Name Mapped'] = df_empty[product_col].map(map_sku)

    sku_summary_2 = df_empty.groupby('Product Name Mapped').size().reset_index(name='qty')
    sku_summary_2['sample_cost'] = sku_summary_2['Product Name Mapped'].map(cost_mapping_2).fillna(0)
    total_sample_cost = (sku_summary_2['qty'] * sku_summary_2['sample_cost']).sum()

    # 收集汇总行
    summary_data.append({
        '文件名称': file_name,
        '订单创建时间': order_date_str,
        '销售总额(N+P+R)': round(sales_total, 2),
        'P列折后价总和': round(p_sum, 2),
        '售出产品成本总和': round(total_p_cost, 2),
        '物流成本总和': round(total_l_cost, 2),
        '寄样总支出(含物流)': round(total_sample_cost, 2)
    })

# -----------------------
# 导出 Excel
# -----------------------
if summary_data:
    # Sheet 1: 汇总
    df_final_summary = pd.DataFrame(summary_data)

    # Sheet 2: 数量明细转置
    df_detail = pd.DataFrame.from_dict(product_detail_data, orient='index').fillna(0).astype(int)
    df_detail = df_detail.T
    df_detail.index.name = '产品名称'

    try:
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            df_final_summary.to_excel(writer, sheet_name='Summary Report', index=False)
            df_detail.to_excel(writer, sheet_name='Product Quantity Detail')
        print(f"\n✅ 处理完成！文件已保存至: {output_filename}")
    except Exception as e:
        print(f"\n❌ 保存失败: {e}")
else:
    print("未发现有效数据进行汇总。")