import pandas as pd
import numpy as np
import os
import re
from datetime import datetime

# 1. 配置路径
current_directory = '../data/'
output_base_dir = '../result/'
if not os.path.exists(output_base_dir):
    os.makedirs(output_base_dir)

# --- 2. 核心映射逻辑 ---
TARGET_CATEGORIES = [
    '精华液', '黑鸦片身体乳', '500ml身体乳A链', '500ml身体乳B链', '防晒霜',
    '车载手机支架', '电动修眉刀', '男士修鼻毛刀', '三合一充电器', '切菜神器', '屏显耳机'
]

def map_product(product_name):
    p = str(product_name)
    if 'Serum' in p: return '精华液'
    if 'Black Opium' in p: return '黑鸦片身体乳'
    if 'Niacinamide' in p: return '500ml身体乳B链'
    if '500g' in p: return '500ml身体乳A链'
    if 'SPF 50 +' in p: return '防晒霜'
    if '車のサンバイザーにス' in p: return '车载手机支架'
    if '電働眉剃り器' in p: return '电动修眉刀'
    if '鼻毛刀' in p or 'フェイス・トリマー' in p: return '男士修鼻毛刀'
    if '充電を備えた' in p: return '三合一充电器'
    if 'Safe Slicer' in p: return '切菜神器'
    if 'ANCノイズキャンセル' in p: return '屏显耳机'
    return '其他'

# --- 3. 获取文件 ---
file_list = [f for f in os.listdir(current_directory) if f.endswith('.xlsx')]
if not file_list:
    print(f"⚠️ 未找到 .xlsx 文件。")
    exit()

final_summary_data = []

for file_name in file_list:
    file_path = os.path.join(current_directory, file_name)
    print(f"\n>>>> 正在处理: {file_name}")

    try:
        # A. 提取日期范围
        df_date = pd.read_excel(file_path, header=None, nrows=1)
        first_line = str(df_date.iloc[0, 0])
        dates = re.findall(r'\d{4}[-/]\d{2}[-/]\d{2}', first_line)
        standard_date_range = f"{dates[0]} ~ {dates[1]}" if len(dates) >= 2 else "Unknown"

        # B. 读取主数据 (header=2 锁定第3行为表头)
        df = pd.read_excel(file_path, header=2)
        df.columns = df.columns.astype(str).str.strip()

        # C. 严格对应图片 D, F, G, J, K, L, M, N, O, P 列的关键字
        def find_col(keywords):
            for k in keywords:
                for c in df.columns:
                    if k in c: return c
            return None

        col_item   = find_col(['商品名称', 'Product Name'])        # B列
        col_vv     = find_col(['曝光次数', 'Exposure'])            # D列
        col_click  = find_col(['点击次数', 'Clicks count'])         # F列
        col_order  = find_col(['SKU订单数', 'Orders'])             # G列
        col_cart_c = find_col(['加车次数', 'Add to cart count'])    # J列
        col_gmv    = find_col(['GMV'])                            # K列
        col_v2o    = find_col(['曝光到成交转化率'])                  # L列
        col_v2c    = find_col(['曝光到点击转化率'])                  # M列
        col_c2cart = find_col(['点击到加车转化率'])                  # N列
        col_c2o    = find_col(['点击到成交转化率'])                  # O列
        col_cart2o = find_col(['加车到成交转化率'])                  # P列

        if not col_item:
            continue

        # D. 数据清洗 (数值类)
        num_cols = [col_vv, col_click, col_order, col_cart_c, col_gmv]
        for col in num_cols:
            if col: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        # E. 数据清洗 (百分比类 L, M, N, O, P)
        pct_cols = [col_v2o, col_v2c, col_c2cart, col_c2o, col_cart2o]
        for col in pct_cols:
            if col:
                df[col] = df[col].astype(str).str.replace('%', '', regex=False)
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) / 100

        # F. 映射分类
        df['产品分类'] = df[col_item].apply(map_product)

        # G. 汇总
        for category in TARGET_CATEGORIES:
            group = df[df['产品分类'] == category].copy()
            if group.empty: continue

            final_summary_data.append({
                '源文件': file_name,
                '日期范围': standard_date_range,
                '产品分类': category,
                '曝光次数(D)': int(group[col_vv].sum()),
                '点击次数(F)': int(group[col_click].sum()),
                'SKU订单数(G)': int(group[col_order].sum()),
                '加车次数(J)': int(group[col_cart_c].sum()),
                'GMV(K)': round(group[col_gmv].sum(), 2),
                '曝光到成交转化率(L)': f"{group[col_v2o].mean() * 100:.2f}%",
                '曝光到点击转化率(M)': f"{group[col_v2c].mean() * 100:.2f}%",
                '点击到加车转化率(N)': f"{group[col_c2cart].mean() * 100:.2f}%",
                '点击到成交转化率(O)': f"{group[col_c2o].mean() * 100:.2f}%",
                '加车到成交转化率(P)': f"{group[col_cart2o].mean() * 100:.2f}%"
            })

    except Exception as e:
        print(f"❌ 出错: {e}")

# --- 5. 导出 ---
if final_summary_data:
    final_df = pd.DataFrame(final_summary_data)
    timestamp = datetime.now().strftime("%m%d%H%M")
    output_name = f"产品汇总_全维度_{timestamp}.xlsx"
    final_df.to_excel(os.path.join(output_base_dir, output_name), index=False)
    print(f"\n🎉 报表生成成功，已包含 D 到 P 列所有红框指标。")