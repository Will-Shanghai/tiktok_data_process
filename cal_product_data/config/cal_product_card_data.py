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

print(f"当前工作目录: {os.getcwd()}")
print(f"预期输出目录: {os.path.abspath(output_base_dir)}")

for file_name in file_list:
    file_path = os.path.join(current_directory, file_name)
    print(f"\n>>>> 正在处理: {file_name}")

    try:
        # A. 提取日期范围
        df_date = pd.read_excel(file_path, header=None, nrows=1)
        first_line = str(df_date.iloc[0, 0])
        dates = re.findall(r'\d{4}[-/]\d{2}[-/]\d{2}', first_line)
        standard_date_range = f"{dates[0]} ~ {dates[1]}" if len(dates) >= 2 else "Unknown"

        # B. 读取主数据
        df = pd.read_excel(file_path, header=2)
        df.columns = df.columns.astype(str).str.strip()

        # C. 锁定列名
        def find_col(keywords):
            for k in keywords:
                for c in df.columns:
                    if k == c or k in c: return c
            return None

        col_item     = find_col(['商品', '商品名称', 'Product Name'])
        col_vv       = find_col(['商品卡曝光次数'])
        col_browse   = find_col(['商品卡的去重页面浏览次数'])
        col_order    = find_col(['商品卡片商品成交件数'])
        col_customer = find_col(['商品卡去重客户数'])
        col_gmv      = find_col(['商品卡 GMV'])
        col_v2c      = find_col(['商品卡点击率'])
        col_v2o      = find_col(['商品卡转化率'])

        if not col_item:
            continue

        # D. 数据清洗函数
        def clean_num(x):
            if pd.isna(x): return 0
            val = str(x).replace('円', '').replace(',', '').replace(' ', '').strip()
            return pd.to_numeric(val, errors='coerce') or 0

        # E. 数值清洗 (扩展到所有数值列)
        all_num_cols = [col_vv, col_browse, col_order, col_customer, col_gmv]
        for col in all_num_cols:
            if col: df[col] = df[col].apply(clean_num)

        # F. 百分比清洗 (增加 strip 确保洁净)
        pct_cols = [col_v2o, col_v2c]
        for col in pct_cols:
            if col:
                df[col] = df[col].astype(str).str.replace('%', '', regex=False).str.strip()
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) / 100

        # G. 映射分类
        df['产品分类'] = df[col_item].apply(map_product)

        # H. 汇总逻辑
        for category in TARGET_CATEGORIES:
            group = df[df['产品分类'] == category].copy()
            if group.empty: continue

            print(f"   找到分类 [{category}]: {len(group)} 行数据")

            # 计算平均率值 (保持你要求的逻辑)
            avg_v2c = group[col_v2c].mean() if col_v2c else 0
            avg_v2o = group[col_v2o].mean() if col_v2o else 0

            final_summary_data.append({
                '源文件': file_name,
                '日期范围': standard_date_range,
                '产品分类': category,
                '曝光次数(AG)': int(group[col_vv].sum()),
                '商品卡的去重页面浏览次数(AH)': int(group[col_browse].sum()) if col_browse else 0,
                '商品卡片商品成交件数(AF)': int(group[col_order].sum()),
                '商品卡去重客户数(AJ)': int(group[col_customer].sum()) if col_customer else 0,
                'GMV(AE)': round(group[col_gmv].sum(), 2),
                '商品卡点击率(AK)': f"{avg_v2c * 100:.2f}%",
                '商品卡转化率(AL)': f"{avg_v2o * 100:.2f}%",
            })

    except Exception as e:
        print(f"❌ 出错: {e}")

# --- 4. 导出 ---
if final_summary_data:
    final_df = pd.DataFrame(final_summary_data)
    timestamp = datetime.now().strftime("%m%d%H%M")
    output_name = f"产品汇总_商品卡维度_{timestamp}.xlsx"
    save_path = os.path.join(output_base_dir, output_name)
    final_df.to_excel(save_path, index=False)
    print(f"\n🎉 报表生成成功！")
    print(f"📍 文件保存路径: {os.path.abspath(save_path)}")
else:
    print("\n❌ 警告：未匹配到任何指定产品的数据。")