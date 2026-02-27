import pandas as pd
import numpy as np
import os
import re
from datetime import datetime

# 1. 配置路径
current_directory = '../data/'
output_base_dir = '../result/'
print(f"✅ 准备处理数据目录: {current_directory}")

# --- 定义标准内部列名和映射关系 ---
STD_ITEM_COL = '商品'
STD_VV_COL = '视频VV'
STD_ORDER_COL = '视频成交订单数'
STD_CTR_COL = '点击率'
STD_CCR_COL = '点击成交转化率'

COLUMN_MAP = {
    STD_ITEM_COL: ['商品', 'Products'],
    STD_VV_COL: ['视频vv', 'VV'],
    STD_ORDER_COL: ['视频成交订单数', 'Orders'],
    STD_CTR_COL: ['点击率（视频）', 'Click-through rate (Video)'],
    STD_CCR_COL: ['点击成交转化率（视频）', 'Click-to-order rate (Video)'],
}

# --- 2. 全局产品分类定义 ---
TARGET_CATEGORIES = [
    '精华液', '黑鸦片身体乳', '500ml身体乳A链', '500ml身体乳B链', '防晒霜',
    '车载手机支架', '电动修眉刀', '男士修鼻毛刀', '三合一充电器', '切菜神器', '屏显耳机'
]

# 3. 获取文件列表
if not os.path.exists(current_directory):
    print(f"❌ 错误：目录 '{current_directory}' 不存在。")
    exit()

file_list = [f for f in os.listdir(current_directory) if f.endswith('.xlsx')]
if not file_list:
    print(f"⚠️ 未找到 .xlsx 文件。")
    exit()

final_summary_data = []

# 4. 循环处理
for file_name in file_list:
    file_path = os.path.join(current_directory, file_name)
    print(f"\n>>>> 正在处理: {file_name}")

    try:
        # 提取日期范围
        df_date = pd.read_excel(file_path, header=None, nrows=1)
        first_line = str(df_date.iloc[0, 0])
        dates = re.findall(r'\d{4}[-/]\d{2}[-/]\d{2}', first_line)
        standard_date_range = f"{dates[0]} ~ {dates[1]}" if len(dates) >= 2 else "Unknown_Date"

        # 读取主数据 (跳过前两行)
        df = pd.read_excel(file_path, skiprows=2)
        df.columns = df.columns.astype(str).str.strip()
    except Exception as e:
        print(f"❌ 读取失败: {e}")
        continue

    # --- 5. 列名标准化 ---
    for std_col, possible_names in COLUMN_MAP.items():
        for name in possible_names:
            if name in df.columns:
                df.rename(columns={name: std_col}, inplace=True)
                break

    # 检查必要列
    if STD_ITEM_COL not in df.columns:
        print(f"⚠️ 文件缺少核心列，跳过")
        continue

    # --- 6. 数据清洗 ---
    for col in [STD_CTR_COL, STD_CCR_COL]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace('%', '').replace('nan', '0')
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) / 100

    for col in [STD_ORDER_COL, STD_VV_COL]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)


    # --- 7. 产品映射逻辑 ---
    def map_product(product_name):
        p = str(product_name)
        if 'Serum' in p: return '精华液'
        if 'Black Opium' in p: return '黑鸦片身体乳'
        if 'Niacinamide' in p: return '500ml身体乳B链'
        if '500g' in p: return '500ml身体乳A链'
        if 'SPF 50 +' in p: return '防晒霜'
        if '車のサンバイザーにス' in p: return '车载手机支架'
        if '電働眉剃り器' in p: return '电动修眉刀'
        if '鼻毛刀フェイス・トリマー男性' in p: return '男士修鼻毛刀'
        if '充電を備えた3-in-1の多機能' in p: return '三合一充电器'
        if '改良版 Safe Slicer Plus' in p: return '切菜神器'
        if 'ANCノイズキャンセル' in p: return '屏显耳机'
        return '其他'


    df['产品分类'] = df[STD_ITEM_COL].apply(map_product)

    # --- 8. 执行统计 ---
    present_categories = [c for c in TARGET_CATEGORIES if c in df['产品分类'].unique()]

    for category in present_categories:
        group = df[df['产品分类'] == category].copy()
        videos_with_orders = group[group[STD_ORDER_COL] > 0]

        count_with_orders = len(videos_with_orders)
        total_videos = len(group)
        total_orders_sum = group[STD_ORDER_COL].sum()
        total_vv_sum = group[STD_VV_COL].sum()

        # 计算比率数值
        video_order_rate = count_with_orders / total_videos if total_videos > 0 else 0
        avg_ctr = videos_with_orders[STD_CTR_COL].mean() if count_with_orders > 0 else 0
        avg_ccr = videos_with_orders[STD_CCR_COL].mean() if count_with_orders > 0 else 0
        avg_vv_ordered = videos_with_orders[STD_VV_COL].mean() if count_with_orders > 0 else 0

        # 构建字典并格式化百分比
        final_summary_data.append({
            '源文件': file_name,
            '日期范围': standard_date_range,
            '产品分类': category,
            '总视频数': total_videos,
            '出单视频数': count_with_orders,
            '总成交订单数': total_orders_sum,
            '总播放量VV': total_vv_sum,
            # 使用 f-string 进行百分比转换
            '视频出单率': f"{video_order_rate:.2%}",
            '出单视频-平均点击率': f"{avg_ctr:.2%}",
            '出单视频-平均转化率': f"{avg_ccr:.2%}",
            '出单视频-平均VV': round(avg_vv_ordered, 2),
        })

# --- 9. 导出结果 ---
if final_summary_data:
    final_df = pd.DataFrame(final_summary_data)

    if not os.path.exists(output_base_dir):
        os.makedirs(output_base_dir)

    timestamp = datetime.now().strftime("%m%d%H%M")

    # 按日期范围分组导出
    for date_range, group_df in final_df.groupby('日期范围'):
        # 处理日期字符串作为文件名时的非法字符
        safe_date = str(date_range).replace(' ~ ', '-').replace(' ', '_').replace('/', '-')
        output_name = f"数据汇总-{safe_date}-{timestamp}.xlsx"

        group_df.to_excel(os.path.join(output_base_dir, output_name), index=False)
        print(f"🎉 报表已生成: {output_name}")
else:
    print("❌ 未匹配到任何目标数据。")