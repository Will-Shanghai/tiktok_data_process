import pandas as pd
import os
import glob

# -----------------------
# 1. 配置区
# -----------------------
# 第一步：获取当前脚本 cal_cost_return_jp.py 的绝对物理路径
current_script_dir = os.path.dirname(os.path.abspath(__file__))

# 第二步：基于脚本位置定位其他文件夹
# os.path.join 会自动处理斜杠，os.path.normpath 会把 ../ 转化为真实的绝对路径
folder_path = os.path.normpath(os.path.join(current_script_dir, '../data/data_JP'))
output_filename = os.path.normpath(os.path.join(current_script_dir, '../result/Weekly_Performance_Report_JP.xlsx'))

file_list = glob.glob(os.path.join(folder_path, '*.csv'))

# 关键列名
status_col = 'Order Status'
quantity_col = 'Normal or Pre-order'
product_col = 'Product Name'
n_col = 'SKU Platform Discount'
p_col = 'SKU Subtotal After Discount'
r_col = 'Original Shipping Fee'

# 汇率配置 (JPY 除以 20 换算)
EXCHANGE_RATE = 20.0

# 功能一：Normal 订单成本配置
cost_mapping_1 = {'黑色睫毛夹': 8.86, '白色睫毛推': 8.9, '电动剃眉刀': 4.6, '车载手机支架': 4.6,
                  '发饰水银13/14点': 16.3, '电动鼻毛刀': 8.8, '头发护理精华': 10, '蓝牙MP3': 43, '车载充电器': 18.4}
logistics_cost_mapping_1 = {'黑色睫毛夹': 22.65, '白色睫毛推': 22.45, '电动剃眉刀': 12.5, '车载手机支架': 17.5,
                            '发饰水银13/14点': 22.2, '电动鼻毛刀': 25.09, '头发护理精华': 27.7, '蓝牙MP3': 24.02, '车载充电器': 23.47}

# 功能二：寄样订单成本配置 (产品+物流总和)
cost_mapping_2 = {
    '黑色睫毛夹': 31.51, '蜻蜓两个装': 26.18, '卷发棒隔热袋': 37.2,
    '屏显耳机': 62, '电动剃眉刀': 17.1, '车载手机支架': 22.1, '发饰水银13/14点': 38.5, '电动鼻毛刀': 33.89
}

# 核心聚合映射：细分规格 -> 产品大类
category_mapping = {
    '黑色睫毛夹': '黑色睫毛夹', '白色睫毛推': '白色睫毛推', '电动剃眉刀': '电动剃眉刀',
    '电动鼻毛刀': '电动鼻毛刀', '车载手机支架': '车载手机支架', '发饰水银13/14点': '水银13点',
    '蜻蜓两个装': '蜻蜓', '卷发棒隔热袋': '卷发棒隔热袋', '屏显耳机': '屏显耳机', '头发护理精华': '头发护理精华',
    '蓝牙MP3': '蓝牙MP3', '车载充电器': '车载充电器'
}


# -----------------------
# 2. 工具函数
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
    if '美容液の大容量版' in name: return '头发护理精华'
    if 'Bluetooth&ヘッドフォンMP 3' in name: return '蓝牙MP3'
    if '車充電器' in name: return '车载充电器'
    return name


def parse_amount(value):
    if pd.isna(value): return 0.0
    s = str(value).replace(' ', '').replace('JPY', '').replace(',', '')
    try:
        return float(s)
    except:
        return 0.0


# -----------------------
# 3. 主处理逻辑
# -----------------------
weekly_summary_list = []
category_detail_list = []
product_detail_data = {}

for file_path in file_list:
    file_name = os.path.basename(file_path)
    print(f"正在处理: {file_name}")

    try:
        df = pd.read_csv(file_path)
    except Exception as e:
        print(f"读取失败 {file_name}: {e}")
        continue

    # 日期处理
    order_date_str = "N/A"
    if 'Created Time' in df.columns and len(df) > 0:
        raw_time = df.iloc[0]['Created Time']
        try:
            order_date_str = pd.to_datetime(raw_time).strftime('%Y-%m-%d')
        except:
            order_date_str = str(raw_time).split(' ')[0]

    # --- 功能一：Normal 订单计算 ---
    # 增加对日文状态词的兼容 (キャンセル済み=已取消, 未払い=未支付)
    cancel_keywords = ['Canceled', 'Unpaid', 'キャンセル済み', '未払い']
    df_normal = df[~df[status_col].isin(cancel_keywords) & (df[quantity_col] == 'Normal')].copy()
    df_normal['Mapped Name'] = df_normal[product_col].map(map_sku)

    for col in [n_col, p_col, r_col]:
        df_normal[col] = df_normal[col].map(parse_amount)

    sku_counts = df_normal['Mapped Name'].value_counts().to_dict()
    product_detail_data[order_date_str] = sku_counts

    norm_agg = df_normal.groupby('Mapped Name').agg({n_col: 'sum', p_col: 'sum', r_col: 'sum'}).reset_index()

    # --- 功能二：寄样订单计算 ---
    df_sample = df[~df[status_col].isin(cancel_keywords) & (df[quantity_col].isna() | (df[quantity_col] == ''))].copy()
    df_sample['Mapped Name'] = df_sample[product_col].map(map_sku)
    sample_counts = df_sample['Mapped Name'].value_counts().to_dict()

    # --- 指标归集 ---
    all_names_in_file = set(list(sku_counts.keys()) + list(sample_counts.keys()))
    file_p_cost, file_l_cost, file_s_cost = 0, 0, 0

    for name in all_names_in_file:
        q_norm = sku_counts.get(name, 0)
        q_samp = sample_counts.get(name, 0)

        row_sales = norm_agg[norm_agg['Mapped Name'] == name]
        s_val = row_sales[n_col].sum() + row_sales[p_col].sum() + row_sales[r_col].sum()
        p_val = row_sales[p_col].sum()

        p_cost = q_norm * cost_mapping_1.get(name, 0)
        l_cost = q_norm * logistics_cost_mapping_1.get(name, 0)
        s_cost = q_samp * cost_mapping_2.get(name, 0)

        file_p_cost += p_cost;
        file_l_cost += l_cost;
        file_s_cost += s_cost

        category_detail_list.append({
            '日期': order_date_str, '大类': category_mapping.get(name, '其他'), '销售额': s_val,
            'P列折后价': p_val, '产品成本': p_cost, '物流成本': l_cost, '寄样支出': s_cost
        })

    sales_total = df_normal[n_col].sum() + df_normal[p_col].sum() + df_normal[r_col].sum()
    weekly_summary_list.append({
        '文件名称': file_name, '订单创建时间': order_date_str, '销售总额(N+P+R)': round(sales_total, 2),
        '汇率后金额': round(sales_total / EXCHANGE_RATE, 2), 'P列折后价总和': round(df_normal[p_col].sum(), 2),
        '售出产品成本总和': round(file_p_cost, 2), '物流成本总和': round(file_l_cost, 2),
        '寄样总支出(含物流)': round(file_s_cost, 2)
    })

# -----------------------
# 4. 安全导出 (macOS & Windows 强力校验)
# -----------------------
if weekly_summary_list:
    df_final_weekly = pd.DataFrame(weekly_summary_list)

    df_cat_base = pd.DataFrame(category_detail_list)
    df_cat_summary = df_cat_base.groupby(['日期', '大类']).sum().reset_index()
    df_cat_summary['汇率后金额'] = (df_cat_summary['销售额'] / EXCHANGE_RATE).round(2)
    df_cat_summary = df_cat_summary[
        ['日期', '大类', '销售额', '汇率后金额', 'P列折后价', '产品成本', '物流成本', '寄样支出']].sort_values(
        by=['日期', '大类'])

    # 3. Product Quantity Detail (销量透视 + 汇总行)
    df_detail = pd.DataFrame.from_dict(product_detail_data, orient='index').fillna(0).astype(int)
    df_detail = df_detail.T
    df_detail.index.name = '产品名称'
    # 【新增】汇总行逻辑
    df_detail.loc['汇总'] = df_detail.sum(axis=0)

    # 【核心】macOS 强效文件占用检查
    temp_excel_file = os.path.join(os.path.dirname(os.path.abspath(output_filename)),
                                   "~$" + os.path.basename(output_filename))

    file_accessible = False
    while not file_accessible:
        if os.path.exists(temp_excel_file):
            print(f"\n⚠️  警告: 检测到 Excel 隐藏临时文件，'{output_filename}' 正在运行。")
            input("👉 请先【关闭】Excel 软件，然后按【回车键】重试程序...")
        elif os.path.exists(output_filename):
            try:
                os.rename(output_filename, output_filename)
                file_accessible = True
            except:
                print(f"\n⚠️  无法写入！文件 '{output_filename}' 被锁定。")
                input("👉 请关闭相关 Excel 窗口后重试...")
        else:
            file_accessible = True

    # 写入 Excel
    try:
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            df_final_weekly.to_excel(writer, sheet_name='Weekly Summary', index=False)
            df_cat_summary.to_excel(writer, sheet_name='Category Aggregation', index=False)
            df_detail.to_excel(writer, sheet_name='Product Quantity Detail')

            # 样式微调
            workbook = writer.book
            ws_detail = writer.sheets['Product Quantity Detail']
            black_font = workbook.add_format({'font_color': 'black', 'bold': False})
            ws_detail.set_row(len(df_detail), None, black_font)

        print(f"\n✅ 处理完成！日本站报表已保存至: {output_filename}")
    except Exception as e:
        print(f"\n❌ 保存失败: {e}")