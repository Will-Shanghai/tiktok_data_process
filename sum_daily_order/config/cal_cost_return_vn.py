import pandas as pd
import glob
import os

# -----------------------
# 1. 全局配置区
# -----------------------
current_script_dir = os.path.dirname(os.path.abspath(__file__))
folder_path = os.path.normpath(os.path.join(current_script_dir, '../data/data_VN'))
output_filename = os.path.normpath(os.path.join(current_script_dir, '../result/Weekly_Performance_Report_VN.xlsx'))
file_list = glob.glob(os.path.join(folder_path, '*.csv'))


# 关键列名配置
status_col = 'Order Status'
quantity_col = 'Normal or Pre-order'
sku_col = 'Seller SKU'
n_col = 'SKU Platform Discount'
p_col = 'SKU Subtotal After Discount'
r_col = 'Original Shipping Fee'

# SKU 到细分规格映射
sku_mapping = {
    'FCBL001': '身体乳单瓶', '2FCBL001': '身体乳双瓶', '3FCBL001': '身体乳三瓶',
    'FCSS001': '防晒霜单瓶', '2FCSS001': '防晒霜双瓶', '3FCSS001': '防晒霜三瓶',
    'FCSERUM001': '精华液单瓶', '2FCSERUM001': '精华液双瓶', '3FCSERUM001': '精华液三瓶',
    'FCB&S001': '身体乳和防晒霜组合套装',
    'FCBLGREENTEE002': '身体乳单瓶', '2FCBLGREENTEE002': '身体乳双瓶', '3FCBLGREENTEE002': '身体乳三瓶',
    'FCBLBO300': '身体乳300ml单瓶', '2FCBLBO300': '身体乳300ml双瓶', '3FCBLBO300': '身体乳300ml三瓶',
    'SERUM001&BO001': '精华液&黑鸦片'
}

category_mapping = {
    '身体乳单瓶': '身体乳', '身体乳双瓶': '身体乳', '身体乳三瓶': '身体乳',
    '精华液单瓶': '精华液', '精华液双瓶': '精华液', '精华液三瓶': '精华液', '精华液&黑鸦片': '精华液',
    '防晒霜单瓶': '防晒霜', '防晒霜双瓶': '防晒霜', '防晒霜三瓶': '防晒霜',
    '身体乳300ml单瓶': '身体乳300ml', '身体乳300ml双瓶': '身体乳300ml', '身体乳300ml三瓶': '身体乳300ml',
    '身体乳和防晒霜组合套装': '组合套装'
}

cost_mapping = {
    '身体乳单瓶': 9.36, '身体乳双瓶': 18.72, '身体乳三瓶': 28.08,
    '防晒霜单瓶': 8.37, '防晒霜双瓶': 16.74, '防晒霜三瓶': 25.11,
    '身体乳和防晒霜组合套装': 30.0, '精华液单瓶': 9.02, '精华液双瓶': 18.04,
    '精华液三瓶': 27.06, '身体乳300ml单瓶': 3.08, '身体乳300ml双瓶': 6.16,
    '身体乳300ml三瓶': 9.24, '精华液&黑鸦片': 12.1
}

logistics_cost_mapping = {
    '身体乳单瓶': 6.22, '身体乳双瓶': 9.53, '身体乳三瓶': 13.1,
    '防晒霜单瓶': 3.15, '防晒霜双瓶': 3.9, '防晒霜三瓶': 4.65,
    '身体乳和防晒霜组合套装': 5.0, '精华液单瓶': 3.21, '精华液双瓶': 4.02,
    '精华液三瓶': 4.83, '身体乳300ml单瓶': 2.03, '身体乳300ml双瓶': 4.06,
    '身体乳300ml三瓶': 6.09, '精华液&黑鸦片': 8.24
}

cost_mapping_function2 = {
    '身体乳单瓶': 23.42, '身体乳双瓶': 32.96, '身体乳三瓶': 43.02,
    '防晒霜单瓶': 19.96, '防晒霜双瓶': 28.53, '防晒霜三瓶': 37.10,
    '精华液单瓶': 20.21, '精华液双瓶': 29.23, '精华液三瓶': 38.25,
    '身体乳300ml单瓶': 14.12, '身体乳300ml双瓶': 23.6, '身体乳300ml三瓶': 30.3,
    '精华液&黑鸦片': 33.95
}

SORT_ORDER_LIST = ['身体乳300ml单瓶', '身体乳300ml双瓶', '身体乳300ml三瓶', '身体乳单瓶', '身体乳双瓶', '身体乳三瓶',
                   '防晒霜单瓶', '防晒霜双瓶', '防晒霜三瓶', '精华液单瓶', '精华液双瓶', '精华液三瓶',
                   '身体乳和防晒霜组合套装', '精华液&黑鸦片']

# 初始化
summary_list = []
category_detail_list = []
product_detail_data = {}

# -----------------------
# 2. 核心处理流程
# -----------------------
for file_path in file_list:
    file_name = os.path.basename(file_path)
    print(f"正在处理: {file_name}")

    try:
        df = pd.read_csv(file_path)
        for col in [n_col, p_col, r_col]:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        order_date_str = 'N/A'
        if 'Created Time' in df.columns and len(df) > 1:
            raw_time = df.iloc[1]['Created Time']
            try:
                order_date_str = pd.to_datetime(raw_time, dayfirst=True).strftime('%d/%m/%Y')
            except:
                order_date_str = str(raw_time).split(' ')[0]
        else:
            order_date_str = file_name

        # Normal 订单处理
        df_normal = df[~df[status_col].isin(['已取消', 'Canceled']) & (df[quantity_col] == 'Normal')].copy()
        df_normal['Product Name'] = df_normal[sku_col].map(lambda x: sku_mapping.get(x, x))

        sku_qty_pre = df_normal['Product Name'].value_counts().to_dict()
        sku_sales_map = df_normal.groupby('Product Name').apply(lambda x: (x[n_col] + x[p_col]).sum()).to_dict()
        sku_p_map = df_normal.groupby('Product Name')[p_col].sum().to_dict()

        combo_name = '身体乳和防晒霜组合套装'
        post_split_qty = sku_qty_pre.copy()
        if combo_name in post_split_qty:
            c_qty = post_split_qty[combo_name]
            for single in ['身体乳单瓶', '防晒霜单瓶']:
                post_split_qty[single] = post_split_qty.get(single, 0) + c_qty
            post_split_qty[combo_name] = 0

        # 寄样订单处理
        df_sample = df[
            ~df[status_col].isin(['已取消', 'Canceled']) & (df[quantity_col].isna() | (df[quantity_col] == ''))].copy()
        df_sample['Product Name'] = df_sample[sku_col].map(lambda x: sku_mapping.get(x, x))
        sample_qty_dict = df_sample['Product Name'].value_counts().to_dict()

        if combo_name in sample_qty_dict:
            s_c_qty = sample_qty_dict[combo_name]
            for single in ['身体乳单瓶', '防晒霜单瓶']:
                sample_qty_dict[single] = sample_qty_dict.get(single, 0) + s_c_qty
            sample_qty_dict[combo_name] = 0

        # 指标归集
        file_p_cost, file_l_cost, file_s_cost = 0, 0, 0
        all_names = set(list(post_split_qty.keys()) + list(sample_qty_dict.keys()) + list(sku_sales_map.keys()))

        for name in all_names:
            q_n = post_split_qty.get(name, 0);
            q_s = sample_qty_dict.get(name, 0)
            p_cost = q_n * cost_mapping.get(name, 0);
            l_cost = q_n * logistics_cost_mapping.get(name, 0);
            s_cost = q_s * cost_mapping_function2.get(name, 0)
            file_p_cost += p_cost;
            file_l_cost += l_cost;
            file_s_cost += s_cost
            category_detail_list.append({
                '日期': order_date_str, '大类': category_mapping.get(name, '其他'),
                '销售额': sku_sales_map.get(name, 0),
                'P列折后价': sku_p_map.get(name, 0), '产品成本': p_cost, '物流成本': l_cost, '寄样支出': s_cost
            })

        summary_list.append({
            '文件名称': file_name, '订单日期': order_date_str,
            '销售总额(N+P)': (df_normal[n_col].sum() + df_normal[p_col].sum()),
            '汇率后(VND/3600)': round((df_normal[n_col].sum() + df_normal[p_col].sum()) / 3600, 2),
            'P列折后价总和': df_normal[p_col].sum(), '售出产品成本总和': file_p_cost, '物流成本总和': file_l_cost,
            '寄样支出总和': file_s_cost
        })
        product_detail_data[order_date_str] = post_split_qty

    except Exception as e:
        print(f"处理失败 {file_name}: {e}")

# -----------------------
# 3. 报表生成与强力校验 (macOS & Windows)
# -----------------------
df_weekly = pd.DataFrame(summary_list)
df_cat_summary = pd.DataFrame(category_detail_list).groupby(['日期', '大类']).sum().reset_index()
df_cat_summary['汇率后(销售额/3600)'] = (df_cat_summary['销售额'] / 3600).round(2)
df_cat_summary = df_cat_summary.sort_values(by='日期', ascending=True)

df_pivot = pd.DataFrame.from_dict(product_detail_data, orient='index').fillna(0).astype(int)
df_pivot = df_pivot[[c for c in SORT_ORDER_LIST if c in df_pivot.columns]].T
df_pivot.index.name = '产品名称'
df_pivot.loc['汇总'] = df_pivot.sum(axis=0)

# 定义 Excel 产生的隐藏临时文件路径
# Mac 下 Excel 打开文件会生成一个 ~$ 开头的同名文件
temp_excel_file = os.path.join(os.path.dirname(os.path.abspath(output_filename)),
                               "~$" + os.path.basename(output_filename))

file_accessible = False
while not file_accessible:
    # 检查1: 是否存在隐藏临时文件 (Mac 专用)
    if os.path.exists(temp_excel_file):
        print(f"\n⚠️  警告: 检测到 Excel 临时文件，'{output_filename}' 正在运行中。")
        input("👉 请先【彻底关闭】Excel 软件，然后按【回车键】重试...")
    # 检查2: 如果原文件存在，尝试重命名 (Windows 专用)
    elif os.path.exists(output_filename):
        try:
            os.rename(output_filename, output_filename)
            file_accessible = True
        except:
            print(f"\n⚠️  无法写入！文件 '{output_filename}' 被占用。")
            input("👉 请关闭 Excel 窗口后，按【回车键】重试...")
    else:
        file_accessible = True

# 写入
with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
    df_weekly.to_excel(writer, sheet_name='Weekly Summary', index=False)
    df_cat_summary.to_excel(writer, sheet_name='Category Aggregation', index=False)
    df_pivot.to_excel(writer, sheet_name='Quantity Pivot')

    workbook = writer.book
    ws_pivot = writer.sheets['Quantity Pivot']
    black_font = workbook.add_format({'font_color': 'black', 'bold': False})
    ws_pivot.set_row(len(df_pivot), None, black_font)

print("-" * 30)
print(f"✅ 处理完成！数据已安全保存至：{output_filename}")