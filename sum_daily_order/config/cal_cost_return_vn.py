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

# SKU 映射表 (保持不变)
sku_mapping = {
    'FCBL001': '身体乳单瓶', '2FCBL001': '身体乳双瓶', '3FCBL001': '身体乳三瓶',
    'FCSS001': '防晒霜单瓶', '2FCSS001': '防晒霜双瓶', '3FCSS001': '防晒霜三瓶',
    'FCSERUM001': '精华液单瓶', '2FCSERUM001': '精华液双瓶', '3FCSERUM001': '精华液三瓶',
    'FCB&S001': '身体乳和防晒霜组合套装',
    'FCBLGREENTEE002': '身体乳单瓶', '2FCBLGREENTEE002': '身体乳双瓶', '3FCBLGREENTEE002': '身体乳三瓶',
    'FCBLBO300': '身体乳300ml单瓶', '2FCBLBO300': '身体乳300ml双瓶', '3FCBLBO300': '身体乳300ml三瓶',
    'SERUM001&BO001': '精华液&黑鸦片', 'FCPerfume001': '观山香薰'
}

category_mapping = {
    '身体乳单瓶': '身体乳', '身体乳双瓶': '身体乳', '身体乳三瓶': '身体乳',
    '精华液单瓶': '精华液', '精华液双瓶': '精华液', '精华液三瓶': '精华液', '精华液&黑鸦片': '精华液',
    '防晒霜单瓶': '防晒霜', '防晒霜双瓶': '防晒霜', '防晒霜三瓶': '防晒霜',
    '身体乳300ml单瓶': '身体乳300ml', '身体乳300ml双瓶': '身体乳300ml', '身体乳300ml三瓶': '身体乳300ml',
    '身体乳和防晒霜组合套装': '组合套装', '观山香薰': '香薰',
    '新产品(待核实)': '未分类'
}

# 成本配置 (保持不变)
cost_mapping = {
    '身体乳单瓶': 9.36, '身体乳双瓶': 18.72, '身体乳三瓶': 28.08,
    '防晒霜单瓶': 8.37, '防晒霜双瓶': 16.74, '防晒霜三瓶': 25.11,
    '身体乳和防晒霜组合套装': 30.0, '精华液单瓶': 9.02, '精华液双瓶': 18.04,
    '精华液三瓶': 27.06, '身体乳300ml单瓶': 3.08, '身体乳300ml双瓶': 6.16,
    '身体乳300ml三瓶': 9.24, '精华液&黑鸦片': 12.1, '观山香薰': 14.4,
    '新产品(待核实)': 0.0
}

logistics_cost_mapping = {
    '身体乳单瓶': 6.22, '身体乳双瓶': 9.53, '身体乳三瓶': 13.1,
    '防晒霜单瓶': 3.15, '防晒霜双瓶': 3.9, '防晒霜三瓶': 4.65,
    '身体乳和防晒霜组合套装': 5.0, '精华液单瓶': 3.21, '精华液双瓶': 4.02,
    '精华液三瓶': 4.83, '身体乳300ml单瓶': 2.03, '身体乳300ml双瓶': 4.06,
    '身体乳300ml三瓶': 6.09, '精华液&黑鸦片': 8.24, '观山香薰': 8.22,
    '新产品(待核实)': 0.0
}

cost_mapping_function2 = {
    '身体乳单瓶': 23.42, '身体乳双瓶': 32.96, '身体乳三瓶': 43.02,
    '防晒霜单瓶': 19.96, '防晒霜双瓶': 28.53, '防晒霜三瓶': 37.10,
    '精华液单瓶': 20.21, '精华液双瓶': 29.23, '精华液三瓶': 38.25,
    '身体乳300ml单瓶': 14.12, '身体乳300ml双瓶': 23.6, '身体乳300ml三瓶': 30.3,
    '精华液&黑鸦片': 33.95, '观山香薰': 22.62,
    '新产品(待核实)': 0.0
}

SORT_ORDER_LIST = ['身体乳300ml单瓶', '身体乳300ml双瓶', '身体乳300ml三瓶', '身体乳单瓶', '身体乳双瓶', '身体乳三瓶',
                   '防晒霜单瓶', '防晒霜双瓶', '防晒霜三瓶', '精华液单瓶', '精华液双瓶', '精华液三瓶',
                   '身体乳和防晒霜组合套装', '精华液&黑鸦片', '观山香薰', '新产品(待核实)']

# 初始化数据容器
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
        df.columns = df.columns.str.strip()
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

        for col in [n_col, p_col, r_col]:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        order_date_str = 'N/A'
        if 'Created Time' in df.columns and len(df) > 0:
            raw_time = df.iloc[0]['Created Time']
            try:
                # 统一转为 DD/MM/YYYY 字符串，后续再通过 pandas 物理排序
                order_date_str = pd.to_datetime(raw_time, dayfirst=True).strftime('%d/%m/%Y')
            except:
                order_date_str = str(raw_time).split(' ')[0]

        df[sku_col] = df[sku_col].replace('', None).fillna('Missing_SKU')

        df_normal = df[
            ~df[status_col].isin(['已取消', 'Canceled', 'Cancelled']) & (df[quantity_col] == 'Normal')].copy()
        df_normal['Mapped Name'] = df_normal[sku_col].map(lambda x: sku_mapping.get(x, '新产品(待核实)'))

        sku_qty_pre = df_normal['Mapped Name'].value_counts().to_dict()
        sku_sales_map = (df_normal[n_col] + df_normal[p_col]).groupby(df_normal['Mapped Name']).sum().to_dict()
        sku_p_map = df_normal.groupby('Mapped Name')[p_col].sum().to_dict()

        # 组合装拆分
        combo_name = '身体乳和防晒霜组合套装'
        post_split_qty = sku_qty_pre.copy()
        if combo_name in post_split_qty:
            c_qty = post_split_qty[combo_name]
            for single in ['身体乳单瓶', '防晒霜单瓶']:
                post_split_qty[single] = post_split_qty.get(single, 0) + c_qty
            post_split_qty[combo_name] = 0

        df_sample = df[~df[status_col].isin(['已取消', 'Canceled', 'Cancelled']) &
                       (df[quantity_col].isna() | (df[quantity_col] == '') | (df[quantity_col] == 'NaN'))].copy()
        df_sample['Mapped Name'] = df_sample[sku_col].map(lambda x: sku_mapping.get(x, '新产品(待核实)'))
        sample_qty_dict = df_sample['Mapped Name'].value_counts().to_dict()

        file_p_cost, file_l_cost, file_s_cost = 0, 0, 0
        all_names = set(list(post_split_qty.keys()) + list(sample_qty_dict.keys()) + list(sku_sales_map.keys()))

        for name in all_names:
            q_n = post_split_qty.get(name, 0)
            q_s = sample_qty_dict.get(name, 0)
            p_cost = q_n * cost_mapping.get(name, 0)
            l_cost = q_n * logistics_cost_mapping.get(name, 0)
            s_cost = q_s * cost_mapping_function2.get(name, 0)
            file_p_cost += p_cost
            file_l_cost += l_cost
            file_s_cost += s_cost

            category_detail_list.append({
                '日期': order_date_str,
                '大类': category_mapping.get(name, '其他'),
                '销售额': sku_sales_map.get(name, 0),
                'P列折后价': sku_p_map.get(name, 0),
                '产品成本': p_cost,
                '物流成本': l_cost,
                '寄样支出': s_cost
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
        print(f"❌ 处理失败 {file_name}: {e}")

# -----------------------
# 3. 报表生成与校验 (新增排序逻辑)
# -----------------------
if not summary_list:
    print("未发现有效数据。")
else:
    # A. Summary 表排序
    df_summary = pd.DataFrame(summary_list)
    df_summary['temp_date'] = pd.to_datetime(df_summary['订单日期'], format='%d/%m/%Y', errors='coerce')
    df_summary = df_summary.sort_values('temp_date').drop(columns=['temp_date'])

    # B. Category Aggregation 表排序
    df_cat_raw = pd.DataFrame(category_detail_list).groupby(['日期', '大类']).sum().reset_index()
    df_cat_raw['temp_date'] = pd.to_datetime(df_cat_raw['日期'], dayfirst=True)
    df_cat_raw['汇率后(销售额/3600)'] = (df_cat_raw['销售额'] / 3600).round(2)
    new_column_order = ['日期', '大类', '销售额', '汇率后(销售额/3600)', 'P列折后价', '产品成本', '物流成本',
                        '寄样支出']
    df_cat_summary = df_cat_raw.sort_values(['temp_date', '大类']).drop(columns=['temp_date'])[new_column_order]

    # C. Quantity Pivot 表排序 (最关键，让日期列按先后排列)
    df_pivot_raw = pd.DataFrame.from_dict(product_detail_data, orient='index').fillna(0).astype(int)
    # 将索引转为时间类型并排序
    df_pivot_raw.index = pd.to_datetime(df_pivot_raw.index, dayfirst=True)
    df_pivot_raw = df_pivot_raw.sort_index()
    # 转回字符串显示格式
    df_pivot_raw.index = df_pivot_raw.index.strftime('%d/%m/%Y')

    # 转置并应用产品名称排序
    valid_cols = [c for c in SORT_ORDER_LIST if c in df_pivot_raw.columns]
    other_cols = [c for c in df_pivot_raw.columns if c not in SORT_ORDER_LIST]
    df_pivot = df_pivot_raw[valid_cols + other_cols].T
    df_pivot.index.name = '产品名称'
    df_pivot.loc['汇总'] = df_pivot.sum(axis=0)

    # D. 文件占用检查 (保持不变)
    temp_excel_file = os.path.join(os.path.dirname(os.path.abspath(output_filename)),
                                   "~$" + os.path.basename(output_filename))
    file_accessible = False
    while not file_accessible:
        if os.path.exists(temp_excel_file):
            print(f"\n⚠️ 警告: 检测到 Excel 临时文件，'{output_filename}' 正在运行中。")
            input("👉 请【彻底关闭】Excel 软件，然后按【回车键】重试...")
        elif os.path.exists(output_filename):
            try:
                os.rename(output_filename, output_filename)
                file_accessible = True
            except:
                print(f"\n⚠️ 无法写入！文件 '{output_filename}' 被占用。")
                input("👉 请关闭 Excel 窗口后，按【回车键】重试...")
        else:
            file_accessible = True

    # E. 最终写入
    with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
        df_summary.to_excel(writer, sheet_name='Weekly Summary', index=False)
        df_cat_summary.to_excel(writer, sheet_name='Category Aggregation', index=False)
        df_pivot.to_excel(writer, sheet_name='Quantity Pivot')

    print("-" * 30)
    print(f"✅ 处理完成！数据已保存至：{output_filename}")