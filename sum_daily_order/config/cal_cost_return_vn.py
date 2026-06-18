import pandas as pd
import glob
import os

# -----------------------
# 1. 静态映射配置
# -----------------------
sku_mapping = {
    'FCBL001': '身体乳单瓶', '2FCBL001': '身体乳双瓶', '3FCBL001': '身体乳三瓶',
    'FCSS001': '防晒霜单瓶', '2FCSS001': '防晒霜双瓶', '3FCSS001': '防晒霜三瓶',
    'FCSERUM001': '精华液单瓶', '2FCSERUM001': '精华液双瓶', '3FCSERUM001': '精华液三瓶',
    'FCB&S001': '身体乳和防晒霜组合套装', 'FCLS001': '唇部精华单支', '2FCLS001': '唇部精华双支',
    'FCBLGREENTEE002': '身体乳单瓶', '2FCBLGREENTEE002': '身体乳双瓶', '3FCBLGREENTEE002': '身体乳三瓶',
    'FCBLBO300': '身体乳300ml单瓶', '2FCBLBO300': '身体乳300ml双瓶', '3FCBLBO300': '身体乳300ml三瓶',
    'SERUM001&BO001': '精华液&黑鸦片', 'FCPerfume001': '观山香薰'
}

category_mapping = {
    '身体乳单瓶': '身体乳', '身体乳双瓶': '身体乳', '身体乳三瓶': '身体乳',
    '精华液单瓶': '精华液单瓶', '精华液双瓶': '精华液双瓶', '精华液三瓶': '精华液三瓶', '精华液&黑鸦片': '精华液&黑鸦片',
    '防晒霜单瓶': '防晒霜', '防晒霜双瓶': '防晒霜', '防晒霜三瓶': '防晒霜',
    '唇部精华单只': '唇部精华', '唇部精华双只': '唇部精华',
    '身体乳300ml单瓶': '身体乳300ml', '身体乳300ml双瓶': '身体乳300ml', '身体乳300ml三瓶': '身体乳300ml',
    '身体乳和防晒霜组合套装': '组合套装', '观山香薰': '香薰', '新产品(待核实)': '未分类'
}

SORT_ORDER_LIST = ['身体乳300ml单瓶', '身体乳300ml双瓶', '身体乳300ml三瓶', '身体乳单瓶', '身体乳双瓶', '身体乳三瓶',
                   '防晒霜单瓶', '防晒霜双瓶', '防晒霜三瓶', '精华液单瓶', '精华液双瓶', '精华液三瓶', '唇部精华单支', '唇部精华双支',
                   '身体乳和防晒霜组合套装', '精华液&黑鸦片', '观山香薰', '新产品(待核实)']

# -----------------------
# 2. 店铺差异化配置

# -----------------------
# 产品成本
common_cost = {
    '身体乳单瓶': 9.36, '身体乳双瓶': 18.72, '身体乳三瓶': 28.08, '防晒霜单瓶': 8.37,
    '精华液单瓶': 9.02, '身体乳300ml单瓶': 3.08, '精华液&黑鸦片': 12.1, '观山香薰': 14.4,
    '身体乳和防晒霜组合套装': 30.0, '唇部精华单支': 6.65, '唇部精华双支': 13.3
}

# 本土店头程+尾程
common_logistics = {
    '身体乳单瓶': 6.22, '身体乳双瓶': 9.53, '身体乳三瓶': 13.1, '防晒霜单瓶': 3.15,
    '精华液单瓶': 3.21, '身体乳300ml单瓶': 2.03, '精华液&黑鸦片': 8.24, '观山香薰': 8.22,
    '身体乳和防晒霜组合套装': 5.0, '唇部精华单支': 3.5, '唇部精华双支': 5.5
}
common_sample_cost = {
    '身体乳单瓶': 23.42, '身体乳双瓶': 32.96, '防晒霜单瓶': 19.96, '精华液单瓶': 20.21,
    '身体乳300ml单瓶': 14.12, '精华液&黑鸦片': 33.95, '观山香薰': 22.62, '唇部精华单支': 10.15, '唇部精华双支': 18.8
}

STORE_CONFIG = {
    'local': {'cn_name': '本土店', 'cost': common_cost, 'logistics': common_logistics, 'sample': common_sample_cost,
              'exchange_rate': 3883},
    'cross_border': {'cn_name': '跨境店', 'cost': common_cost, 'logistics': common_logistics,
                     'sample': common_sample_cost, 'exchange_rate': 3883}
}


# -----------------------
# 3. 核心处理函数
# -----------------------
def run_vietnam_report(store_key):
    conf = STORE_CONFIG[store_key]
    current_script_dir = os.path.dirname(os.path.abspath(__file__))
    folder_path = os.path.normpath(os.path.join(current_script_dir, f'../data/data_VN/{store_key}'))
    output_filename = os.path.normpath(
        os.path.join(current_script_dir, f'../result/Weekly_Performance_Report_VN_{store_key}.xlsx'))

    file_list = glob.glob(os.path.join(folder_path, '*.csv'))
    if not file_list:
        print(f"⚠️  [{conf['cn_name']}] 文件夹为空，跳过处理。")
        return

    print(f"🚀 正在处理 [{conf['cn_name']}] ({len(file_list)}个文件)...")

    summary_list = []
    category_detail_list = []
    product_detail_data = {}
    all_samples_collector = []

    status_col, quantity_col, sku_col = 'Order Status', 'Normal or Pre-order', 'Seller SKU'
    n_col, p_col, r_col = 'SKU Platform Discount', 'SKU Subtotal After Discount', 'Original Shipping Fee'

    for file_path in file_list:
        file_name = os.path.basename(file_path)
        try:
            df = pd.read_csv(file_path)
            df.columns = df.columns.str.strip()
            # --- 修正点：将 .map 改为 .applymap 以兼容旧版本 ---
            df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

            for col in [n_col, p_col, r_col]:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

            order_date_str = 'N/A'
            if 'Created Time' in df.columns and len(df) > 0:
                try:
                    order_date_str = pd.to_datetime(df.iloc[0]['Created Time'], dayfirst=True).strftime('%d/%m/%Y')
                except:
                    order_date_str = str(df.iloc[0]['Created Time']).split(' ')[0]

            df[sku_col] = df[sku_col].replace('', None).fillna('Missing_SKU')
            cancel_tags = ['已取消', 'Canceled', 'Cancelled']

            # --- A. 处理正常销售订单 ---
            df_normal = df[~df[status_col].isin(cancel_tags) & (df[quantity_col] == 'Normal')].copy()
            df_normal['Mapped Name'] = df_normal[sku_col].map(lambda x: sku_mapping.get(x, '新产品(待核实)'))

            sku_qty_pre = df_normal['Mapped Name'].value_counts().to_dict()
            sku_sales_map = (df_normal[n_col] + df_normal[p_col]).groupby(df_normal['Mapped Name']).sum().to_dict()
            sku_p_map = df_normal.groupby('Mapped Name')[p_col].sum().to_dict()

            combo_name = '身体乳和防晒霜组合套装'
            post_split_qty = sku_qty_pre.copy()
            if combo_name in post_split_qty:
                c_qty = post_split_qty[combo_name]
                for single in ['身体乳单瓶', '防晒霜单瓶']:
                    post_split_qty[single] = post_split_qty.get(single, 0) + c_qty
                post_split_qty[combo_name] = 0

            # --- B. 处理样品订单 ---
            df_sample = df[~df[status_col].isin(cancel_tags) & (df[quantity_col] != 'Normal')].copy()
            df_sample['Mapped Name'] = df_sample[sku_col].map(lambda x: sku_mapping.get(x, '新产品(待核实)'))

            if not df_sample.empty:
                sample_temp = df_sample[['Mapped Name']].copy()
                sample_temp['日期'] = order_date_str
                all_samples_collector.append(sample_temp)

            sample_qty_dict = df_sample['Mapped Name'].value_counts().to_dict()

            # --- C. 成本汇总逻辑 ---
            file_p_cost, file_l_cost, file_s_cost = 0, 0, 0
            all_names = set(list(post_split_qty.keys()) + list(sample_qty_dict.keys()))

            for name in all_names:
                q_n = post_split_qty.get(name, 0)
                q_s = sample_qty_dict.get(name, 0)
                p_cost = q_n * conf['cost'].get(name, 0)
                l_cost = q_n * conf['logistics'].get(name, 0)
                s_cost = q_s * conf['sample'].get(name, 0)

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
                '汇率后': round((df_normal[n_col].sum() + df_normal[p_col].sum()) / conf['exchange_rate'], 2),
                'P列折后价总和': df_normal[p_col].sum(), '售出产品成本总和': file_p_cost,
                '物流成本总和': file_l_cost, '寄样支出总和': file_s_cost
            })
            product_detail_data[order_date_str] = post_split_qty

        except Exception as e:
            print(f"❌ 处理失败 {file_name}: {e}")

    # -----------------------
    # 4. 数据导出逻辑 ---
    # -----------------------
    if summary_list:
        df_summary = pd.DataFrame(summary_list)
        df_summary['t'] = pd.to_datetime(df_summary['订单日期'], format='%d/%m/%Y', errors='coerce')
        df_summary = df_summary.sort_values('t').drop(columns=['t'])

        df_cat_raw = pd.DataFrame(category_detail_list).groupby(['日期', '大类']).sum().reset_index()
        df_cat_raw['t'] = pd.to_datetime(df_cat_raw['日期'], dayfirst=True)
        df_cat_raw['汇率后(销售额)'] = (df_cat_raw['销售额'] / conf['exchange_rate']).round(2)
        cols_order = ['日期', '大类', '销售额', '汇率后(销售额)', 'P列折后价', '产品成本', '物流成本', '寄样支出']
        df_cat = df_cat_raw[cols_order].copy()
        df_cat = df_cat.assign(t=df_cat_raw['t']).sort_values(['t', '大类']).drop(columns=['t'])

        df_pivot_raw = pd.DataFrame.from_dict(product_detail_data, orient='index').fillna(0).astype(int)
        df_pivot_raw.index = pd.to_datetime(df_pivot_raw.index, dayfirst=True)
        df_pivot_raw = df_pivot_raw.sort_index()
        df_pivot_raw.index = df_pivot_raw.index.strftime('%d/%m/%Y')
        valid_cols = [c for c in SORT_ORDER_LIST if c in df_pivot_raw.columns]
        other_cols = [c for c in df_pivot_raw.columns if c not in SORT_ORDER_LIST]
        df_pivot = df_pivot_raw[valid_cols + other_cols].T
        df_pivot.index.name = '商品名称'
        df_pivot.loc['汇总'] = df_pivot.sum(axis=0)

        if all_samples_collector:
            df_all_samples = pd.concat(all_samples_collector)
            df_sample_pivot = df_all_samples.pivot_table(
                index='Mapped Name',
                columns='日期',
                aggfunc='size',
                fill_value=0
            )
            df_sample_pivot.columns = pd.to_datetime(df_sample_pivot.columns, dayfirst=True)
            df_sample_pivot = df_sample_pivot.sort_index(axis=1)
            df_sample_pivot.columns = df_sample_pivot.columns.strftime('%d/%m/%Y')
            df_sample_pivot['总样品数'] = df_sample_pivot.sum(axis=1)
            df_sample_pivot.index.name = '商品名称'
            df_sample_pivot = df_sample_pivot.sort_values(by='总样品数', ascending=False)
        else:
            df_sample_pivot = pd.DataFrame([["今日无样品数据"]], columns=["提示"])

        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            df_summary.to_excel(writer, sheet_name='Weekly Summary', index=False)
            df_cat.to_excel(writer, sheet_name='Category Aggregation', index=False)
            df_pivot.to_excel(writer, sheet_name='Quantity Pivot', index=True)
            df_sample_pivot.to_excel(writer, sheet_name='Sample Statistics', index=True)

        print(f"✅ [{conf['cn_name']}] 报表已生成！")


if __name__ == "__main__":
    for store in ['local', 'cross_border']:
        run_vietnam_report(store)
    print("\n✨ 越南站所有店铺处理完毕。")