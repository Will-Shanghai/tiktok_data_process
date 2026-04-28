import pandas as pd
import os
import glob


# -----------------------
# 1. 基础映射与工具配置
# -----------------------
def map_sku(name):
    if pd.isna(name): return name
    # 按照关键词匹配中文简称
    if '最新アップグレード版' in name: return '黑色睫毛夹'
    if '2025年最新版アップグレード 5D' in name: return '白色睫毛推'
    if ('正規品 虫除け リアル' in name) or ('Dragonfly' in name): return '蜻蜓两个装'
    if '収納 ヘアアイロンポーチ 耐热300度' in name: return '卷发棒隔热袋'
    if '髪飾り 13 / 14点' in name: return '发饰水银13/14点'
    if '手作り 11点' in name: return '发饰水银11点'
    if 'Bluetooth 5.4ヘッドフォン' in name: return '屏显耳机'
    if '電働眉剃' in name: return '电动剃眉刀'
    if '車のサンバイ' in name: return '车载手机支架'
    if ('電働1台多用耳毛刀眉毛刀' in name) or ('耳毛 鼻毛切り' in name): return '电动鼻毛刀'
    if '3D風景ミ二チュア シ-ンコレクション' in name: return '剪影贴纸'
    if '高濃度のマグネシウム配合' in name: return '镁膏'
    if 'Taba ソフトストロ' in name: return '草莓捏捏乐'
    if '脱毛器' in name: return '脱毛器'
    if '手帳 DIY' in name: return '场景贴纸'
    if ('美容液' in name) or ('290g' in name): return '头发护理精华'
    if 'Bluetooth&ヘッドフォンMP 3' in name: return '蓝牙MP3'
    if '车充电器' in name: return '车载充电器'
    if '睫毛 フェイスケア' in name: return '新款睫毛夹'
    if 'イヤホンクリップ' in name: return '耳机夹'
    if 'グレインココナッツの手作りボール' in name: return '椰子捏捏乐'
    if 'ワイヤレス翻訳ヘッドフォンは、' in name: return '无线翻译耳机'
    if 'シガーライターソケット' in name: return '点烟器'
    return name


def parse_amount(value):
    if pd.isna(value): return 0.0
    s = str(value).replace(' ', '').replace('JPY', '').replace(',', '')
    try:
        return float(s)
    except:
        return 0.0


# -----------------------
# 2. 店铺差异化配置
# -----------------------
EXCHANGE_RATE = 20.0

# 产品成本
common_cost_1 = {
    '黑色睫毛夹': 8.86, '白色睫毛推': 8.9, '电动剃眉刀': 4.6, '车载手机支架': 4.6, '镁膏': 12,
    '发饰水银13/14点': 16.3, '发饰水银11点': 16.3, '电动鼻毛刀': 7.8, '头发护理精华': 10,
    '蓝牙MP3': 43, '车载充电器': 18.4, '脱毛器': 50, '场景贴纸': 16.8
}

# 本土店产品头程 + 尾程
common_logistics_1 = {
    '黑色睫毛夹': 22.65, '白色睫毛推': 22.45, '电动剃眉刀': 12.5, '车载手机支架': 17.5,
    '发饰水银13/14点': 22.2, '发饰水银11点': 22.2, '电动鼻毛刀': 14, '头发护理精华': 41,
    '蓝牙MP3': 24.02, '车载充电器': 23.47, '场景贴纸': 19
}

# 直邮店产品头程 + 尾程
common_logistics_2 = {
    '黑色睫毛夹': 22.65, '白色睫毛推': 22.45, '电动剃眉刀': 12.5, '车载手机支架': 17.5, '镁膏': 22,
    '发饰水银13/14点': 22.2, '发饰水银11点': 22.2, '电动鼻毛刀': 25, '头发护理精华': 25,
    '蓝牙MP3': 24.02, '车载充电器': 23.47
}

# 本土店寄样成本
sample_cost_1 = {
    '黑色睫毛夹': 31.51, '蜻蜓两个装': 26.18, '卷发棒隔热袋': 37.2, '头发护理精华': 51,  '场景贴纸': 35.8,
    '屏显耳机': 62, '电动剃眉刀': 17.1, '车载手机支架': 22.1, '发饰水银13/14点': 38.5, '电动鼻毛刀': 21.8
}

# 直邮店寄样成本
sample_cost_2 = {
    '黑色睫毛夹': 31.51, '蜻蜓两个装': 26.18, '卷发棒隔热袋': 37.2, '头发护理精华': 35, '镁膏': 37,
    '屏显耳机': 62, '电动剃眉刀': 17.1, '车载手机支架': 22.1, '发饰水银13/14点': 38.5, '电动鼻毛刀': 32.8
}

STORE_CONFIG = {
    'local': {
        'name_cn': '本土店',
        'cost_mapping_1': common_cost_1,
        'logistics_mapping_1': common_logistics_1,
        'cost_mapping_2': sample_cost_1
    },
    'direct': {
        'name_cn': '直邮店',
        'cost_mapping_1': common_cost_1,
        'logistics_mapping_1': common_logistics_2,
        'cost_mapping_2': sample_cost_2
    }
}


# -----------------------
# 3. 核心处理函数
# -----------------------
def run_report(store_key):
    config = STORE_CONFIG[store_key]
    current_script_dir = os.path.dirname(os.path.abspath(__file__))
    folder_path = os.path.normpath(os.path.join(current_script_dir, f'../data/data_JP/{store_key}'))
    output_filename = os.path.normpath(
        os.path.join(current_script_dir, f'../result/Weekly_Performance_Report_JP_{store_key}.xlsx'))

    file_list = glob.glob(os.path.join(folder_path, '*.csv'))
    if not file_list:
        print(f"⚠️  [{config['name_cn']}] 路径下未找到文件: {folder_path}")
        return

    print(f"🚀 开始处理 [{config['name_cn']}]，共 {len(file_list)} 个文件")

    status_col = 'Order Status'
    quantity_col = 'Normal or Pre-order'
    product_col = 'Product Name'
    n_col = 'SKU Platform Discount'
    p_col = 'SKU Subtotal After Discount'
    r_col = 'Original Shipping Fee'

    weekly_summary_list = []
    category_detail_list = []
    product_detail_data = {}
    all_samples_collector = []

    unrecognized_names = set()

    for file_path in file_list:
        file_name = os.path.basename(file_path)
        try:
            df = pd.read_csv(file_path)
        except Exception as e:
            print(f"读取失败 {file_name}: {e}")
            continue

        order_date_str = "N/A"
        if 'Created Time' in df.columns and len(df) > 0:
            try:
                order_date_str = pd.to_datetime(df.iloc[0]['Created Time']).strftime('%Y-%m-%d')
            except:
                order_date_str = str(df.iloc[0]['Created Time']).split(' ')[0]

        cancel_keywords = ['Canceled', 'Unpaid', 'キャンセル済み', '未払い', '已取消', '未支付', '部分发货后取消']

        # 1. 正常订单
        df_normal = df[~df[status_col].isin(cancel_keywords) & (df[quantity_col] == 'Normal')].copy()
        df_normal['Mapped Name'] = df_normal[product_col].map(map_sku)
        for col in [n_col, p_col, r_col]:
            df_normal[col] = df_normal[col].map(parse_amount)

        sku_counts = df_normal['Mapped Name'].value_counts().to_dict()
        product_detail_data[order_date_str] = sku_counts
        norm_agg = df_normal.groupby('Mapped Name').agg({n_col: 'sum', p_col: 'sum', r_col: 'sum'}).reset_index()

        # 2. 样品处理
        df_sample = df[
            ~df[status_col].isin(cancel_keywords) & (df[quantity_col].isna() | (df[quantity_col] == ''))].copy()
        df_sample['Mapped Name'] = df_sample[product_col].map(map_sku)

        if not df_sample.empty:
            sample_temp = df_sample[['Mapped Name']].copy()
            sample_temp['日期'] = order_date_str
            all_samples_collector.append(sample_temp)

        sample_counts = df_sample['Mapped Name'].value_counts().to_dict()

        # 3. 成本计算
        file_p_cost, file_l_cost, file_s_cost = 0, 0, 0
        all_names = set(list(sku_counts.keys()) + list(sample_counts.keys()))

        for name in all_names:
            q_norm = sku_counts.get(name, 0)
            q_samp = sample_counts.get(name, 0)
            row_sales = norm_agg[norm_agg['Mapped Name'] == name]

            s_val = row_sales[n_col].sum() + row_sales[p_col].sum() + row_sales[r_col].sum()
            p_val = row_sales[p_col].sum()

            p_cost = q_norm * config['cost_mapping_1'].get(name, 0)
            l_cost = q_norm * config['logistics_mapping_1'].get(name, 0)
            s_cost = q_samp * config['cost_mapping_2'].get(name, 0)

            file_p_cost += p_cost
            file_l_cost += l_cost
            file_s_cost += s_cost

            cat_name = name
            if any('\u3040' <= c <= '\u30ff' for c in str(cat_name)) and not any(
                    '\u4e00' <= c <= '\u9fff' for c in str(cat_name)):
                if str(cat_name) != 'nan':
                    unrecognized_names.add(cat_name)

            category_detail_list.append({
                '日期': order_date_str, '大类': cat_name, '销售额': s_val,
                'P列折后价': p_val, '产品成本': p_cost, '物流成本': l_cost, '寄样支出': s_cost
            })

        sales_total = df_normal[n_col].sum() + df_normal[p_col].sum() + df_normal[r_col].sum()
        weekly_summary_list.append({
            '文件名称': file_name, '订单创建时间': order_date_str, '销售总额(N+P+R)': round(sales_total, 2),
            '汇率后金额': round(sales_total / EXCHANGE_RATE, 2), 'P列折后价总和': round(df_normal[p_col].sum(), 2),
            '售出产品成本总和': round(file_p_cost, 2), '物流成本总和': round(file_l_cost, 2),
            '寄样总支出(含物流)': round(file_s_cost, 2)
        })

    # --- 4. 数据导出逻辑修正（重点：时间排序） ---
    if weekly_summary_list:
        # Sheet 1: Weekly Summary
        df_final_weekly = pd.DataFrame(weekly_summary_list)
        df_final_weekly['temp_date'] = pd.to_datetime(df_final_weekly['订单创建时间'])
        df_final_weekly = df_final_weekly.sort_values('temp_date').drop(columns=['temp_date'])

        # Sheet 2: Category Aggregation
        df_cat_base = pd.DataFrame(category_detail_list)
        df_cat_summary = df_cat_base.groupby(['日期', '大类']).sum().reset_index()
        df_cat_summary['汇率后金额'] = (df_cat_summary['销售额'] / EXCHANGE_RATE).round(2)
        cols_order = ['日期', '大类', '销售额', '汇率后金额', 'P列折后价', '产品成本', '物流成本', '寄样支出']
        df_cat_summary = df_cat_summary[cols_order].copy()
        df_cat_summary = df_cat_summary.assign(temp_date=pd.to_datetime(df_cat_summary['日期'])) \
            .sort_values(by=['temp_date', '大类']) \
            .drop(columns=['temp_date'])

        # Sheet 3: Product Quantity Detail (增加日期列排序逻辑)
        df_detail_raw = pd.DataFrame.from_dict(product_detail_data, orient='index').fillna(0).astype(int)
        # 显式按照索引（即日期）进行排序
        df_detail_raw.index = pd.to_datetime(df_detail_raw.index)
        df_detail_raw = df_detail_raw.sort_index()
        # 转回字符串格式并转置
        df_detail_raw.index = df_detail_raw.index.strftime('%Y-%m-%d')
        df_detail = df_detail_raw.T
        df_detail.index.name = '产品名称'
        # 计算汇总行（排除表头，最后添加）
        df_detail.loc['汇总'] = df_detail.sum(axis=0)

        # Sheet 4: Sample Statistics (增加日期列排序逻辑)
        if all_samples_collector:
            df_all_samples = pd.concat(all_samples_collector)
            df_sample_pivot = df_all_samples.pivot_table(index='Mapped Name', columns='日期', aggfunc='size',
                                                         fill_value=0)
            # 对列名（日期）进行排序
            sorted_cols = sorted(df_sample_pivot.columns, key=lambda x: pd.to_datetime(x))
            df_sample_pivot = df_sample_pivot[sorted_cols]
            # 格式化日期列名
            df_sample_pivot.columns = [pd.to_datetime(c).strftime('%Y-%m-%d') for c in df_sample_pivot.columns]
            df_sample_pivot['总样品数'] = df_sample_pivot.sum(axis=1)
            df_sample_pivot.index.name = '产品名称'
            df_sample_pivot = df_sample_pivot.sort_values(by='总样品数', ascending=False)
            df_sample_pivot.loc['汇总'] = df_sample_pivot.sum(axis=0)
        else:
            df_sample_pivot = pd.DataFrame([["无样品数据"]], columns=["提示"])

        try:
            if not os.path.exists(os.path.dirname(output_filename)):
                os.makedirs(os.path.dirname(output_filename))

            with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
                df_final_weekly.to_excel(writer, sheet_name='Weekly Summary', index=False)
                df_cat_summary.to_excel(writer, sheet_name='Category Aggregation', index=False)
                df_detail.to_excel(writer, sheet_name='Product Quantity Detail', index=True)
                df_sample_pivot.to_excel(writer, sheet_name='Sample Statistics', index=True)
            print(f"✅ [{config['name_cn']}] 处理完成，数据已按日期线性排序。")

            if unrecognized_names:
                print(f"💡 提醒：以下商品仍显示日文（未匹配简称）:")
                for n in sorted(list(unrecognized_names)):
                    print(f"   - {n}")
        except Exception as e:
            print(f"❌ [{config['name_cn']}] 保存失败: {e}")


if __name__ == "__main__":
    for store in ['local', 'direct']:
        run_report(store)
    print("\n✨ 任务执行完毕。")