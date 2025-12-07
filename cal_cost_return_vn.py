import pandas as pd
import glob
import os

# -----------------------
# 文件夹路径
# -----------------------
folder_path = './data_VN'

# 获取文件夹下所有 CSV 文件
file_list = glob.glob(os.path.join(folder_path, '*.csv'))

# 关键列 (全局配置)
status_col = 'Order Status'
quantity_col = 'Normal or Pre-order'
sku_col = 'Seller SKU'
n_col = 'SKU Platform Discount'
p_col = 'SKU Subtotal After Discount'
r_col = 'Original Shipping Fee'

# SKU 映射 (全局配置)
sku_mapping = {
    'FCBL001': '身体乳单瓶',
    '2FCBL001': '身体乳双瓶',
    '3FCBL001': '身体乳三瓶',
    'FCSS001': '防晒霜单瓶',
    '2FCSS001': '防晒霜双瓶',
    '3FCSS001': '防晒霜三瓶',
    'FCSERUM001': '精华液单瓶',
    '2FCSERUM001': '精华液双瓶',
    '3FCSERUM001': '精华液三瓶',
    'FCB&S001': '身体乳和防晒霜组合套装',
    'FCBLGREENTEE002': '身体乳单瓶',
    '2FCBLGREENTEE002': '身体乳双瓶',
    '3FCBLGREENTEE002': '身体乳三瓶',
    'FCBLBO300': '身体乳300ml单瓶',
    '2FCBLBO300': '身体乳300ml双瓶',
    '3FCBLBO300': '身体乳300ml三瓶'
}

# 产品成本 (功能一) (全局配置)
cost_mapping = {
    '身体乳单瓶': 9.36,
    '身体乳双瓶': 18.72,
    '身体乳三瓶': 28.08,
    '防晒霜单瓶': 8.37,
    '防晒霜双瓶': 16.74,
    '防晒霜三瓶': 25.11,
    '身体乳和防晒霜组合套装': 30.0,
    '精华液单瓶': 9.02,
    '精华液双瓶': 18.04,
    '精华液三瓶': 27.06,
    '身体乳300ml单瓶': 3.08,
    '身体乳300ml双瓶': 6.16,
    '身体乳300ml三瓶': 9.24
}

# 物流成本 (功能一) (全局配置)
logistics_cost_mapping = {
    '身体乳单瓶': 6.22,
    '身体乳双瓶': 9.53,
    '身体乳三瓶': 13.1,
    '防晒霜单瓶': 3.15,
    '防晒霜双瓶': 3.9,
    '防晒霜三瓶': 4.65,
    '身体乳和防晒霜组合套装': 5.0,
    '精华液单瓶': 3.21,
    '精华液双瓶': 4.02,
    '精华液三瓶': 4.83,
    '身体乳300ml单瓶': 2.03,
    '身体乳300ml双瓶': 4.06,
    '身体乳300ml三瓶': 6.09
}

# 寄样成本映射 (功能二) (全局配置)
cost_mapping_function2 = {
    '身体乳单瓶': 23.42,
    '身体乳双瓶': 32.96,
    '身体乳三瓶': 43.02,
    '防晒霜单瓶': 19.96,
    '防晒霜双瓶': 28.53,
    '防晒霜三瓶': 37.10,
    '精华液单瓶': 20.21,
    '精华液双瓶': 29.23,
    '精华液三瓶': 38.25,
    '身体乳300ml单瓶': 14.12,
    '身体乳300ml双瓶': 23.6,
    '身体乳300ml三瓶': 30.3,
}

# 固定顺序输出 (功能一/二) (全局配置)
SORT_ORDER_LIST = ['身体乳300ml单瓶', '身体乳300ml双瓶', '身体乳300ml三瓶', '身体乳单瓶', '身体乳双瓶', '身体乳三瓶',
                   '防晒霜单瓶', '防晒霜双瓶', '防晒霜三瓶',
                   '精华液单瓶', '精华液双瓶', '精华液三瓶', '身体乳和防晒霜组合套装']

# 初始化结果列表，用于存储每个文件的汇总数据 (Sheet 1)
summary_data = []

# 初始化产品数量明细字典 (用于 Sheet 2)
product_detail_data = {}

# -----------------------
# 循环处理每个文件
# -----------------------
for file_path in file_list:
    # 提取文件名
    file_name = os.path.basename(file_path)
    print(f"\n==================== 处理文件: {file_name} ====================\n")

    try:
        # 读取 CSV
        df = pd.read_csv(file_path)
    except Exception as e:
        print(f"读取文件 {file_name} 失败，跳过: {e}")
        continue

    # 初始化当前文件的关键指标
    current_summary = {
        '文件名称': file_name,
        '订单创建时间': 'N/A',
        '销售总额(N+P总和)': 0.0,
        '销售总额(N+P总和)/3600': 0.0,
        'P列折后价总和': 0.0,
        '售出产品成本总和': 0.0,
        '物流成本总和': 0.0,
        '寄样总支出（产品+物流）总和': 0.0
    }

    # 用于 Sheet 2 索引的订单创建日期（字符串形式）
    order_date_str = 'N/A'

    # -----------------------
    # 读取第二行 Created Time 列作为订单创建时间
    # -----------------------
    if 'Paid Time' in df.columns and len(df) > 1:
        raw_creation_time = df.iloc[1]['Created Time']
        try:
            # 明确指定输入格式，解决 UserWarning
            datetime_obj = pd.to_datetime(raw_creation_time, format='%d/%m/%Y %H:%M:%S', errors='coerce')

            if pd.isna(datetime_obj):
                raise ValueError("Date parsing failed")

            # 格式化日期为 DD/MM/YYYY
            order_date_str = datetime_obj.strftime('%d/%m/%Y')

            current_summary['订单创建时间'] = order_date_str
            print(f"订单创建时间: {order_date_str}")

        except Exception:
            # 如果转换失败，则回退到只取空格前的内容
            if isinstance(raw_creation_time, str) and ' ' in raw_creation_time:
                order_date_str = raw_creation_time.split(' ')[0]
            else:
                order_date_str = raw_creation_time
            current_summary['订单创建时间'] = order_date_str
            print(f"订单创建时间: {order_date_str} (未进行标准格式化)")
    else:
        print("未找到 Paid Time 列或文件行数不足")
        # 确保日期缺失时有一个唯一键
        order_date_str = file_name

        # -------------------------------------------
    # 功能一：排除已取消，保留 Normal 订单 (用于销售、产品成本、物流成本计算)
    # -------------------------------------------
    df_normal = df[(df[status_col] != '已取消') & (df[quantity_col] == 'Normal')].copy()

    # 清洗 N、P、R 列，确保它们是数字类型
    df_normal[n_col] = pd.to_numeric(df_normal[n_col], errors='coerce')
    df_normal[p_col] = pd.to_numeric(df_normal[p_col], errors='coerce')
    df_normal[r_col] = pd.to_numeric(df_normal[r_col], errors='coerce')

    df_normal['Product Name'] = df_normal[sku_col].map(lambda x: sku_mapping.get(x, x))

    # 统计拆分前数量
    sku_summary = df_normal.groupby('Product Name').size().reset_index(name='Product Quantity_PreSplit')
    sku_summary['Product Quantity_PostSplit'] = sku_summary['Product Quantity_PreSplit']

    # 组合套装拆分
    combo_name = '身体乳和防晒霜组合套装'
    if combo_name in sku_summary['Product Name'].values:
        combo_qty = sku_summary.loc[sku_summary['Product Name'] == combo_name, 'Product Quantity_PostSplit'].values[0]
        for single in ['身体乳单瓶', '防晒霜单瓶']:
            if single in sku_summary['Product Name'].values:
                sku_summary.loc[sku_summary['Product Name'] == single, 'Product Quantity_PostSplit'] += combo_qty
            else:
                sku_summary = pd.concat([sku_summary,
                                         pd.DataFrame({'Product Name': [single],
                                                       'Product Quantity_PreSplit': [0],
                                                       'Product Quantity_PostSplit': [combo_qty]})],
                                        ignore_index=True)
        sku_summary.loc[sku_summary['Product Name'] == combo_name, 'Product Quantity_PostSplit'] = 0

    # 产品成本和物流成本计算
    sku_summary['Product Cost'] = sku_summary['Product Name'].map(cost_mapping).fillna(0)
    sku_summary['Total Product Cost'] = (sku_summary['Product Quantity_PostSplit'] * sku_summary['Product Cost']).round(
        2)
    sku_summary['Shipping Cost'] = sku_summary['Product Name'].map(logistics_cost_mapping).fillna(0)
    sku_summary['Total Shipping Cost'] = (
                sku_summary['Product Quantity_PostSplit'] * sku_summary['Shipping Cost']).round(2)

    # 总成本 (仅用于显示，不用于 Excel 汇总)
    sku_summary['Total Cost Including Logistics'] = (
                sku_summary['Total Product Cost'] + sku_summary['Total Shipping Cost']).round(2)

    # 固定顺序输出
    sku_summary['Sort Order'] = sku_summary['Product Name'].apply(
        lambda x: SORT_ORDER_LIST.index(x) if x in SORT_ORDER_LIST else len(SORT_ORDER_LIST))
    sku_summary = sku_summary.sort_values(by='Sort Order').drop(columns='Sort Order').reset_index(drop=True)

    # ------------------ 打印输出 (保留) ------------------
    print("\n拆分前 vs 拆分后数量对比表：")
    print(sku_summary[['Product Name', 'Product Quantity_PreSplit', 'Product Quantity_PostSplit']])

    print("\n拆分后的产品成本明细及汇总：")
    print(sku_summary[['Product Name', 'Product Quantity_PostSplit', 'Product Cost', 'Total Product Cost']])
    total_product_cost = sku_summary['Total Product Cost'].sum()
    print(f"售出产品成本总和: {total_product_cost:.2f}")

    print("\n拆分后的物流成本明细及汇总：")
    print(sku_summary[['Product Name', 'Product Quantity_PostSplit', 'Shipping Cost', 'Total Shipping Cost']])
    total_shipping_cost = sku_summary['Total Shipping Cost'].sum()
    print(f"物流成本总和: {total_shipping_cost:.2f}")

    # N 列折扣总和、P 列折后价总和及 N+P 总和
    n_sum = df_normal[n_col].sum()
    p_sum = df_normal[p_col].sum()
    sales_total = n_sum + p_sum
    print(f"\nN列折扣总和: {n_sum:.2f}")
    print(f"P列折后价总和: {p_sum:.2f}")
    print(f"销售总额(N+P总和): {sales_total:.2f}")
    # ------------------------------------------------------------

    # 填充功能一的关键指标 (用于 Sheet 1 汇总)
    current_summary['P列折后价总和'] = p_sum
    current_summary['销售总额(N+P总和)'] = sales_total
    current_summary['销售总额(N+P总和)/3600'] = round(sales_total / 3600.0, 2) if sales_total else 0.0
    current_summary['售出产品成本总和'] = total_product_cost
    current_summary['物流成本总和'] = total_shipping_cost

    # ------------------ 收集 Sheet 2 产品数量明细 ------------------
    # 创建当前文件的产品数量 Series，以 Product Name 为索引
    current_product_qty = sku_summary.set_index('Product Name')['Product Quantity_PostSplit']

    # 将 Series 转换为字典，使用订单创建日期作为键
    # 键是 '订单创建时间' (order_date_str)，值是产品名称到数量的映射
    product_detail_data[order_date_str] = current_product_qty.to_dict()
    # ------------------------------------------------------------

    # -------------------------------------------
    # 功能二：未取消且E列为空的订单，拆分组合套装并计算成本 (用于寄样成本计算)
    # -------------------------------------------

    # 1️⃣ 筛选未取消订单且E列为空（NaN 或空字符串）
    df_not_cancelled_empty = df[(df[status_col] != '已取消') &
                                (df[quantity_col].isna() | (df[quantity_col] == ''))].copy()

    # 2️⃣ 映射产品名称
    df_not_cancelled_empty['Product Name'] = df_not_cancelled_empty[sku_col].map(lambda x: sku_mapping.get(x, x))

    # 3️⃣ 按产品统计拆分前数量
    empty_summary = df_not_cancelled_empty.groupby('Product Name').size().reset_index(name='Product Quantity_PreSplit')
    empty_summary['Product Quantity_PostSplit'] = empty_summary['Product Quantity_PreSplit']

    # 4️⃣ 拆分组合套装
    if combo_name in empty_summary['Product Name'].values:
        combo_qty = empty_summary.loc[empty_summary['Product Name'] == combo_name, 'Product Quantity_PostSplit'].values[
            0]
        for single in ['身体乳单瓶', '防晒霜单瓶']:
            if single in empty_summary['Product Name'].values:
                empty_summary.loc[empty_summary['Product Name'] == single, 'Product Quantity_PostSplit'] += combo_qty
            else:
                empty_summary = pd.concat([empty_summary,
                                           pd.DataFrame({'Product Name': [single],
                                                         'Product Quantity_PreSplit': [0],
                                                         'Product Quantity_PostSplit': [combo_qty]})],
                                          ignore_index=True)
        empty_summary.loc[empty_summary['Product Name'] == combo_name, 'Product Quantity_PostSplit'] = 0

    # 5️⃣ 寄样成本计算
    empty_summary['Product Cost'] = empty_summary['Product Name'].map(cost_mapping_function2).fillna(0)
    empty_summary['Total Product Cost'] = (
            empty_summary['Product Quantity_PostSplit'] * empty_summary['Product Cost']).round(2)

    # 6️⃣ 固定顺序输出
    empty_summary['Sort Order'] = empty_summary['Product Name'].apply(
        lambda x: SORT_ORDER_LIST.index(x) if x in SORT_ORDER_LIST else len(SORT_ORDER_LIST))
    empty_summary = empty_summary.sort_values(by='Sort Order').drop(columns='Sort Order').reset_index(drop=True)

    # ------------------ 打印输出 (保留) ------------------
    print("\n拆分前 vs 拆分后数量对比（功能二）：")
    print(empty_summary[['Product Name', 'Product Quantity_PreSplit', 'Product Quantity_PostSplit']])

    print("\n产品成本明细及汇总（功能二）：")
    print(empty_summary[['Product Name', 'Product Quantity_PostSplit', 'Product Cost', 'Total Product Cost']])
    total_sample_cost = empty_summary['Total Product Cost'].sum()
    print(f"寄样总支出（产品+物流）总和: {total_sample_cost:.2f}")
    # ------------------------------------------------------------

    # 填充功能二的关键指标 (用于 Sheet 1 汇总)
    current_summary['寄样总支出（产品+物流）总和'] = total_sample_cost

    # 收集当前文件的所有汇总数据 (Sheet 1)
    summary_data.append(current_summary)

# -------------------------------------------
# 结果汇总并输出 Excel 文件
# -------------------------------------------
if summary_data:
    # --- Sheet 1: 汇总数据 (Summary Report) ---
    final_df = pd.DataFrame(summary_data)

    # 重新排列和命名列，确保输出顺序与要求一致
    final_df = final_df[[
        '文件名称',
        '订单创建时间',
        '销售总额(N+P总和)',
        '销售总额(N+P总和)/3600',
        'P列折后价总和',
        '售出产品成本总和',
        '物流成本总和',
        '寄样总支出（产品+物流）总和'
    ]]

    # --- Sheet 2: 产品数量明细 (Product Detail) ---
    # 1. 将字典转换为 DataFrame，订单创建日期作为索引 (行)
    detail_df = pd.DataFrame.from_dict(product_detail_data, orient='index').fillna(0).astype(int)

    # 2. 确保列名 (产品名称) 按照 SORT_ORDER_LIST 的顺序排列
    ordered_columns = [col for col in SORT_ORDER_LIST if col in detail_df.columns]
    detail_df = detail_df[ordered_columns]

    # 3. 执行转置 (.T)，使产品名称成为索引 (行)，订单创建时间成为列
    detail_df = detail_df.T

    # 4. 重命名索引 (行标题)
    detail_df.index.name = '产品名称'

    output_filename = 'summary_report.xlsx'

    # 使用 ExcelWriter 写入多 Sheet
    try:
        # 即使您已安装 xlsxwriter，但由于环境问题，这里改为使用 openpyxl
        # (如果 openpyxl 也未安装，请使用 pip install openpyxl 安装)
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            # 写入 Sheet 1
            final_df.to_excel(writer, sheet_name='Summary Report', index=False)

            # 写入 Sheet 2 (已转置)
            detail_df.to_excel(writer, sheet_name='Product Quantity Detail')  # index=True by default

        print("\n" + "=" * 50)
        print(f"✅ 所有文件处理完毕，汇总报告已成功保存到：{output_filename}")
        print("报告包含两个 Sheet: 'Summary Report' 和 'Product Quantity Detail' (产品名称已竖排)")
        print("=" * 50)
    except Exception as e:
        print(f"\n❌ 写入 Excel 文件失败: {e}")
        print(
            "提示：如果继续出现 'No module named...' 错误，请确保运行脚本的 Python 环境已激活并正确安装了 'xlsxwriter' 或 'openpyxl' 模块。")
else:
    print("\n没有找到可处理的文件或所有文件均处理失败。未生成 Excel 报告。")