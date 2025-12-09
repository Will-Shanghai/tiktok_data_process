import pandas as pd
import numpy as np
import os

# 1. 配置文件的路径：读取当前目录（'.'）下的所有文件
current_directory = '../data/'
print(f"✅ 准备处理当前目录下的所有Excel文件...")

# --- 定义标准内部列名和映射关系 ---
# 使用标准内部名称来统一代码中对列的引用
STD_ITEM_COL = '商品'
STD_VV_COL = '视频VV'
STD_ORDER_COL = '视频成交订单数'
STD_CTR_COL = '点击率'
STD_CCR_COL = '点击成交转化率'

# 映射字典：兼容中英双语列名
COLUMN_MAP = {
    STD_ITEM_COL: ['商品', 'Products'],
    STD_VV_COL: ['视频vv', 'VV'],
    STD_ORDER_COL: ['视频成交订单数', 'Orders'],
    STD_CTR_COL: ['点击率（视频）', 'Click-through rate (Video)'],
    STD_CCR_COL: ['点击成交转化率（视频）', 'Click-to-order rate (Video)'],
}

# 定义需要分析的产品类别
TARGET_CATEGORIES = ['精华液', '黑鸦片身体乳', '500ml身体乳A链', '500ml身体乳B链', '防晒霜']

# 2. 获取当前目录下所有 Excel 文件
try:
    file_list = [f for f in os.listdir(current_directory) if f.endswith('.xlsx')]
except Exception as e:
    print(f"❌ 错误：读取当前目录文件列表失败。错误信息: {e}")
    exit()

if not file_list:
    print("⚠️ 警告：当前目录下未找到任何 .xlsx 文件。程序终止。")
    exit()

print(f"✅ 找到 {len(file_list)} 个 .xlsx 文件，开始循环处理...")

# 初始化最终汇总数据的列表
final_summary_data = []

# 循环处理每个文件
for file_name in file_list:
    file_path = os.path.join(current_directory, file_name)
    print(f"\n================ 开始处理文件: {file_name} ================")

    # 3. 提取日期范围并读取数据
    try:
        # 临时读取文件，获取 A1 单元格日期
        df_date = pd.read_excel(file_path, header=None, sheet_name=0)
        first_line = str(df_date.iloc[0, 0]).strip()
        first_line = first_line.replace('[日期范围]:', '').strip().replace('"', '').replace(',', '')

        print("-------------------------------------------------")
        print(f"📅 提取的日期范围: {first_line}")
        print("-------------------------------------------------")

        # 读取主数据，跳过前两行
        df = pd.read_excel(file_path, skiprows=2)
        print("✅ Excel 数据文件读取成功！")

    except Exception as e:
        print(f"❌ 错误：读取文件 '{file_name}' 失败。错误信息: {e}。跳过此文件。")
        continue

    # --- 关键兼容性处理：检查列名并重命名 ---
    df.columns = df.columns.str.strip()  # 清理列名空格

    missing_cols_std = []

    for std_col, possible_names in COLUMN_MAP.items():
        found = False
        for name in possible_names:
            if name in df.columns:
                # 找到匹配的列，将其重命名为标准列名
                df.rename(columns={name: std_col}, inplace=True)
                found = True
                break

        if not found:
            # 如果核心列找不到任何匹配项，记录下来
            missing_cols_std.append(std_col)

    if missing_cols_std:
        print(f"❌ 致命错误：文件 '{file_name}' 中缺少核心数据（标准列）：{missing_cols_std}。跳过此文件。")
        continue
    # --- 兼容性处理结束 ---

    # --- 数据清洗与预处理 ---

    # 4. 清理百分比数据 (CTR 和 CCR)，使用标准列名
    cols_to_clean = [STD_CTR_COL, STD_CCR_COL]
    for col in cols_to_clean:
        df[col] = (
            df[col]
            .astype(str)
            .str.replace('%', '', regex=False)
            .replace('', '0')
        )
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) / 100

    # 5. 确保订单数和 VV 列是数值类型，使用标准列名
    numeric_cols = [STD_ORDER_COL, STD_VV_COL]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)


    # 6. 定义映射函数
    def map_product(product_name):
        """根据商品名称映射到指定的产品分类"""
        if pd.isna(product_name):
            return '其他/未匹配'

        product_name_str = str(product_name)

        if 'Serum' in product_name_str:
            return '精华液'
        elif 'Black Opium' in product_name_str:
            return '黑鸦片身体乳'
        elif 'Niacinamide' in product_name_str:
            return '500ml身体乳B链'
        elif '500g' in product_name_str:
            return '500ml身体乳A链'
        elif 'SPF 50 +' in product_name_str:
            return '防晒霜'
        else:
            return '其他/未匹配'


    # 应用映射函数
    df['产品分类'] = df[STD_ITEM_COL].apply(map_product)

    # --- 分组计算与统计 (只针对出单视频) ---
    df_analysis = df[df['产品分类'].isin(TARGET_CATEGORIES)].copy()
    results = {}

    for category in TARGET_CATEGORIES:
        group = df_analysis[df_analysis['产品分类'] == category].copy()
        total_videos = len(group)

        # B. 出单视频子集
        videos_with_orders = group[group[STD_ORDER_COL] > 0].copy()
        count_with_orders = len(videos_with_orders)

        # C. 视频出单率 (结果保持浮点数形式)
        order_rate = count_with_orders / total_videos if total_videos > 0 else 0

        # --- 新增：计算该分类下所有视频的 VV 总和 ---
        total_vv = group[STD_VV_COL].sum()

        # D. 计算各项指标 (仅对出单视频子集)
        if count_with_orders > 0:
            average_ctr = videos_with_orders[STD_CTR_COL].mean()
            average_ccr = videos_with_orders[STD_CCR_COL].mean()
            average_vv = videos_with_orders[STD_VV_COL].mean()
        else:
            # 如果没有出单视频，所有指标设为 NaN
            average_ctr = np.nan
            average_ccr = np.nan
            average_vv = np.nan

        # 存储结果到字典
        results[category] = {
            '总视频数': total_videos,
            '出单视频数': count_with_orders,
            '视频VV总和': total_vv,  # 新增
            '视频出单率': order_rate,
            '平均点击率': average_ctr,
            '平均点击成交率': average_ccr,
            '视频VV的平均值': average_vv,
        }

        # E. 收集数据到最终汇总列表
        final_summary_data.append({
            '文件名': file_name,
            '日期范围': first_line,  # 附加日期范围信息
            '产品分类': category,
            '总视频数': total_videos,
            '出单视频数': count_with_orders,
            '视频出单率': order_rate,  # 保持浮点数
            '平均点击率': average_ctr,  # 保持浮点数
            '平均点击成交率': average_ccr,  # 保持浮点数
            '视频VV的平均值': average_vv,  # 保持数值
            '视频VV总和': total_vv,  # 新增
        })

    # 12. 打印当前文件结果 (屏幕输出)
    print("\n--- 关键数据分析结果 (屏幕输出) ---")
    for category, data in results.items():
        print(f"\n======== 【{category}】 ========")
        print(f"  总视频数: {data['总视频数']:.0f}")
        print(f"  出单视频数: {data['出单视频数']:.0f}")
        # 新增：打印视频VV总和
        print(f"  视频VV总和: {round(data['视频VV总和']):,.0f}")
        print(f"  视频出单率: {data['视频出单率']:.2%}")

        print("--- 仅限出单视频指标 ---")

        if pd.isna(data['平均点击率']):
            print(f"  平均点击率: N/A (无出单视频)")
        else:
            print(f"  平均点击率: {data['平均点击率']:.2%}")

        if pd.isna(data['平均点击成交率']):
            print(f"  平均点击成交率: N/A (无出单视频)")
        else:
            print(f"  平均点击成交率: {data['平均点击成交率']:.2%}")

        if pd.isna(data['视频VV的平均值']):
            print(f"  视频VV的平均值: N/A (无出单视频)")
        else:
            print(f"  视频VV的平均值: {round(data['视频VV的平均值']):,.0f}")

    print("====================================================")

# 13. 所有文件处理完毕后，生成最终报告文件
if final_summary_data:
    final_df = pd.DataFrame(final_summary_data)
    # 设置带目录的输出文件名
    output_file_name = '../result/analysis_summary_report.xlsx'

    # 将 DataFrame 导出为 Excel
    try:
        # 确保目标目录存在，如果不存在则创建
        output_dir = os.path.dirname(output_file_name)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir, exist_ok=True)
            print(f"📁 目标输出目录 '{output_dir}' 已创建。")

        final_df.to_excel(output_file_name, index=False, sheet_name='汇总报告')
        print(f"\n\n🎉 所有文件处理完成！")
        print(f"📊 汇总数据已成功保存至文件: {output_file_name}")
    except Exception as e:
        print(f"❌ 错误：保存汇总报告失败。请检查是否有权限创建目录或文件。错误信息: {e}")
else:
    print("\n🎉 所有文件处理完成！但未收集到有效数据。")