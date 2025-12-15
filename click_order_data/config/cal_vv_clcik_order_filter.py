import pandas as pd
import numpy as np
import os
import re
from datetime import datetime  # 引入 datetime 模块用于获取时间戳

# 1. 配置文件的路径：读取当前目录（'.'）下的所有文件
current_directory = '../data/'
print(f"✅ 准备处理数据目录: {current_directory}")

# --- 定义标准内部列名和映射关系 ---
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

# --- 2. 站点产品分类隔离定义 ---
VN_CATEGORIES = ['精华液', '黑鸦片身体乳', '500ml身体乳A链', '500ml身体乳B链', '防晒霜']
JP_CATEGORIES = ['车载手机支架', '电动修眉刀']


# --- 3. 站点识别函数（基于 A3 单元格内容） ---
def identify_site_by_a3(a3_content):
    """
    根据数据框的第三行第一列（A3）的内容判断站点是越南站 (VN) 还是日本站 (JP)。
    """
    a3_content = str(a3_content).strip()

    # 假设 '达人昵称', '商品' 等中文/越南语词汇代表越南站点 (VN)
    if '达人昵称' in a3_content or '商品' in a3_content:
        return "VN", VN_CATEGORIES

    # 假设 'Creator', 'Products' 等英文/日文词汇代表日本站点 (JP)
    elif 'Creator' in a3_content or 'Products' in a3_content:
        return "JP", JP_CATEGORIES

    else:
        # 如果 A3 不包含明确的中文/越南语特征，我们默认为日本站
        print(f"⚠️ 警告：无法通过 A3 单元格内容 '{a3_content}' 明确识别站点。默认为日本站 (JP)。")
        return "JP", JP_CATEGORIES


# 4. 获取文件列表 (路径已更新为 '../data/')
try:
    if not os.path.exists(current_directory):
        print(f"❌ 错误：数据目录 '{current_directory}' 不存在。请创建该目录并放入文件。")
        exit()

    file_list = [f for f in os.listdir(current_directory) if f.endswith('.xlsx')]
except Exception as e:
    print(f"❌ 错误：读取数据目录文件列表失败。错误信息: {e}")
    exit()

if not file_list:
    print(f"⚠️ 警告：数据目录 '{current_directory}' 下未找到任何 .xlsx 文件。程序终止。")
    exit()

print(f"✅ 找到 {len(file_list)} 个 .xlsx 文件，开始循环处理...")

# 初始化最终汇总数据的列表
final_summary_data = []

# 循环处理每个文件
for file_name in file_list:
    file_path = os.path.join(current_directory, file_name)
    print(f"\n================ 开始处理文件: {file_name} ================")

    # 5. 提取日期范围并读取数据
    try:
        # 临时读取文件，获取 A1 单元格日期
        df_date = pd.read_excel(file_path, header=None, sheet_name=0)
        first_line = str(df_date.iloc[0, 0]).strip()

        # 清理日期字符串
        date_range_str = first_line.replace('[日期范围]:', '').strip().replace('"', '').replace(',', '')

        # 匹配日期范围字符串，用于文件名 (例如 '2025-11-24 ~ 2025-11-30')
        dates = re.findall(r'\d{4}[-/]\d{2}[-/]\d{2}', date_range_str)
        if len(dates) >= 2:
            # 统一日期格式为 YYYY-MM-DD ~ YYYY-MM-DD
            standard_date_range = f"{dates[0].replace('/', '-') or dates[0]} ~ {dates[1].replace('/', '-') or dates[1]}"
        else:
            standard_date_range = "Unknown_Date"

        print("-------------------------------------------------")
        print(f"📅 提取的日期范围: {date_range_str}")
        print("-------------------------------------------------")

        # 读取主数据，跳过前两行 (第三行成为表头)
        df = pd.read_excel(file_path, skiprows=2)
        print("✅ Excel 数据文件读取成功！")

    except Exception as e:
        print(f"❌ 错误：读取文件 '{file_name}' 失败。错误信息: {e}。跳过此文件。")
        continue

    # --- 6. 站点识别与产品隔离 ---
    df.columns = df.columns.astype(str).str.strip()  # 清理原始列名
    a3_content = df.columns[0].strip()  # A3 单元格的内容就是第一个列名

    current_site, TARGET_CATEGORIES_FOR_FILE = identify_site_by_a3(a3_content)
    print(
        f"🌏 站点识别结果：A3='{a3_content}'，判定为 {current_site} 站点。将分析 {len(TARGET_CATEGORIES_FOR_FILE)} 种产品。")

    # --- 7. 关键兼容性处理：检查列名并重命名 ---
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
            missing_cols_std.append(std_col)

    if missing_cols_std:
        print(f"❌ 致命错误：文件 '{file_name}' 中缺少核心数据（标准列）：{missing_cols_std}。跳过此文件。")
        continue
    # --- 兼容性处理结束 ---

    # --- 8. 数据清洗与预处理 (略) ---
    cols_to_clean = [STD_CTR_COL, STD_CCR_COL]
    for col in cols_to_clean:
        df[col] = (df[col].astype(str).str.replace('%', '', regex=False).replace('', '0'))
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) / 100

    numeric_cols = [STD_ORDER_COL, STD_VV_COL]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)


    # 9. 定义映射函数 (使用所有产品的映射逻辑)
    def map_product(product_name):
        """根据商品名称映射到指定的产品分类 (包含所有站点的产品)"""
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
        elif '車のサンバイザーにス' in product_name_str:
            return '车载手机支架'
        elif '電働眉剃り器婦人眉剃' in product_name_str:
            return '电动修眉刀'
        else:
            return '其他/未匹配'


    # 应用映射函数
    df['产品分类'] = df[STD_ITEM_COL].apply(map_product)

    # --- 10. 分组计算与统计 ---
    df_analysis = df[df['产品分类'].isin(TARGET_CATEGORIES_FOR_FILE)].copy()

    if df_analysis.empty:
        print(
            f"⚠️ 警告：文件 '{file_name}' 属于 {current_site} 站点，但未找到任何目标产品 ({TARGET_CATEGORIES_FOR_FILE}) 的数据。跳过统计。")
        continue

    results = {}

    for category in TARGET_CATEGORIES_FOR_FILE:
        group = df_analysis[df_analysis['产品分类'] == category].copy()
        total_videos = len(group)

        # B. 出单视频子集
        videos_with_orders = group[group[STD_ORDER_COL] > 0].copy()
        count_with_orders = len(videos_with_orders)

        # C. 视频出单率
        order_rate = count_with_orders / total_videos if total_videos > 0 else 0

        # D. 所有视频的 VV 总和
        total_vv = group[STD_VV_COL].sum()

        # E. 仅出单视频的平均指标
        if count_with_orders > 0:
            average_ctr = videos_with_orders[STD_CTR_COL].mean()
            average_ccr = videos_with_orders[STD_CCR_COL].mean()
            average_vv = videos_with_orders[STD_VV_COL].mean()
        else:
            average_ctr = np.nan
            average_ccr = np.nan
            average_vv = np.nan

        # 存储结果到字典 (用于屏幕输出)
        results[category] = {
            '总视频数': total_videos,
            '出单视频数': count_with_orders,
            '视频VV总和': total_vv,
            '视频出单率': order_rate,
            '平均点击率': average_ctr,
            '平均点击成交率': average_ccr,
            '视频VV的平均值': average_vv,
        }

        # F. 收集数据到最终汇总列表
        final_summary_data.append({
            '文件名': file_name,
            '站点': current_site,
            '日期范围': standard_date_range,
            '产品分类': category,
            '总视频数': total_videos,
            '出单视频数': count_with_orders,
            '视频VV总和': total_vv,
            '视频出单率': order_rate,
            '平均点击率': average_ctr,
            '平均点击成交率': average_ccr,
            '视频VV的平均值': average_vv,
        })

    # 11. 打印当前文件结果 (屏幕输出)
    print("\n--- 关键数据分析结果 (屏幕输出) ---")
    for category, data in results.items():
        print(f"\n======== 【{category} ({current_site})】 ========")
        print(f"  总视频数: {data['总视频数']:.0f}")
        print(f"  出单视频数: {data['出单视频数']:.0f}")
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

# 12. 所有文件处理完毕后，生成最终报告文件
if final_summary_data:
    final_df = pd.DataFrame(final_summary_data)

    # --- 获取运行时间戳 ---
    current_time = datetime.now()
    # 格式化为 MM DD HH MM (例如 12091759)
    timestamp = current_time.strftime("%m%d%H%M")

    output_base_dir = '../result/'
    if not os.path.exists(output_base_dir):
        os.makedirs(output_base_dir, exist_ok=True)
        print(f"📁 目标输出目录 '{output_base_dir}' 已创建。")

    # 按 '站点' 和 '日期范围' 分组，生成独立文件
    grouped = final_df.groupby(['站点', '日期范围'])

    print("\n\n🎉 所有文件处理完成！开始生成独立报告...")

    for (site, date_range), group_df in grouped:
        # 构造文件名: [站点]-[日期范围]-[时间戳].xlsx
        # 例如: VN-2025-12-01-2025-12-07-12091759.xlsx
        safe_date_range = date_range.replace(' ~ ', '-').replace(' ', '_')

        # 完整的输出文件名
        output_file_name = os.path.join(output_base_dir, f"{site}-{safe_date_range}-{timestamp}.xlsx")

        try:
            group_df.to_excel(output_file_name, index=False, sheet_name=f"{site}汇总")
            print(f"📊 报告已成功保存至文件: {output_file_name}")
        except Exception as e:
            print(f"❌ 错误：保存报告 '{output_file_name}' 失败。错误信息: {e}")

else:
    print("\n🎉 所有文件处理完成！但未收集到有效数据。")