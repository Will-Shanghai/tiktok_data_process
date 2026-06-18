from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter


# ========= 1. 路径配置 =========
# 适配你的项目结构：
# cal_ads_data/
#   config/
#     ads_statistics.py
#   data/
#     0521-ads.xlsx
#     0522-ads.xlsx
#     0523-ads.xlsx
#   result/
#
# 无论在 PyCharm 里从哪里运行，都会基于当前脚本位置自动定位项目目录。
CONFIG_DIR = Path(__file__).resolve().parent
PROJECT_DIR = CONFIG_DIR.parent
DATA_DIR = PROJECT_DIR / "data"
RESULT_DIR = PROJECT_DIR / "result"
OUTPUT_PATH = RESULT_DIR / "ads_statistics_result.xlsx"


# ========= 2. 字段配置 =========
# 当前附件中，“商品卡片”是“创意作品类型”列中的一个取值，取值示例为：商品卡片、视频。
PRODUCT_CARD_COL = "创意作品类型"
PRODUCT_CARD_VALUES = {"商品卡片", "商品卡"}

ORDER_COL = "SKU 订单数"
REVENUE_COL = "总收入"
CTR_COL = "商品广告点击率"
CVR_COL = "广告转化率"
VIDEO_2S_RATE_COL = "广告视频播放达 2 秒播放率"

REQUIRED_COLUMNS = [
    PRODUCT_CARD_COL,
    ORDER_COL,
    REVENUE_COL,
    CTR_COL,
    CVR_COL,
    VIDEO_2S_RATE_COL,
]

NUMERIC_COLUMNS = [
    ORDER_COL,
    REVENUE_COL,
    CTR_COL,
    CVR_COL,
    VIDEO_2S_RATE_COL,
]

RESULT_COLUMNS = [
    "文件名",
    "总 SKU 订单数",
    "总收入",
    "广告点击率平均值",
    "广告转化率平均值",
    "广告视频 2 秒播放率平均值",
    "商品卡订单数",
    "商品卡总收入",
    "商品卡广告点击率平均值"
]

PERCENT_COLUMNS = [
    "广告点击率平均值",
    "广告转化率平均值",
    "广告视频 2 秒播放率平均值",
    "商品卡广告点击率平均值",
]


def get_excel_files(data_dir):
    """获取 data 目录下所有 xlsx 文件，并跳过 Excel 临时文件。"""
    if not data_dir.exists():
        raise FileNotFoundError(f"未找到 data 目录：{data_dir}")

    excel_files = sorted(
        file
        for file in data_dir.glob("*.xlsx")
        if not file.name.startswith(("~$", ".~"))
    )

    if not excel_files:
        raise FileNotFoundError(f"data 目录下没有可读取的 .xlsx 文件：{data_dir}")

    return excel_files


def parse_number(value):
    """把数字、百分比字符串、'-' 等统一转换成可计算的数值。"""
    if pd.isna(value):
        return np.nan

    if isinstance(value, str):
        value = value.strip().replace(",", "")
        if value in {"", "-", "--", "—"}:
            return np.nan

        # 兼容类似 "2.07%" 的百分比文本；转换后为 0.0207。
        if value.endswith("%"):
            return float(value[:-1]) / 100

    return float(value)


def check_required_columns(df, required_columns, file_path):
    """检查必要字段是否存在，避免列名变化时静默算错。"""
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(
            f"文件 {file_path.name} 缺少必要字段："
            + "、".join(missing)
            + "\n当前实际字段为："
            + "、".join(map(str, df.columns))
        )


def clean_numeric_columns(df, columns):
    """批量清洗需要参与计算的数值列。"""
    for col in columns:
        df[col] = df[col].map(parse_number).astype("float64")
    return df


def read_one_excel(file_path):
    """读取单个 Excel 文件，并完成字段检查、数值清洗。"""
    df = pd.read_excel(file_path, sheet_name=0, engine="openpyxl")
    check_required_columns(df, REQUIRED_COLUMNS, file_path)
    df = clean_numeric_columns(df, NUMERIC_COLUMNS)
    return df


def build_statistics_row(file_path, df):
    """基于单个文件生成一行统计结果。"""
    product_card_mask = (
        df[PRODUCT_CARD_COL]
        .astype("string")
        .str.strip()
        .isin(PRODUCT_CARD_VALUES)
    )
    product_card_df = df.loc[product_card_mask]

    return {
        "文件名": file_path.name,
        # 所有 SKU 订单数汇总
        "总 SKU 订单数": df[ORDER_COL].sum(skipna=True),
        # 所有收入汇总
        "总收入": df[REVENUE_COL].sum(skipna=True),
        # 整体平均值
        "广告点击率平均值": df[CTR_COL].mean(skipna=True),
        "广告转化率平均值": df[CVR_COL].mean(skipna=True),
        "广告视频 2 秒播放率平均值": df[VIDEO_2S_RATE_COL].mean(skipna=True),
        # 仅商品卡数据
        "商品卡订单数": product_card_df[ORDER_COL].sum(skipna=True),
        "商品卡总收入": product_card_df[REVENUE_COL].sum(skipna=True),
        "商品卡广告点击率平均值": product_card_df[CTR_COL].mean(skipna=True),
    }


def format_for_console(df):
    """仅用于控制台展示，把比例列转成百分比字符串，便于阅读。"""
    result = df.copy()
    rate_columns = [
        "广告点击率平均值",
        "广告转化率平均值",
        "广告视频 2 秒播放率平均值",
        "商品卡广告点击率平均值",
    ]

    for col in dict.fromkeys(rate_columns):
        if col in result.columns:
            result[col] = result[col].map(lambda x: "" if pd.isna(x) else f"{x:.2%}")

    return result


def apply_excel_percentage_format(writer, sheet_name, df):
    """把结果表中的比例列设置为 Excel 百分比格式。"""
    worksheet = writer.sheets[sheet_name]

    for col_name in PERCENT_COLUMNS:
        if col_name not in df.columns:
            continue

        col_index = df.columns.get_loc(col_name) + 1
        col_letter = get_column_letter(col_index)

        for row_index in range(2, len(df) + 2):
            worksheet[f"{col_letter}{row_index}"].number_format = "0.00%"


def main():
    # ========= 3. 读取 data 目录下所有 Excel =========
    excel_files = get_excel_files(DATA_DIR)
    print("本次读取文件：")
    for file in excel_files:
        print(f"- {file.name}")

    # ========= 4. 逐个文件生成统计结果 =========
    result_rows = []
    for file in excel_files:
        df = read_one_excel(file)
        result_rows.append(build_statistics_row(file, df))

    result_df = pd.DataFrame(result_rows, columns=RESULT_COLUMNS)

    # ========= 5. 控制台输出 =========
    print("\n========== 统计结果 ==========")
    print(format_for_console(result_df).to_string(index=False))

    # ========= 6. 保存到 result 目录 =========
    RESULT_DIR.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        sheet_name = "统计结果"
        result_df.to_excel(writer, sheet_name=sheet_name, index=False)
        apply_excel_percentage_format(writer, sheet_name, result_df)

    print(f"\n统计结果已保存到：{OUTPUT_PATH}")


if __name__ == "__main__":
    main()
