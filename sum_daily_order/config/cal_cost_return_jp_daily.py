# -*- coding: utf-8 -*-
"""
TikTok 日本按日汇总订单成本报表脚本
- 从飞书表格读取日本店 SKU 映射及成本配置
- 支持日本店铺（本土店、跨境店、直邮店）
- 使用固定日本汇率，不读取飞书汇率表
- 按 Paid Time 拆分自然日，输出按日汇总和按日累计订单成本报表
- 输出四张 Excel 报表（按日汇总、按日产品明细、销量矩阵、样品统计）
- 使用 .env 管理飞书凭证，支持本地缓存降级

【使用前必读】
1. 创建飞书应用并开启权限：sheet:client.doc.read（读取表格）
2. 将应用加入配置表的协作者（至少只读权限）
3. 配置 .env 文件中的 FEISHU_APP_ID 和 FEISHU_APP_SECRET
4. 确认订单 CSV 中的 SKU ID 列名，并修改下方 SKU_ID_COLUMN 变量
5. 在飞书表格中创建以下 Sheet：
   - "日本_本土店", "日本_跨境店", "日本_直邮店"（SKU配置）
"""

import os
import glob
import math
from datetime import datetime, timedelta
from urllib.parse import quote

import pandas as pd
import requests
from dotenv import load_dotenv

def find_project_root(start_dir):
    """向上查找 sum_daily_order 根目录，避免依赖固定绝对路径。"""
    env_root = os.getenv("TIKTOK_REPORT_ROOT")
    if env_root:
        root = os.path.abspath(env_root)
        for name in ("config", "data", "result"):
            os.makedirs(os.path.join(root, name), exist_ok=True)
        return root

    current = os.path.abspath(start_dir)
    while True:
        if all(os.path.isdir(os.path.join(current, name)) for name in ("config", "data", "result")):
            return current
        parent = os.path.dirname(current)
        if parent == current:
            raise FileNotFoundError("无法定位 sum_daily_order 根目录，请确认 config/data/result 目录仍在同一层级。")
        current = parent

# ============================================================
# 0. 加载环境变量（.env 文件）
# ============================================================
CURRENT_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = find_project_root(CURRENT_SCRIPT_DIR)

for env_path in [
    os.path.join(PROJECT_ROOT, "config", ".env"),
    os.path.join(CURRENT_SCRIPT_DIR, ".env"),
]:
    if os.path.exists(env_path):
        load_dotenv(env_path)
load_dotenv()

# -------- 飞书应用凭证（必须配置在 .env 中）--------
FEISHU_APP_ID = os.getenv("FEISHU_APP_ID")
FEISHU_APP_SECRET = os.getenv("FEISHU_APP_SECRET")

# -------- 飞书表格信息 --------
FEISHU_SHEET_TOKEN = "G3HCsMq7UhSjIptrTI5c0dUnnyh"
FEISHU_RANGE_SKU = "A:G"                  # SKU配置表的列范围（A到G）
FEISHU_REQUIRED_SCOPE_HINT = "请在飞书开放平台给应用开通 sheets:spreadsheet:readonly 或 sheets:spreadsheet:read 权限，并重新发布/生效。"

# -------- 日本固定汇率（1 日元兑人民币）--------
JAPAN_EXCHANGE_RATE = 0.042336

# -------- 订单 CSV 中 SKU ID 列名（请根据实际调整）--------
SKU_ID_COLUMN = "SKU ID"   # 订单 CSV 中的 SKU ID 列名
PRODUCT_CATEGORY_COLUMN = "产品大类"
LOGISTICS_CAPACITY_COLUMN = "每单物流承载数量"

# -------- 本地缓存目录 --------
CACHE_DIR = os.path.join(PROJECT_ROOT, "config", "cache")
os.makedirs(CACHE_DIR, exist_ok=True)
CACHE_EXPIRY_HOURS = 24   # 缓存有效期（小时）
USE_CONFIG_CACHE_FIRST = os.getenv("USE_CONFIG_CACHE_FIRST", "").lower() in {"1", "true", "yes", "y"}
SHEET_ID_CACHE = {}

# -------- 日本店铺配置 --------
# 如果某个店铺暂时不需要运行，把 enabled 改成 False 即可。
JAPAN_STORES = [
    {
        "enabled": True,
        "country_code": "jp",
        "country_name": "日本",
        "store_key": "local",
        "store_name": "本土店",
        "store_dir": "local",
        "sheet_name": "日本_本土店",
    },
    {
        "enabled": True,
        "country_code": "jp",
        "country_name": "日本",
        "store_key": "cross-border",
        "store_name": "跨境店",
        "store_dir": "cross-border",
        "sheet_name": "日本_跨境店",
    },
    {
        "enabled": True,
        "country_code": "jp",
        "country_name": "日本",
        "store_key": "direct",
        "store_name": "直邮店",
        "store_dir": "direct",
        "sheet_name": "日本_直邮店",
    },
]

# ============================================================
# 1. 飞书表格读取 + 缓存工具
# ============================================================

def get_feishu_token(app_id, app_secret):
    """获取飞书 tenant_access_token"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    payload = {"app_id": app_id, "app_secret": app_secret}
    try:
        resp = requests.post(url, json=payload, timeout=10)
        if resp.status_code != 200:
            raise Exception(f"飞书Token请求失败，HTTP {resp.status_code}")
        data = resp.json()
        if data.get("code") != 0:
            raise Exception(f"飞书Token错误: {data}")
        return data["tenant_access_token"]
    except Exception as e:
        raise Exception(f"获取飞书 token 失败: {e}")

def format_feishu_error(data):
    """把飞书 API 错误整理成控制台可读信息。"""
    if not isinstance(data, dict):
        return str(data)

    code = data.get("code")
    msg = data.get("msg")
    error = data.get("error") or {}
    message = error.get("message")
    log_id = error.get("log_id")

    parts = [f"code={code}", f"msg={msg}"]
    if message:
        parts.append(f"message={message}")
    if log_id:
        parts.append(f"log_id={log_id}")
    if code == 99991672:
        parts.append(FEISHU_REQUIRED_SCOPE_HINT)
    return "；".join(parts)

def get_feishu_json(url, headers, action):
    """请求飞书接口并保留详细错误，避免只看到 HTTP 400。"""
    resp = requests.get(url, headers=headers, timeout=10)
    try:
        data = resp.json()
    except Exception:
        data = {"code": resp.status_code, "msg": resp.text[:500]}

    if resp.status_code != 200 or data.get("code") != 0:
        raise Exception(f"{action}失败，HTTP {resp.status_code}，{format_feishu_error(data)}")
    return data

def get_feishu_sheet_id(token, sheet_token, sheet_name):
    """飞书 values API 需要 sheet_id，这里支持用工作表名称自动解析。"""
    cache_key = (sheet_token, sheet_name)
    if cache_key in SHEET_ID_CACHE:
        return SHEET_ID_CACHE[cache_key]

    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://open.feishu.cn/open-apis/sheets/v3/spreadsheets/{sheet_token}/sheets/query"
    data = get_feishu_json(url, headers, "查询飞书工作表列表")
    sheets = data.get("data", {}).get("sheets", [])

    available = []
    for item in sheets:
        props = item.get("sheet") or item.get("properties") or item
        title = props.get("title") or props.get("sheet_name")
        sheet_id = props.get("sheet_id") or props.get("sheetId")
        if title and sheet_id:
            SHEET_ID_CACHE[(sheet_token, title)] = sheet_id
            SHEET_ID_CACHE[(sheet_token, sheet_id)] = sheet_id
            available.append(f"{title}({sheet_id})")
            if title == sheet_name or sheet_id == sheet_name:
                return sheet_id

    raise ValueError(f"未找到飞书 Sheet '{sheet_name}'。当前可用 Sheet: {', '.join(available) or '无'}")

def read_feishu_sheet(token, sheet_token, sheet_name, range_str):
    """读取飞书表格指定 Sheet 的数据，返回 DataFrame（第一行为列名）"""
    headers = {"Authorization": f"Bearer {token}"}
    sheet_id = get_feishu_sheet_id(token, sheet_token, sheet_name)
    range_expr = quote(f"{sheet_id}!{range_str}", safe="")
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{sheet_token}/values/{range_expr}"
    try:
        data = get_feishu_json(url, headers, "读取飞书表格")
        values = data.get("data", {}).get("valueRange", {}).get("values", [])
        if not values:
            raise ValueError(f"飞书表格 Sheet '{sheet_name}' 返回空数据")
        df = pd.DataFrame(values[1:], columns=values[0])
        return df
    except Exception as e:
        raise Exception(f"读取飞书表格 Sheet '{sheet_name}' 失败: {e}")

def clean_sku_id_column(df):
    """
    清洗 DataFrame 中的 SKU ID 列（如果存在）：
    - 删除该列为空的行
    - 去除首尾空格
    - 将浮点型 '.0' 后缀移除
    """
    if 'SKU ID' not in df.columns:
        return df
    df = df.dropna(subset=['SKU ID'])
    df['SKU ID'] = df['SKU ID'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    df = df[df['SKU ID'] != '']
    return df

def clean_config_dataframe(df):
    """统一清洗配置表：SKU 转字符串，成本列转数字。"""
    df = clean_sku_id_column(df)
    cost_cols = ['产品成本(元)', '物流成本(元)', '寄样成本(元)']
    raw_sample_cost = df['寄样成本(元)'].copy() if '寄样成本(元)' in df.columns else None

    for col in cost_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    if '寄样成本(元)' not in df.columns and {'产品成本(元)', '物流成本(元)'}.issubset(df.columns):
        df['寄样成本(元)'] = df['产品成本(元)'] + df['物流成本(元)']
    if {'产品成本(元)', '物流成本(元)', '寄样成本(元)'}.issubset(df.columns):
        empty_sample_cost = df['寄样成本(元)'].isna() | (df['寄样成本(元)'] <= 0)
        df.loc[empty_sample_cost, '寄样成本(元)'] = (
            df.loc[empty_sample_cost, '产品成本(元)'] + df.loc[empty_sample_cost, '物流成本(元)']
        )
    if LOGISTICS_CAPACITY_COLUMN not in df.columns:
        df[LOGISTICS_CAPACITY_COLUMN] = 1
    df[LOGISTICS_CAPACITY_COLUMN] = pd.to_numeric(df[LOGISTICS_CAPACITY_COLUMN], errors='coerce').fillna(1)
    df.loc[df[LOGISTICS_CAPACITY_COLUMN] <= 0, LOGISTICS_CAPACITY_COLUMN] = 1

    if raw_sample_cost is not None:
        formula_mask = raw_sample_cost.astype(str).str.match(r'^=?C\d+\+D\d+$', na=False)
        if formula_mask.any() and {'产品成本(元)', '物流成本(元)', '寄样成本(元)'}.issubset(df.columns):
            df.loc[formula_mask, '寄样成本(元)'] = (
                df.loc[formula_mask, '产品成本(元)'] + df.loc[formula_mask, '物流成本(元)']
            )
    if PRODUCT_CATEGORY_COLUMN not in df.columns:
        df[PRODUCT_CATEGORY_COLUMN] = df.get('中文简称', '')
    for col in ['中文简称', PRODUCT_CATEGORY_COLUMN]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    if '中文简称' in df.columns:
        df[PRODUCT_CATEGORY_COLUMN] = df[PRODUCT_CATEGORY_COLUMN].replace({'': pd.NA, 'nan': pd.NA})
        df[PRODUCT_CATEGORY_COLUMN] = df[PRODUCT_CATEGORY_COLUMN].fillna(df['中文简称'])
    return df

def get_cache_file_path(sheet_name):
    """根据 Sheet 名称生成缓存文件路径"""
    safe_name = sheet_name.replace(' ', '_').replace('/', '_')
    return os.path.join(CACHE_DIR, f"config_{safe_name}.csv")

def get_config_dataframe(sheet_name, force_refresh=False):
    """
    获取指定 Sheet 的配置 DataFrame，优先飞书，失败则用本地缓存
    """
    cache_file = get_cache_file_path(sheet_name)
    cache_valid = False
    if os.path.exists(cache_file) and not force_refresh:
        mtime = datetime.fromtimestamp(os.path.getmtime(cache_file))
        if datetime.now() - mtime < timedelta(hours=CACHE_EXPIRY_HOURS):
            cache_valid = True

    if cache_valid and USE_CONFIG_CACHE_FIRST and not force_refresh:
        try:
            df = pd.read_csv(cache_file, dtype=str, encoding='utf-8-sig')
            df = clean_config_dataframe(df)
            print(f"   ✅ 使用本地缓存配置（Sheet: {sheet_name}）")
            return df
        except Exception as e:
            print(f"   ⚠️ 缓存读取失败: {e}，尝试刷新")

    try:
        if not FEISHU_APP_ID or not FEISHU_APP_SECRET:
            raise EnvironmentError(
                "未配置 FEISHU_APP_ID/FEISHU_APP_SECRET，无法从飞书刷新配置。"
            )
        token = get_feishu_token(FEISHU_APP_ID, FEISHU_APP_SECRET)
        df = read_feishu_sheet(token, FEISHU_SHEET_TOKEN, sheet_name, FEISHU_RANGE_SKU)
        df = df.dropna(how='all')
        df = clean_config_dataframe(df)
        df.to_csv(cache_file, index=False, encoding='utf-8-sig')
        print(f"   ✅ 成功从飞书拉取配置并缓存（Sheet: {sheet_name}）")
        return df
    except Exception as e:
        print(f"   ❌ 飞书拉取失败（Sheet: {sheet_name}）: {e}")
        if os.path.exists(cache_file):
            try:
                df = pd.read_csv(cache_file, dtype=str, encoding='utf-8-sig')
                df = clean_config_dataframe(df)
                print(f"   ⚠️ 飞书不可用，使用过期缓存（Sheet: {sheet_name}）")
                return df
            except:
                pass
        raise Exception(f"无法获取配置数据（Sheet: {sheet_name}）")

# ============================================================
# 2. 辅助函数
# ============================================================

def parse_amount(value):
    """将字符串金额转为浮点数，移除 'JPY', 逗号, 空格"""
    if pd.isna(value):
        return 0.0
    s = str(value).replace(' ', '').replace('JPY', '').replace(',', '')
    try:
        return float(s)
    except:
        return 0.0

def clean_sku_str(sku_series):
    """将 Series 中的 SKU 转换为字符串并去除 .0 后缀和空格"""
    return sku_series.astype(str).str.replace(r'\.0$', '', regex=True).str.strip()

def parse_paid_date(value):
    """从 Paid Time 解析自然日，解析失败返回 NA。"""
    if pd.isna(value):
        return pd.NA
    s = str(value).strip()
    if not s:
        return pd.NA
    try:
        return pd.to_datetime(s).strftime('%Y-%m-%d')
    except:
        return pd.NA

def safe_change_rate(current, previous):
    """计算环比变化率；上一期为 0 时返回空值，避免误导。"""
    if previous == 0:
        return ""
    return round((current - previous) / previous, 4)

def format_percent(value):
    """把小数格式化为百分比文本，便于 Excel 中直接阅读。"""
    if value == "" or pd.isna(value):
        return ""
    return f"{value:.2%}"

def build_order_logistics_allocation(df_normal, logistics_col):
    """按订单计算物流成本，并分摊到订单内的产品规格。

    规则：
    - 同一订单内，同一产品大类按“每单物流承载数量”计算候选物流成本。
    - 整单只取一次物流成本，即取各产品大类候选物流成本中的最大值。
    - 混合产品订单按候选物流成本比例分摊，保证整单物流总额不重复。
    """
    result_cols = ["日期", "Product Category", "Mapped Name", "物流成本"]
    if df_normal.empty:
        return pd.DataFrame(columns=result_cols)

    df_calc = df_normal.copy()
    if "Order ID" in df_calc.columns:
        order_key = df_calc["Order ID"].astype(str).str.strip()
        missing_order = order_key.isna() | order_key.eq("") | order_key.eq("nan")
        order_key = order_key.mask(missing_order, "ROW_" + df_calc.index.astype(str))
    else:
        order_key = "ROW_" + df_calc.index.astype(str)
    df_calc["_OrderKey"] = order_key
    df_calc[LOGISTICS_CAPACITY_COLUMN] = pd.to_numeric(
        df_calc.get(LOGISTICS_CAPACITY_COLUMN, 1),
        errors="coerce",
    ).fillna(1)
    df_calc.loc[df_calc[LOGISTICS_CAPACITY_COLUMN] <= 0, LOGISTICS_CAPACITY_COLUMN] = 1

    category_order = df_calc.groupby(["_OrderKey", "日期", "Product Category"]).agg(
        category_qty=("SKU_ID", "count"),
        original_logistics=(logistics_col, "sum"),
        unit_logistics=(logistics_col, "max"),
        capacity=(LOGISTICS_CAPACITY_COLUMN, "max"),
    ).reset_index()
    category_order["_capacity_logistics"] = category_order.apply(
        lambda row: math.ceil(row["category_qty"] / row["capacity"]) * row["unit_logistics"],
        axis=1,
    )
    category_order["candidate_logistics"] = category_order[["original_logistics", "_capacity_logistics"]].min(axis=1)
    order_max = category_order.groupby("_OrderKey")["candidate_logistics"].transform("max")
    order_sum = category_order.groupby("_OrderKey")["candidate_logistics"].transform("sum")
    category_order["category_logistics"] = category_order.apply(
        lambda row: 0 if order_sum.loc[row.name] == 0 else order_max.loc[row.name] * row["candidate_logistics"] / order_sum.loc[row.name],
        axis=1,
    )

    sku_order = df_calc.groupby(["_OrderKey", "日期", "Product Category", "Mapped Name"]).agg(
        sku_qty=("SKU_ID", "count"),
        sku_original_logistics=(logistics_col, "sum"),
    ).reset_index()
    sku_order = sku_order.merge(
        category_order[["_OrderKey", "日期", "Product Category", "category_qty", "original_logistics", "category_logistics"]],
        on=["_OrderKey", "日期", "Product Category"],
        how="left",
    )
    sku_order["物流成本"] = sku_order.apply(
        lambda row: (
            0
            if row["category_qty"] == 0
            else row["category_logistics"] * (
                row["sku_original_logistics"] / row["original_logistics"]
                if row["original_logistics"] != 0
                else row["sku_qty"] / row["category_qty"]
            )
        ),
        axis=1,
    )
    return sku_order.groupby(["日期", "Product Category", "Mapped Name"])["物流成本"].sum().reset_index()

def build_period_comparison_frames(df_daily_product, daily_order_records):
    """按 CSV 文件生成周期汇总；多文件时追加相邻周期对比。"""
    if df_daily_product.empty or "文件名" not in df_daily_product.columns:
        return []

    period_summary = df_daily_product.groupby("文件名").agg({
        "销量": "sum",
        "寄样数": "sum",
        "销售额": "sum",
        "汇率后金额": "sum",
        "P列折后价": "sum",
        "产品成本": "sum",
        "物流成本": "sum",
        "寄样支出": "sum",
        "日期": ["min", "max"],
    })
    period_summary.columns = [
        "销量", "寄样数", "销售额", "汇率后金额", "P列折后价",
        "产品成本", "物流成本", "寄样支出", "开始日期", "结束日期"
    ]

    if daily_order_records:
        df_order_count = (
            pd.DataFrame(daily_order_records)
            .drop_duplicates(subset=["文件名", "日期", "Order ID"])
            .groupby("文件名")["Order ID"]
            .count()
            .rename("订单数")
        )
        period_summary = period_summary.join(df_order_count, how="left")
    else:
        period_summary["订单数"] = period_summary["销量"]

    period_summary["订单数"] = period_summary["订单数"].fillna(0)
    period_summary["利润"] = (
        period_summary["汇率后金额"]
        - period_summary["产品成本"]
        - period_summary["物流成本"]
        - period_summary["寄样支出"]
    )
    period_summary["利润率"] = period_summary.apply(
        lambda row: 0 if row["汇率后金额"] == 0 else row["利润"] / row["汇率后金额"],
        axis=1,
    )
    period_summary = period_summary.reset_index()
    period_summary["_开始日期排序"] = pd.to_datetime(period_summary["开始日期"], errors="coerce")
    period_summary["_结束日期排序"] = pd.to_datetime(period_summary["结束日期"], errors="coerce")
    period_summary = (
        period_summary
        .sort_values(["_开始日期排序", "_结束日期排序", "文件名"])
        .drop(columns=["_开始日期排序", "_结束日期排序"])
        .set_index("文件名")
    )

    metrics = ["订单数", "销量", "寄样数", "销售额", "汇率后金额", "P列折后价", "产品成本", "物流成本", "寄样支出", "利润", "利润率"]
    frames = []
    ordered_files = list(period_summary.index)

    summary_rows = []
    for file_name in ordered_files:
        row = {"文件名": file_name}
        for metric in metrics:
            if metric == "利润率":
                row[metric] = format_percent(period_summary.loc[file_name, metric])
            else:
                row[metric] = round(period_summary.loc[file_name, metric], 2)
        summary_rows.append(row)
    frames.append(("周期汇总", pd.DataFrame(summary_rows)))

    for idx in range(1, len(ordered_files)):
        previous_file = ordered_files[idx - 1]
        current_file = ordered_files[idx]
        previous = period_summary.loc[previous_file]
        current = period_summary.loc[current_file]
        change_value_row = {"文件名": f"变化值：{current_file} - {previous_file}"}
        change_rate_row = {"文件名": f"变化率：{current_file} / {previous_file} - 1"}
        for metric in metrics:
            if metric == "利润率":
                change_value_row[metric] = format_percent(current[metric] - previous[metric])
                change_rate_row[metric] = ""
            else:
                previous_value = round(previous[metric], 2)
                current_value = round(current[metric], 2)
                change_value_row[metric] = round(current_value - previous_value, 2)
                change_rate_row[metric] = format_percent(safe_change_rate(current_value, previous_value))
        title = f"{current_file} - {previous_file}"
        frames.append((title, pd.DataFrame([change_value_row, change_rate_row])))
    return frames

def center_excel_sheet(writer, sheet_name, row_count, col_count):
    """把已写入的工作表区域设置为居中显示。"""
    worksheet = writer.sheets[sheet_name]
    if not hasattr(writer, "_center_format"):
        writer._center_format = writer.book.add_format({"align": "center", "valign": "vcenter"})
    center_format = writer._center_format
    worksheet.set_column(0, max(col_count - 1, 0), 14, center_format)
    for row_idx in range(row_count):
        worksheet.set_row(row_idx, None, center_format)

def insert_blank_rows_between_files(df):
    """输出展示用：不同文件名之间插入空行，不参与任何计算。"""
    if df.empty or "文件名" not in df.columns:
        return df

    parts = []
    blank_row = {col: "" for col in df.columns}
    for _, group in df.groupby("文件名", sort=False, dropna=False):
        if parts:
            parts.append(pd.DataFrame([blank_row], columns=df.columns))
        parts.append(group)
    return pd.concat(parts, ignore_index=True)

def build_product_profit_by_period(df_daily_product):
    """按文件+产品汇总 GMV、成本和利润，方便判断单品盈利情况。"""
    if df_daily_product.empty:
        return pd.DataFrame([["无产品汇总数据"]], columns=["提示"])

    summary = df_daily_product.groupby(["文件名", "产品名称"]).agg({
        "销量": "sum",
        "寄样数": "sum",
        "销售额": "sum",
        "汇率后金额": "sum",
        "P列折后价": "sum",
        "产品成本": "sum",
        "物流成本": "sum",
        "寄样支出": "sum",
    }).reset_index()
    summary["总成本"] = summary["产品成本"] + summary["物流成本"] + summary["寄样支出"]
    summary["利润"] = summary["汇率后金额"] - summary["总成本"]
    summary["利润率"] = summary.apply(
        lambda row: "" if row["汇率后金额"] == 0 else format_percent(row["利润"] / row["汇率后金额"]),
        axis=1,
    )

    numeric_cols = ["销售额", "汇率后金额", "P列折后价", "产品成本", "物流成本", "寄样支出", "总成本", "利润"]
    for col in numeric_cols:
        summary[col] = summary[col].round(2)
    return summary[[
        "文件名", "产品名称", "销量", "寄样数", "销售额", "汇率后金额", "P列折后价",
        "产品成本", "物流成本", "寄样支出", "总成本", "利润", "利润率"
    ]]

def build_product_quantity_by_period(product_quantity_records):
    """按文件汇总每个产品的销量，放在每日销量矩阵上方。"""
    df_quantity_records = pd.DataFrame(product_quantity_records)
    if df_quantity_records.empty:
        return pd.DataFrame([["无销量数据"]], columns=["提示"])

    matrix = df_quantity_records.pivot_table(
        index="产品名称",
        columns="文件名",
        values="销量",
        aggfunc="sum",
        fill_value=0,
    ).astype(int)
    matrix["汇总"] = matrix.sum(axis=1)
    matrix.loc["汇总"] = matrix.sum(axis=0)
    return matrix

# ============================================================
# 3. 核心处理函数
# ============================================================

def run_report(combo, config_df, exchange_rate):
    """
    处理单个国家+店铺组合
    :param combo: 包含 country_code, country_name, store_key, store_name, sheet_name 的字典
    :param config_df: 该组合的配置 DataFrame
    :param exchange_rate: 该国家的汇率（1单位外币兑人民币）
    """
    country_code = combo["country_code"]
    country_name = combo["country_name"]
    store_key = combo["store_key"]
    store_name = combo["store_name"]
    store_dir = combo["store_dir"]
    display_name = f"{country_name}_{store_name}"

    # 配置表中的列名（统一）
    logistics_col = "物流成本(元)"
    sample_cost_col = "寄样成本(元)"

    folder_path = os.path.normpath(os.path.join(PROJECT_ROOT, 'data', f'data_{country_code.upper()}', store_dir))
    output_filename = os.path.normpath(
        os.path.join(PROJECT_ROOT, f'result/Daily_Performance_Report_{country_code.upper()}_{store_key}.xlsx'))

    file_list = glob.glob(os.path.join(folder_path, '*.csv'))
    if not file_list:
        print(f"⚠️  [{display_name}] 路径下未找到文件: {folder_path}")
        return

    print(f"🚀 开始处理 [{display_name}]，共 {len(file_list)} 个文件")

    # 订单CSV关键列名
    status_col = 'Order Status'
    quantity_col = 'Normal or Pre-order'
    n_col = 'SKU Platform Discount'
    p_col = 'SKU Subtotal After Discount'
    r_col = 'Original Shipping Fee'

    daily_product_detail_list = []
    product_quantity_records = []
    sample_quantity_records = []
    daily_order_records = []
    unmatched_skus = set()

    for file_path in file_list:
        file_name = os.path.basename(file_path)
        try:
            df = pd.read_csv(file_path)
        except Exception as e:
            print(f"读取失败 {file_name}: {e}")
            continue

        cancel_keywords = ['Canceled', 'Unpaid', 'キャンセル済み', '未払い', '已取消', '未支付', '部分发货后取消']

        # -------- 正常订单 --------
        df_normal = df[~df[status_col].isin(cancel_keywords) & (df[quantity_col] == 'Normal')].copy()
        if 'Paid Time' not in df_normal.columns:
            print(f"⚠️ 文件 {file_name} 中未找到 Paid Time 列，将跳过此文件")
            continue
        df_normal['日期'] = df_normal['Paid Time'].map(parse_paid_date)
        df_normal = df_normal.dropna(subset=['日期']).copy()

        for col in [n_col, p_col, r_col]:
            df_normal[col] = df_normal[col].map(parse_amount)

        if SKU_ID_COLUMN not in df_normal.columns:
            print(f"⚠️ 文件 {file_name} 中未找到 SKU ID 列 '{SKU_ID_COLUMN}'，将跳过此文件")
            continue
        df_normal['SKU_ID'] = clean_sku_str(df_normal[SKU_ID_COLUMN])

        # 左连接配置表
        df_normal = df_normal.merge(
            config_df[['SKU ID', PRODUCT_CATEGORY_COLUMN, '中文简称', '产品成本(元)', logistics_col, sample_cost_col, LOGISTICS_CAPACITY_COLUMN]],
            left_on='SKU_ID',
            right_on='SKU ID',
            how='left'
        )

        missing = df_normal[df_normal['中文简称'].isna()]
        if not missing.empty:
            for sku in missing['SKU_ID'].unique():
                unmatched_skus.add(sku)

        df_normal['Mapped Name'] = df_normal['中文简称'].fillna(df_normal['SKU_ID'])
        df_normal['Product Category'] = df_normal[PRODUCT_CATEGORY_COLUMN].fillna(df_normal['Mapped Name'])
        if 'Order ID' in df_normal.columns:
            daily_order_records.extend(
                df_normal[['日期', 'Order ID']]
                .dropna()
                .drop_duplicates()
                .assign(文件名=file_name)
                .to_dict('records')
            )

        norm_agg = df_normal.groupby('Mapped Name').agg({
            n_col: 'sum',
            p_col: 'sum',
            r_col: 'sum',
            '产品成本(元)': 'first',
            logistics_col: 'first',
            sample_cost_col: 'first',
            'SKU_ID': 'count'
        }).reset_index()
        norm_agg.rename(columns={'SKU_ID': '销量'}, inplace=True)

        # -------- 样品订单 --------
        df_sample = df[
            ~df[status_col].isin(cancel_keywords) & (df[quantity_col].isna() | (df[quantity_col] == ''))
        ].copy()
        if not df_sample.empty:
            if 'Paid Time' not in df_sample.columns:
                print(f"⚠️ 文件 {file_name} 样品订单中未找到 Paid Time 列，将跳过样品统计")
                df_sample = pd.DataFrame()
            else:
                df_sample['日期'] = df_sample['Paid Time'].map(parse_paid_date)
                df_sample = df_sample.dropna(subset=['日期']).copy()

        if not df_sample.empty:
            df_sample['SKU_ID'] = clean_sku_str(df_sample[SKU_ID_COLUMN])
            df_sample = df_sample.merge(
                config_df[['SKU ID', PRODUCT_CATEGORY_COLUMN, '中文简称', sample_cost_col]],
                left_on='SKU_ID',
                right_on='SKU ID',
                how='left'
            )
            missing_s = df_sample[df_sample['中文简称'].isna()]
            if not missing_s.empty:
                for sku in missing_s['SKU_ID'].unique():
                    unmatched_skus.add(sku)
            df_sample['Mapped Name'] = df_sample['中文简称'].fillna(df_sample['SKU_ID'])
            df_sample['Product Category'] = df_sample[PRODUCT_CATEGORY_COLUMN].fillna(df_sample['Mapped Name'])

        # -------- 按日期 + 产品汇总 --------
        normal_daily_product = df_normal.groupby(['日期', 'Product Category', 'Mapped Name']).agg({
            n_col: 'sum',
            p_col: 'sum',
            r_col: 'sum',
            '产品成本(元)': 'first',
            'SKU_ID': 'count'
        }).reset_index()
        normal_daily_product.rename(columns={'SKU_ID': '销量'}, inplace=True)
        normal_daily_product['销售额'] = (
            normal_daily_product[n_col] + normal_daily_product[p_col] + normal_daily_product[r_col]
        )
        normal_daily_product['产品成本'] = normal_daily_product['销量'] * normal_daily_product['产品成本(元)']
        logistics_daily_product = build_order_logistics_allocation(df_normal, logistics_col)
        normal_daily_product = normal_daily_product.merge(
            logistics_daily_product,
            on=['日期', 'Product Category', 'Mapped Name'],
            how='left'
        )
        normal_daily_product['物流成本'] = normal_daily_product['物流成本'].fillna(0)

        sample_daily_product = pd.DataFrame(columns=['日期', 'Product Category', 'Mapped Name', '寄样数', '寄样支出'])
        if not df_sample.empty:
            sample_daily_product = df_sample.groupby(['日期', 'Product Category', 'Mapped Name']).agg({
                sample_cost_col: 'first',
                'SKU_ID': 'count'
            }).reset_index()
            sample_daily_product.rename(columns={'SKU_ID': '寄样数'}, inplace=True)
            sample_daily_product['寄样支出'] = sample_daily_product['寄样数'] * sample_daily_product[sample_cost_col]
            sample_daily_product = sample_daily_product[['日期', 'Product Category', 'Mapped Name', '寄样数', '寄样支出']]

        merged_daily_product = normal_daily_product[[
            '日期', 'Product Category', 'Mapped Name', '销量', '销售额', p_col, '产品成本', '物流成本'
        ]].merge(
            sample_daily_product,
            on=['日期', 'Product Category', 'Mapped Name'],
            how='outer'
        )
        for col in ['销量', '销售额', p_col, '产品成本', '物流成本', '寄样数', '寄样支出']:
            if col in merged_daily_product.columns:
                merged_daily_product[col] = pd.to_numeric(merged_daily_product[col], errors='coerce').fillna(0)
        merged_daily_product['汇率后金额'] = (merged_daily_product['销售额'] * exchange_rate).round(2)

        for _, row in merged_daily_product.iterrows():
            daily_product_detail_list.append({
                '文件名': file_name,
                '日期': row['日期'],
                '产品大类': row['Product Category'],
                '产品名称': row['Mapped Name'],
                '销量': int(row['销量']),
                '销售额': round(row['销售额'], 2),
                '汇率后金额': round(row['汇率后金额'], 2),
                'P列折后价': round(row[p_col], 2),
                '产品成本': round(row['产品成本'], 2),
                '物流成本': round(row['物流成本'], 2),
                '寄样数': int(row['寄样数']),
                '寄样支出': round(row['寄样支出'], 2)
            })

        for _, row in normal_daily_product.iterrows():
            product_quantity_records.append({
                '文件名': file_name,
                '日期': row['日期'],
                '产品名称': row['Product Category'],
                '销量': int(row['销量'])
            })

        for _, row in sample_daily_product.iterrows():
            sample_quantity_records.append({
                '文件名': file_name,
                '日期': row['日期'],
                '产品名称': row['Product Category'],
                '寄样数': int(row['寄样数'])
            })

    # -------- 输出 Excel --------
    if not daily_product_detail_list:
        print(f"⚠️ [{display_name}] 没有有效数据，不生成报表。")
        return

    # Sheet1: Daily Summary
    df_sku_detail = pd.DataFrame(daily_product_detail_list)
    df_daily_product = df_sku_detail.groupby(['文件名', '日期', '产品大类']).agg({
        '销量': 'sum',
        '销售额': 'sum',
        '汇率后金额': 'sum',
        'P列折后价': 'sum',
        '产品成本': 'sum',
        '物流成本': 'sum',
        '寄样数': 'sum',
        '寄样支出': 'sum'
    }).reset_index().rename(columns={'产品大类': '产品名称'})
    df_daily_summary = df_daily_product.groupby(['文件名', '日期']).agg({
        '销量': 'sum',
        '销售额': 'sum',
        '汇率后金额': 'sum',
        'P列折后价': 'sum',
        '产品成本': 'sum',
        '物流成本': 'sum',
        '寄样数': 'sum',
        '寄样支出': 'sum'
    }).reset_index()
    if daily_order_records:
        df_order_count = pd.DataFrame(daily_order_records) \
            .drop_duplicates(subset=['文件名', '日期', 'Order ID']) \
            .groupby(['文件名', '日期'])['Order ID'].count().reset_index()
        df_order_count.rename(columns={'Order ID': '订单数'}, inplace=True)
    else:
        df_order_count = df_daily_product[df_daily_product['销量'] > 0].groupby(['文件名', '日期'])['销量'].sum().reset_index()
        df_order_count.rename(columns={'销量': '订单数'}, inplace=True)
    df_daily_summary = df_daily_summary.merge(df_order_count, on=['文件名', '日期'], how='left')
    df_daily_summary['订单数'] = df_daily_summary['订单数'].fillna(0).astype(int)
    df_daily_summary = df_daily_summary[[
        '文件名', '日期', '订单数', '销量', '寄样数', '销售额', '汇率后金额', 'P列折后价', '产品成本', '物流成本', '寄样支出'
    ]]
    df_daily_summary = df_daily_summary.assign(temp_date=pd.to_datetime(df_daily_summary['日期'])) \
        .sort_values(['文件名', 'temp_date']) \
        .drop(columns=['temp_date'])
    total_row = {'文件名': '全部文件合计', '日期': '汇总'}
    for col in df_daily_summary.columns:
        if col not in ['文件名', '日期']:
            total_row[col] = df_daily_summary[col].sum()
    df_daily_summary = pd.concat([df_daily_summary, pd.DataFrame([total_row])], ignore_index=True)

    # Sheet2: Daily Product Detail
    df_daily_product = df_daily_product.assign(temp_date=pd.to_datetime(df_daily_product['日期'])) \
        .sort_values(['temp_date', '文件名', '产品名称']) \
        .drop(columns=['temp_date'])
    df_daily_product = df_daily_product[[
        '文件名', '日期', '产品名称', '销量', '寄样数', '销售额', '汇率后金额', 'P列折后价', '产品成本', '物流成本', '寄样支出'
    ]]
    df_sku_detail = df_sku_detail.assign(temp_date=pd.to_datetime(df_sku_detail['日期'])) \
        .sort_values(['temp_date', '文件名', '产品大类', '产品名称']) \
        .drop(columns=['temp_date'])
    df_sku_detail = df_sku_detail[[
        '文件名', '日期', '产品大类', '产品名称', '销量', '寄样数', '销售额', '汇率后金额', 'P列折后价', '产品成本', '物流成本', '寄样支出'
    ]]

    period_comparison_frames = build_period_comparison_frames(df_daily_product, daily_order_records)
    df_product_profit_by_period = build_product_profit_by_period(df_daily_product)
    df_product_quantity_by_period = build_product_quantity_by_period(product_quantity_records)

    # Sheet3: Product Quantity Matrix
    df_quantity_records = pd.DataFrame(product_quantity_records)
    if not df_quantity_records.empty:
        df_quantity_matrix = df_quantity_records.pivot_table(
            index=['文件名', '产品名称'],
            columns='日期',
            values='销量',
            aggfunc='sum',
            fill_value=0
        ).astype(int)
        sorted_cols = sorted(df_quantity_matrix.columns, key=lambda x: pd.to_datetime(x))
        df_quantity_matrix = df_quantity_matrix[sorted_cols]
        df_quantity_matrix['汇总'] = df_quantity_matrix.sum(axis=1)
        total_row = df_quantity_matrix.sum(axis=0).to_frame().T
        total_row.index = pd.MultiIndex.from_tuples([('全部文件合计', '汇总')], names=df_quantity_matrix.index.names)
        df_quantity_matrix = pd.concat([df_quantity_matrix, total_row]).reset_index()
    else:
        df_quantity_matrix = pd.DataFrame([["无销量数据"]], columns=["提示"])

    # Sheet4: Sample Statistics
    df_sample_records = pd.DataFrame(sample_quantity_records)
    if not df_sample_records.empty:
        df_sample_matrix = df_sample_records.pivot_table(
            index=['文件名', '产品名称'],
            columns='日期',
            values='寄样数',
            aggfunc='sum',
            fill_value=0
        ).astype(int)
        sorted_cols = sorted(df_sample_matrix.columns, key=lambda x: pd.to_datetime(x))
        df_sample_matrix = df_sample_matrix[sorted_cols]
        df_sample_matrix['汇总'] = df_sample_matrix.sum(axis=1)
        total_row = df_sample_matrix.sum(axis=0).to_frame().T
        total_row.index = pd.MultiIndex.from_tuples([('全部文件合计', '汇总')], names=df_sample_matrix.index.names)
        df_sample_matrix = pd.concat([df_sample_matrix, total_row]).reset_index()
    else:
        df_sample_matrix = pd.DataFrame([["无样品数据"]], columns=["提示"])

    # 保存
    try:
        if not os.path.exists(os.path.dirname(output_filename)):
            os.makedirs(os.path.dirname(output_filename))
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            df_daily_summary_display = insert_blank_rows_between_files(df_daily_summary)
            df_daily_summary_display.to_excel(writer, sheet_name='Daily Summary', index=False)
            center_excel_sheet(writer, 'Daily Summary', len(df_daily_summary_display) + 1, len(df_daily_summary_display.columns))

            detail_sheet = 'Daily Product Detail'
            df_product_profit_display = insert_blank_rows_between_files(df_product_profit_by_period)
            df_daily_product_display = insert_blank_rows_between_files(df_daily_product)
            df_product_profit_display.to_excel(writer, sheet_name=detail_sheet, index=False)
            detail_startrow = len(df_product_profit_display) + 3
            df_daily_product_display.to_excel(writer, sheet_name=detail_sheet, startrow=detail_startrow, index=False)
            detail_rows = detail_startrow + len(df_daily_product_display) + 1
            detail_cols = max(len(df_product_profit_display.columns), len(df_daily_product_display.columns))
            center_excel_sheet(writer, detail_sheet, detail_rows, detail_cols)

            sku_detail_sheet = 'SKU Detail'
            df_sku_detail_display = insert_blank_rows_between_files(df_sku_detail)
            df_sku_detail_display.to_excel(writer, sheet_name=sku_detail_sheet, index=False)
            center_excel_sheet(writer, sku_detail_sheet, len(df_sku_detail_display) + 1, len(df_sku_detail_display.columns))

            quantity_sheet = 'Product Quantity Matrix'
            df_product_quantity_by_period.to_excel(writer, sheet_name=quantity_sheet, index=True)
            quantity_startrow = len(df_product_quantity_by_period) + 3
            df_quantity_matrix_display = insert_blank_rows_between_files(df_quantity_matrix)
            df_quantity_matrix_display.to_excel(writer, sheet_name=quantity_sheet, startrow=quantity_startrow, index=False)
            quantity_rows = quantity_startrow + len(df_quantity_matrix_display) + 1
            quantity_cols = max(len(df_product_quantity_by_period.columns) + 1, len(df_quantity_matrix_display.columns))
            center_excel_sheet(writer, quantity_sheet, quantity_rows, quantity_cols)

            df_sample_matrix_display = insert_blank_rows_between_files(df_sample_matrix)
            df_sample_matrix_display.to_excel(writer, sheet_name='Sample Statistics', index=False)
            center_excel_sheet(writer, 'Sample Statistics', len(df_sample_matrix_display) + 1, len(df_sample_matrix_display.columns))

            if period_comparison_frames:
                startrow = 0
                sheet_name = 'Period Comparison'
                max_cols = 0
                for title, frame in period_comparison_frames:
                    frame.to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False)
                    max_cols = max(max_cols, len(frame.columns))
                    startrow += len(frame) + 3
                center_excel_sheet(writer, sheet_name, startrow, max_cols)
            else:
                period_hint = pd.DataFrame([{'提示': '周期对比需要至少 2 个有效 CSV 文件'}])
                period_hint.to_excel(writer, sheet_name='Period Comparison', index=False)
                center_excel_sheet(writer, 'Period Comparison', len(period_hint) + 1, len(period_hint.columns))
        print(f"✅ [{display_name}] 处理完成，报表已保存至 {output_filename}")

        if unmatched_skus:
            print(f"💡 [{display_name}] 提醒：以下 SKU ID 在配置表中未找到匹配项:")
            for sku in sorted(list(unmatched_skus)):
                print(f"   - {sku}")
    except Exception as e:
        print(f"❌ [{display_name}] 保存失败: {e}")

def run_japan_daily(stores=None):
    """运行日本日报。stores 为空时使用脚本内置 JAPAN_STORES。"""
    print("=" * 60)
    print("TikTok 日本按日汇总订单成本报表工具")
    print("=" * 60)

    exchange_rate = JAPAN_EXCHANGE_RATE
    print(f"✅ 使用固定日本汇率: {exchange_rate}")

    # 遍历日本店铺组合
    for combo in (stores or JAPAN_STORES):
        if not combo.get("enabled", True):
            continue

        sheet_name = combo["sheet_name"]
        display_name = f"{combo['country_name']}_{combo['store_name']}"

        print(f"\n📋 处理组合: {display_name} (汇率: {exchange_rate})")

        try:
            config_df = get_config_dataframe(sheet_name)
            print(f"   ✅ 配置表加载成功，共 {len(config_df)} 条记录")
            run_report(combo, config_df, exchange_rate)
        except Exception as e:
            print(f"   ❌ 处理失败: {e}")
            continue

    print("\n" + "=" * 60)
    print("✨ 所有任务执行完毕。")
    print("=" * 60)

# ============================================================
# 4. 主程序
# ============================================================

if __name__ == "__main__":
    run_japan_daily()
