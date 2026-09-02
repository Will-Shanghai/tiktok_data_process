#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Build Mexico SKU mapping table from:
1. 墨西哥直邮店主要产品成本核算.xlsx
2. 0701-0829.xlsx

Output columns:
产品大类 | SKU ID | SKU中文简称 | 产品成本(元) | 长cm | 宽cm | 高cm | 计费重kg
"""

from __future__ import annotations

from pathlib import Path
import re

import pandas as pd


ROOT = Path("/Users/macbook/Documents/Projects/tiktok_data_process")
ORDER_FILE = Path("/Users/macbook/Downloads/0701-0829.xlsx")
COST_FILE = Path("/Users/macbook/Downloads/墨西哥直邮店主要产品成本核算.xlsx")
OUTPUT_FILE = ROOT / "outputs" / "墨西哥直邮SKU映射表.xlsx"
MX_COST_EXCHANGE_RATE = 0.3819


VARIANT_MAP = {
    "Black": "黑色",
    "White": "白色",
    "hite": "白色",
    "Green": "绿色",
    "Gray": "浅灰",
    "Purple": "紫色",
    "Orange": "橙色",
    "Amber": "琥珀",
    "Nude": "米白",
    "Coffee": "咖啡",
    "BlackCoffee": "黑咖啡",
    "GrayCoffee": "浅灰咖啡",
    "Ham": "汉堡",
    "Fries": "薯条",
    "Popcorn": "爆米花",
    "Mix": "混合装",
    "Mix5": "五件混合装",
    "B": "B款",
}

PRODUCT_CATEGORY_OVERRIDES = {
    "FCGSS003": "不锈钢盖调料罐",
    "FCGSS004": "不锈钢盖调料罐",
    "FCKAD001": "硅藻泥吸水垫",
    "FCKAD002": "硅藻泥吸水垫",
}

PRODUCT_PREFIX_OVERRIDES = {
    "FCGSS003": "不锈钢盖调料罐-三格",
    "FCGSS004": "不锈钢盖调料罐-四格",
    "FCKAD001": "硅藻泥吸水垫",
    "FCKAD002": "硅藻泥吸水垫",
}

INLINE_VARIANT_PREFIXES = {"FCGSS003", "FCGSS004"}

PRODUCT_VARIANT_OVERRIDES = {
    "FCGSS003": {"Black": "黑色", "White": "白色"},
    "FCGSS004": {"Black": "黑色", "White": "白色"},
    "FCKAD001": {
        "Black": "黑",
        "Gray": "浅灰",
        "Nude": "米白",
        "Coffee": "咖啡",
        "BlackCoffee": "黑咖啡",
        "GrayCoffee": "浅灰咖啡",
    },
    "FCKAD002": {
        "Black": "黑",
        "Gray": "浅灰",
        "Nude": "米白",
        "Coffee": "咖啡",
        "BlackCoffee": "黑咖啡",
        "GrayCoffee": "浅灰咖啡",
    },
}


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_sku_id(value) -> str:
    text = clean_text(value)
    text = re.sub(r"\.0$", "", text)
    return text


def normalize_sku_text(value: str) -> str:
    text = clean_sku_id(value)
    return text.lower()


def split_seller_sku(seller_sku: str) -> tuple[str, str]:
    seller_sku = clean_text(seller_sku)
    if not seller_sku or seller_sku.lower().startswith("seller sku input"):
        return "", ""
    if "-" not in seller_sku:
        return seller_sku, ""
    base, suffix = seller_sku.split("-", 1)
    return base.strip(), suffix.strip()


def translate_variant(suffix: str, base_sku: str = "") -> str:
    suffix = clean_text(suffix)
    if not suffix:
        return ""
    base_map = PRODUCT_VARIANT_OVERRIDES.get(base_sku, {})
    if suffix in base_map:
        return base_map[suffix]
    if re.fullmatch(r"\d+\s*cm", suffix, flags=re.I):
        return suffix.lower().replace(" ", "")
    if re.fullmatch(r"\d+cm", suffix, flags=re.I):
        return suffix.lower()
    return VARIANT_MAP.get(suffix, suffix)


def derive_product_category(product_name: str) -> str:
    text = clean_text(product_name)
    text = re.sub(r"[\s\-]*\d+\s*cm$", "", text, flags=re.I)
    text = re.sub(r"(?:[零一二三四五六七八九十]+|[0-9]+)件套$", "", text)
    text = re.sub(r"(?:[零一二三四五六七八九十]+|[0-9]+)件装$", "", text)
    text = re.sub(r"(?:[零一二三四五六七八九十]+|[0-9]+)件$", "", text)
    text = re.sub(r"(?:大号|中号|小号)$", "", text)
    text = re.sub(r"(?:套餐|套装)$", "", text)
    text = re.sub(r"\s+", "", text)
    return text or product_name


def normalize_cost_table(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["SKU"] = df["SKU"].map(clean_sku_id)
    df["SKU_NORM"] = df["SKU"].map(normalize_sku_text)
    df = df[df["SKU"].ne("")]
    df["产品名称"] = df["产品名称"].map(clean_text)
    df["产品成本"] = pd.to_numeric(df.get("产品成本"), errors="coerce")
    df["长cm"] = pd.to_numeric(df.get("长cm"), errors="coerce")
    df["宽cm"] = pd.to_numeric(df.get("宽cm"), errors="coerce")
    df["高cm"] = pd.to_numeric(df.get("高cm"), errors="coerce")
    df["重量g"] = pd.to_numeric(df.get("重量g"), errors="coerce")
    df["父记录"] = df.get("父记录")
    # Prefer the base row without parent record when duplicates exist.
    df["_has_parent"] = df["父记录"].notna() & df["父记录"].astype(str).str.strip().ne("")
    df = (
        df.sort_values(["SKU", "_has_parent", "产品成本"], ascending=[True, True, True])
        .drop_duplicates(subset=["SKU"], keep="first")
        .drop(columns=["_has_parent"])
    )
    return df


def build_mapping(order_path: Path, cost_path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    cost_df = normalize_cost_table(pd.read_excel(cost_path))
    order_df = pd.read_excel(order_path, sheet_name=0)
    order_df.columns = [clean_text(c) for c in order_df.columns]

    required = ["SKU ID", "Seller SKU"]
    missing = [c for c in required if c not in order_df.columns]
    if missing:
        raise ValueError(f"订单表缺少必要列: {', '.join(missing)}")

    order_df["SKU ID"] = order_df["SKU ID"].map(clean_sku_id)
    order_df["Seller SKU"] = order_df["Seller SKU"].map(clean_text)
    order_df = order_df[order_df["SKU ID"].str.fullmatch(r"\d+", na=False)].copy()
    order_df = order_df[order_df["Seller SKU"].ne("")]

    seen = set()
    rows = []
    for _, row in order_df.iterrows():
        sku_id = row["SKU ID"]
        if sku_id in seen:
            continue
        seen.add(sku_id)

        seller_sku = row["Seller SKU"]
        base_sku, suffix = split_seller_sku(seller_sku)
        seller_norm = normalize_sku_text(seller_sku)
        cost_row = cost_df[cost_df["SKU_NORM"].eq(seller_norm)]
        if cost_row.empty:
            cost_row = cost_df[cost_df["SKU"].eq(base_sku)]

        if cost_row.empty:
            rows.append({
                "产品大类": "",
                "SKU ID": sku_id,
                "SKU中文简称": "",
                "产品成本(元)": "",
                "长cm": "",
                "宽cm": "",
                "高cm": "",
        "计费重kg": "",
                "Seller SKU": seller_sku,
                "匹配状态": "未匹配",
                "未匹配原因": "成本表无对应SKU",
            })
            continue

        cost_row = cost_row.iloc[0]
        variant = translate_variant(suffix, base_sku=base_sku)
        short_name = PRODUCT_PREFIX_OVERRIDES.get(base_sku, cost_row["产品名称"])
        if variant:
            if base_sku in INLINE_VARIANT_PREFIXES:
                short_name = f"{short_name}{variant}"
            else:
                short_name = f"{short_name}-{variant}"

        billable_weight = ""
        if "计费重kg" in cost_row.index and pd.notna(cost_row["计费重kg"]):
            billable_weight = round(float(cost_row["计费重kg"]), 2)
        elif "计费重量kg" in cost_row.index and pd.notna(cost_row["计费重量kg"]):
            billable_weight = round(float(cost_row["计费重量kg"]), 2)
        elif "计费重量g" in cost_row.index and pd.notna(cost_row["计费重量g"]):
            billable_weight = round(float(cost_row["计费重量g"]) / 1000, 2)
        elif "重量g" in cost_row.index and pd.notna(cost_row["重量g"]):
            billable_weight = round(float(cost_row["重量g"]) / 1000, 2)

        product_category = PRODUCT_CATEGORY_OVERRIDES.get(base_sku, derive_product_category(cost_row["产品名称"]))

        rows.append({
            "产品大类": product_category,
            "SKU ID": sku_id,
            "SKU中文简称": short_name,
            "产品成本(元)": round(float(cost_row["产品成本"]) * MX_COST_EXCHANGE_RATE, 2) if pd.notna(cost_row["产品成本"]) else "",
            "长cm": round(float(cost_row["长cm"]), 2) if pd.notna(cost_row["长cm"]) else "",
            "宽cm": round(float(cost_row["宽cm"]), 2) if pd.notna(cost_row["宽cm"]) else "",
            "高cm": round(float(cost_row["高cm"]), 2) if pd.notna(cost_row["高cm"]) else "",
            "计费重kg": billable_weight,
            "Seller SKU": seller_sku,
            "匹配状态": "已匹配",
            "未匹配原因": "",
        })

    mapping_df = pd.DataFrame(rows)
    if mapping_df.empty:
        return mapping_df, pd.DataFrame()

    matched_df = mapping_df[mapping_df["匹配状态"].eq("已匹配")].copy()
    unmatched_df = mapping_df[mapping_df["匹配状态"].ne("已匹配")].copy()
    matched_df = matched_df[["产品大类", "SKU ID", "SKU中文简称", "产品成本(元)", "长cm", "宽cm", "高cm", "计费重kg"]]
    matched_df = matched_df.sort_values(["产品大类", "SKU中文简称", "SKU ID"], kind="stable").reset_index(drop=True)
    return matched_df, unmatched_df


def main() -> None:
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    matched_df, unmatched_df = build_mapping(ORDER_FILE, COST_FILE)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        matched_df.to_excel(writer, sheet_name="SKU映射", index=False)
        if not unmatched_df.empty:
            unmatched_df.to_excel(writer, sheet_name="未匹配", index=False)

    print(f"✅ 已输出: {OUTPUT_FILE}")
    print(f"✅ 已匹配: {len(matched_df)} 条")
    if not unmatched_df.empty:
        print(f"⚠️ 未匹配: {len(unmatched_df)} 条")
        print(unmatched_df[["SKU ID", "Seller SKU"]].to_string(index=False))


if __name__ == "__main__":
    main()
