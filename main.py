# -*- coding: utf-8 -*-
"""
TikTok Shop daily report launcher.

This is the packaging-friendly entry point for Windows/macOS users. It reads
config/app_config.xlsx, then runs the enabled JP/VN daily report jobs.
"""

import argparse
import os
import sys
from pathlib import Path

import pandas as pd


def get_runtime_root():
    """Return the folder that contains config/data/result for this run."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    source_root = Path(__file__).resolve().parent
    sum_daily_root = source_root / "sum_daily_order"
    if (sum_daily_root / "config").is_dir() and (sum_daily_root / "data").is_dir():
        return sum_daily_root
    return source_root


APP_ROOT = get_runtime_root()
CONFIG_PATH = APP_ROOT / "config" / "app_config.xlsx"


def ensure_runtime_dirs():
    for folder in ["config", "data", "result"]:
        (APP_ROOT / folder).mkdir(parents=True, exist_ok=True)


def parse_enabled(value):
    return str(value).strip().lower() in {"1", "true", "yes", "y", "是", "启用"}


def normalize_country_code(value):
    value = str(value).strip().upper()
    aliases = {
        "日本": "JP",
        "越南": "VN",
    }
    return aliases.get(value, value)


def load_app_config():
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"未找到配置文件: {CONFIG_PATH}\n"
            "请确认 config/app_config.xlsx 与程序在同一个工具目录下。"
        )

    df = pd.read_excel(CONFIG_PATH, sheet_name="Stores", dtype=str).fillna("")
    required_cols = [
        "enabled",
        "country_code",
        "country_name",
        "store_key",
        "store_name",
        "store_dir",
        "sheet_name",
    ]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"app_config.xlsx 的 Stores Sheet 缺少列: {', '.join(missing)}")

    stores = []
    for _, row in df.iterrows():
        if not parse_enabled(row["enabled"]):
            continue
        store = {
            "enabled": True,
            "country_code": normalize_country_code(row["country_code"]),
            "country_name": str(row["country_name"]).strip(),
            "store_key": str(row["store_key"]).strip(),
            "store_name": str(row["store_name"]).strip(),
            "store_dir": str(row["store_dir"]).strip(),
            "sheet_name": str(row["sheet_name"]).strip(),
        }
        if all(store.values()):
            stores.append(store)
    return stores


def choose_site():
    print("\n请选择要运行的站点：")
    print("1. 日本 JP")
    print("2. 越南 VN")
    print("3. 全部 all")
    choice = input("请输入数字后回车: ").strip()
    return {"1": "JP", "2": "VN", "3": "all"}.get(choice, "")


def split_stores_by_site(stores):
    grouped = {"JP": [], "VN": []}
    for store in stores:
        country_code = normalize_country_code(store["country_code"])
        if country_code in grouped:
            grouped[country_code].append(store)
    return grouped


def run(site):
    ensure_runtime_dirs()

    os.environ["TIKTOK_REPORT_ROOT"] = str(APP_ROOT)
    os.environ.setdefault("USE_CONFIG_CACHE_FIRST", "1")

    stores = load_app_config()
    grouped = split_stores_by_site(stores)

    if site == "all":
        selected_sites = ["JP", "VN"]
    else:
        selected_sites = [normalize_country_code(site)]

    if "JP" in selected_sites:
        if grouped["JP"]:
            from sum_daily_order.config.cal_cost_return_jp_daily import run_japan_daily

            run_japan_daily(grouped["JP"])
        else:
            print("⚠️ app_config.xlsx 中没有启用的日本店铺，已跳过。")

    if "VN" in selected_sites:
        if grouped["VN"]:
            from sum_daily_order.config.cal_cost_return_vn_daily import run_vietnam_daily

            run_vietnam_daily(grouped["VN"])
        else:
            print("⚠️ app_config.xlsx 中没有启用的越南店铺，已跳过。")


def main():
    parser = argparse.ArgumentParser(description="TikTok Shop daily report launcher")
    parser.add_argument("--site", choices=["JP", "VN", "all"], help="直接运行指定站点")
    parser.add_argument("--list", action="store_true", help="列出 app_config.xlsx 中启用的店铺")
    args = parser.parse_args()

    print("=" * 60)
    print("TikTok Shop 日报生成工具")
    print("=" * 60)
    print(f"工具目录: {APP_ROOT}")
    print(f"配置文件: {CONFIG_PATH}")

    try:
        if args.list:
            stores = load_app_config()
            if not stores:
                print("没有启用的店铺。")
                return
            for store in stores:
                print(f"- {store['country_name']}_{store['store_name']} ({store['country_code']}/{store['store_key']})")
            return

        site = args.site or choose_site()
        if site not in {"JP", "VN", "all"}:
            print("未选择有效站点，程序结束。")
            return
        run(site)
    except Exception as exc:
        print(f"\n❌ 程序运行失败: {exc}")
    finally:
        if getattr(sys, "frozen", False):
            input("\n按回车键退出...")


if __name__ == "__main__":
    main()
