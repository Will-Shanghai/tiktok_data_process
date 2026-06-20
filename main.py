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


DEFAULT_STORES = [
    {
        "enabled": 1,
        "country_code": "JP",
        "country_name": "日本",
        "store_key": "local",
        "store_name": "本土店",
        "store_dir": "local",
        "sheet_name": "日本_本土店",
        "说明": "1=启用，0=不运行",
    },
    {
        "enabled": 1,
        "country_code": "JP",
        "country_name": "日本",
        "store_key": "cross-border",
        "store_name": "跨境店",
        "store_dir": "cross-border",
        "sheet_name": "日本_跨境店",
        "说明": "日本跨境目录名为 cross-border",
    },
    {
        "enabled": 1,
        "country_code": "JP",
        "country_name": "日本",
        "store_key": "direct",
        "store_name": "直邮店",
        "store_dir": "direct",
        "sheet_name": "日本_直邮店",
        "说明": "",
    },
    {
        "enabled": 1,
        "country_code": "VN",
        "country_name": "越南",
        "store_key": "local",
        "store_name": "本土店",
        "store_dir": "local",
        "sheet_name": "越南_本土店",
        "说明": "",
    },
    {
        "enabled": 1,
        "country_code": "VN",
        "country_name": "越南",
        "store_key": "cross_border",
        "store_name": "跨境店",
        "store_dir": "cross_border",
        "sheet_name": "越南_跨境店",
        "说明": "越南跨境目录名为 cross_border",
    },
]


def get_runtime_root():
    """Return the folder that contains config/data/result for this run."""
    env_root = os.getenv("TIKTOK_REPORT_ROOT")
    if env_root:
        return Path(env_root).resolve()

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
    for folder in [
        "data/data_JP/local",
        "data/data_JP/cross-border",
        "data/data_JP/direct",
        "data/data_VN/local",
        "data/data_VN/cross_border",
        "config/cache",
    ]:
        (APP_ROOT / folder).mkdir(parents=True, exist_ok=True)


def create_default_app_config():
    """Create config/app_config.xlsx when a packaged user has not copied one yet."""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    stores_df = pd.DataFrame(DEFAULT_STORES)
    notes_df = pd.DataFrame(
        [
            ["enabled", "1 表示运行，0 表示跳过。"],
            ["country_code", "JP=日本，VN=越南。"],
            ["store_key", "输出文件名使用，例如 Daily_Performance_Report_JP_direct.xlsx。"],
            ["store_dir", "订单 CSV 所在目录名，例如 data/data_JP/direct。"],
            ["sheet_name", "飞书配置表 Sheet 名称。"],
        ],
        columns=["字段", "说明"],
    )

    with pd.ExcelWriter(CONFIG_PATH, engine="openpyxl") as writer:
        stores_df.to_excel(writer, sheet_name="Stores", index=False)
        notes_df.to_excel(writer, sheet_name="README", index=False)

    print(f"⚠️ 未找到配置文件，已自动生成默认配置: {CONFIG_PATH}")
    print("   请按需修改 config/app_config.xlsx 的 enabled 列。")


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
        create_default_app_config()

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


def has_feishu_env():
    env_file = APP_ROOT / "config" / ".env"
    if env_file.exists():
        return True
    return bool(os.getenv("FEISHU_APP_ID") and os.getenv("FEISHU_APP_SECRET"))


def cache_file_for_sheet(sheet_name):
    safe_name = str(sheet_name).replace(" ", "_").replace("/", "_")
    return APP_ROOT / "config" / "cache" / f"config_{safe_name}.csv"


def check_config_sources(stores):
    """Give packaged users a clear fix when neither Feishu env nor cache exists."""
    if has_feishu_env():
        return

    missing = [store["sheet_name"] for store in stores if not cache_file_for_sheet(store["sheet_name"]).exists()]
    if not missing:
        return

    missing_text = "\n".join(f"   - {name}" for name in missing)
    raise RuntimeError(
        "当前没有配置飞书凭证，也缺少本地配置缓存，无法读取 SKU 成本配置。\n"
        f"缺少缓存的 Sheet:\n{missing_text}\n\n"
        "解决办法二选一：\n"
        "1. 把源码里的 sum_daily_order/config/cache 文件夹复制到 exe 同级的 config/cache；\n"
        "2. 在 exe 同级的 config/.env 中配置 FEISHU_APP_ID 和 FEISHU_APP_SECRET。"
    )


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

    enabled_stores = [
        store
        for site_code in selected_sites
        for store in grouped.get(site_code, [])
    ]
    check_config_sources(enabled_stores)

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
