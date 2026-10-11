"""Fetch missing public conversion configuration without replacing local files."""

import csv
import io
import json
from pathlib import Path

import requests


CONFIG_SOURCES = (
    "https://gitee.com/dlutwill/tiktok_data_process/raw/master/sum_daily_conversion/config/",
    "https://raw.githubusercontent.com/Will-Shanghai/tiktok_data_process/master/sum_daily_conversion/config/",
)


def validate_config(name, content):
    text = content.decode("utf-8-sig")
    if name == "config.json":
        config = json.loads(text)
        if not isinstance(config, dict):
            raise ValueError("配置必须是 JSON 对象")
        for key in ("data_dir", "mapping_file", "output_dir", "log_dir"):
            if not isinstance(config.get(key), str) or not config[key].strip():
                raise ValueError(f"配置缺少有效字段 {key}")
        if config.get("output_mode", "local_excel") not in ("local_excel", "feishu"):
            raise ValueError("无效的 output_mode")
    else:
        reader = csv.DictReader(io.StringIO(text))
        fields = set(reader.fieldnames or [])
        if not {"店铺", "商品ID"} <= fields or not fields.intersection({"商品名称", "商品名", "飞书sheet名"}):
            raise ValueError("映射 CSV 列名不正确")
        if not any((row.get("店铺") or "").strip() for row in reader):
            raise ValueError("映射 CSV 没有店铺记录")
    return text.encode("utf-8")


def ensure_conversion_configs(root):
    """Download only missing default files; never fetch credentials or business data."""
    root = Path(root)
    config_dir = root / "config"
    config_dir.mkdir(parents=True, exist_ok=True)
    for name in ("config.json", "product_sheet_mapping.csv"):
        target = config_dir / name
        if target.exists():
            continue
        if name == "product_sheet_mapping.csv":
            config = json.loads((config_dir / "config.json").read_text(encoding="utf-8-sig"))
            mapping = Path(config.get("mapping_file", "config/product_sheet_mapping.csv"))
            configured_target = mapping if mapping.is_absolute() else root / mapping
            if configured_target.resolve() != target.resolve():
                # A custom mapping belongs to the user, not the remote default.
                if not configured_target.exists():
                    raise FileNotFoundError(f"本地配置指定的自定义映射不存在：{configured_target}")
                continue
        failures = []
        for source in CONFIG_SOURCES:
            try:
                with requests.get(source + name, timeout=(5, 15)) as response:
                    response.raise_for_status()
                    content = validate_config(name, response.content)
                # Exclusive creation preserves a file created by another process.
                try:
                    with target.open("xb") as handle:
                        handle.write(content)
                except FileExistsError:
                    pass
                print(f"[INFO] 配置已就绪：{target}")
                break
            except (requests.RequestException, ValueError, UnicodeError) as exc:
                failures.append(f"{source}: {type(exc).__name__}")
        else:
            raise RuntimeError(
                f"无法自动下载 {name}。请检查网络或仓库访问权限，"
                f"也可手动复制到 {target}。已有本地配置保持不变。\n" + "\n".join(failures)
            )
