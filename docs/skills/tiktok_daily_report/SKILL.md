# TikTok Daily Report Skill

## Purpose

Use this skill when working on the TikTok Shop daily order report tool in:

- `/Users/macbook/Documents/Projects/tiktok_data_process`

This skill is for maintaining, extending, debugging, or packaging the daily report workflow used for Japan and Vietnam TikTok Shop order-cost reporting.

## When To Use

Use this skill when the request is about any of the following:

- JP / VN daily order cost report scripts
- `main.py` unified launcher
- `sum_daily_order/config/app_config.xlsx`
- report sheet structure such as `Daily Summary`, `Daily Product Detail`, `SKU Detail`
- cost logic: product cost, logistics cost, sample cost
- multi-file period comparison
- Windows `exe` packaging and teammate delivery
- folder placement for store CSV files

Do not use this skill for:

- ad-report scripts under `cal_ads_data`
- product-card scripts under `cal_product_data`
- click-order scripts under `click_order_data`
- unrelated experimental scaffolds unless the user explicitly asks

## Current Project Boundary

Current maintained entrypoint:

- `main.py`

Current maintained country scripts:

- `sum_daily_order/config/cal_cost_return_jp_daily.py`
- `sum_daily_order/config/cal_cost_return_vn_daily.py`

Current support status:

- JP: local, cross-border, direct
- VN: local, cross_border
- MX: directories and config rows are reserved, but the daily script is not fully connected yet

Important:

- Do not assume older `app.py`, `tiktok_data_tool/`, or `docs/TOOL_USAGE.md` are the primary runtime path
- Prefer the current `main.py` based workflow unless the user explicitly wants to revive the old structure

## Runtime Layout

Source mode:

```text
sum_daily_order/
├── config/
│   ├── app_config.xlsx
│   └── cache/
├── data/
│   ├── data_JP/
│   ├── data_VN/
│   └── data_MX/
└── result/
```

Packaged mode:

```text
TikTokDailyReport/
├── TikTokDailyReport_v1.1.0.exe
├── config/
├── data/
└── result/
```

The runtime root is:

- source mode: `sum_daily_order/`
- packaged mode: the folder that contains the `exe`

The code already supports this with `TIKTOK_REPORT_ROOT` and runtime root detection.

## Core Business Rules

### Date Rule

- Daily grouping is based on `Paid Time`
- Orders without a valid paid time are usually excluded from paid-order summaries

### Store Folder Rule

CSV files must be placed in the correct store directory.

Examples:

- JP local: `data/data_JP/local`
- JP cross-border: `data/data_JP/cross-border`
- JP direct: `data/data_JP/direct`
- VN local: `data/data_VN/local`
- VN cross-border: `data/data_VN/cross_border`

Wrong folder placement causes wrong cost rules to be applied.

### Product Mapping Rule

- Product-level reporting uses `产品大类`
- SKU / variant-level detail is preserved separately
- Name mapping must follow the actual business naming used in the Feishu config

### Cost Rule

- Product cost is quantity-based
- Logistics cost is order-based
- Sample cost is sample-quantity-based

If sample cost from config is blank or unreadable because of formula evaluation, treat:

- `寄样成本(元) = 产品成本(元) + 物流成本(元)`

### Logistics Rule

Do not simply multiply logistics by line quantity.

Use the configured `每单物流承载数量`:

- same-order same-category items share logistics according to capacity
- mixed-product orders count one order-level logistics cost
- mixed-product logistics is proportionally allocated back to product categories

Read the config range wide enough to include the logistics capacity column.

## Report Semantics

### Daily Summary

Grouped by:

- `文件名 + 日期`

Used for:

- order count
- quantity
- sample quantity
- sales
- cost totals

### Daily Product Detail

Top section:

- `文件名 + 日期 + 产品大类`

Bottom section:

- `文件名 + 产品大类 + 产品名称`

This sheet should still let the user inspect daily product performance.
Do not remove the `日期` dimension from the top section unless the user explicitly changes the requirement.

### SKU Detail

Grouped by:

- `文件名 + 日期 + 产品大类 + 产品名称`

Used for SKU / variant troubleshooting.

### Product Quantity Matrix

Top section:

- product x file summary

Lower section:

- `文件名 + 产品名称 + 日期`

### Sample Statistics

Grouped as:

- `文件名 + 产品名称 + 日期`

### Period Comparison

- one file: still output single-period summary
- two or more files: output cross-file comparison

Do not require at least two files just to show a summary.

## Output Formatting Rules

Keep the current workbook behavior unless the user asks otherwise:

- center align written cells
- percentage columns shown as percent
- blank separator rows inserted between different files in multi-file sections
- file names shown directly in comparison tables instead of generic labels

## Feishu Config Rules

Config can come from:

- `config/.env` + live Feishu fetch
- local cache under `config/cache/`

If no `.env` exists, prefer using local cache so the packaged tool can still run offline.

When the user says Feishu changes are not reflected:

1. check whether the program used cache first
2. check whether `.env` exists in the runtime `config/`
3. check whether the cache file for that sheet was refreshed

## Working Style For This Repo

When implementing changes in this repo:

1. read the current report shape before editing
2. preserve the user's agreed business semantics
3. avoid broad refactors outside `main.py` and the country daily scripts
4. do not commit local CSV data, cache files, or result workbooks
5. verify with at least syntax compilation after code changes

Preferred checks:

```bash
python -m py_compile sum_daily_order/config/cal_cost_return_jp_daily.py
python -m py_compile sum_daily_order/config/cal_cost_return_vn_daily.py
```

If the user asks for a push:

- commit only code/doc changes related to the request
- exclude local data and generated artifacts

## Common Pitfalls

- Mistaking estimated gross profit for final finance profit
- Putting CSV files into the wrong store directory
- Forgetting that JP uses `cross-border` while VN uses `cross_border`
- Reading too few config columns and missing `每单物流承载数量`
- Assuming Feishu formula cells always return computed numeric values
- Removing `日期` from `Daily Product Detail` and making daily product inspection impossible
- Merging overlapping-date files into one summary and double-counting

## Recommended First Read Order

When starting a fresh session on this repo, read in this order:

1. `README.md`
2. `TikTokDailyReport_使用说明.txt`
3. `progress.md`
4. `main.py`
5. the relevant country daily script

## Expected Response Style

When helping on this project:

- explain business impact plainly
- point out whether an issue is logic, config, cache, or path related
- separate "current behavior" from "recommended change"
- be careful with file-count, date-range, and store-directory assumptions
- if the user says "先不要写代码", stay at solution level first
- if the user says "开始改代码", implement directly
