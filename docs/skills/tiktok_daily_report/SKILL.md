---
name: tiktok-data-process
description: Maintain and debug the TikTokShopDataTool project in /Users/macbook/Documents/Projects/tiktok_data_process, including order daily reports, product daily conversion, Feishu config, and Windows EXE packaging.
---

# TikTokShopDataTool Skill

Use this skill when working in:

```text
/Users/macbook/Documents/Projects/tiktok_data_process
```

This project is now a multi-function TikTok Shop data tool, not only a daily report tool.

## Current Entry Point

The maintained entry point is:

```text
main.py
```

The launcher asks the user to choose:

- `1` order cost daily report
- `2` product daily conversion
- `3` all

If the user chooses order daily reports, it then asks for site: JP, VN, MX, or all.

## Module Boundaries

Keep the two modules separate:

```text
sum_daily_order/
├── config/
├── data/
└── result/

sum_daily_conversion/
├── config/
├── data/
└── result/
```

Do not put business config back into a root-level `config/`.

Do not put order `config/data/result` beside the EXE root. In packaged mode they belong under `sum_daily_order/`.

## Order Daily Reports

Main scripts:

- `sum_daily_order/config/cal_cost_return_jp_daily.py`
- `sum_daily_order/config/cal_cost_return_vn_daily.py`
- `sum_daily_order/config/cal_cost_return_mx_daily.py`

Store config:

- `sum_daily_order/config/app_config.xlsx`

Feishu/cache config:

- live Feishu via `sum_daily_order/config/.env`
- cached config under `sum_daily_order/config/cache/`

Supported stores:

- JP: local, cross_border, direct_old, direct_new
- VN: local, cross_border
- MX: local, direct_old, direct_new

Important rules:

- The display name is now `除运费外销售额`, not `P列折后价`.
- `除运费外销售额 = SKU Platform Discount + SKU Subtotal After Discount`.
- Feishu merged cells in `产品大类` must be forward-filled after reading config.
- Product-level summaries use `产品大类`; SKU-level details keep variant names.
- Profits are estimated gross margin, not final financial profit.

Mexico-specific rules:

- MX local store has no normal-order IVA and no sample IVA.
- MX local sales should exclude `Original Shipping Fee`.
- MX local logistics = `头程物流成本(元) + 尾程物流成本(元)`.
- MX direct uses direct-mail weight/logistics rules and IVA conversion.
- If MX direct lacks real weight or dimensions, logistics cost should be 0 and a warning should be printed.

Japan direct rules:

- JP direct old/new stores are separate directories.
- JP direct uses goods type and weight card rules.
- If an order contains special/sensitive goods, use the special/sensitive price card.
- Only all-normal-goods orders use the normal-goods price card.

## Product Daily Conversion

Main script:

```text
sum_daily_conversion/config/cal_product_daily_conversion.py
```

Config:

```text
sum_daily_conversion/config/config.json
sum_daily_conversion/config/product_sheet_mapping.csv
```

Input:

```text
sum_daily_conversion/data/
```

Output:

```text
sum_daily_conversion/result/
```

Diagnostics:

```text
sum_daily_conversion/result/logs/
```

Keep `result/logs/`. It stores dry-run previews, skipped unmapped products, and mapping diagnostics.

### Mapping matching rules (important)

`product_sheet_mapping.csv` is keyed by `(店铺, 商品ID)` and supports two data layouts:

1. **Store subdirectories** — `data/日本本土店/xxx.xlsx`. The store comes from the folder
   name, so matching with `(store, product_id)` is exact.
2. **Flat directory** — `data/xxx.xlsx`. There is no store info in the path, so the store
   degrades to the placeholder `商品每日转化统计` (`DEFAULT_FLAT_STORE_NAME`).

`ProductMapping.resolve` order:

1. Exact `(store, product_id)` hit → use it.
2. Store **is** a known store in the mapping but this product is not configured for it →
   genuinely unmapped, reason `未配置映射`. Do **not** fall back across stores here, or a
   product from store A could be written into store B's sheet.
3. Store is the flat placeholder / unknown → fall back to `product_id`, then to `商品名`.
4. Fallback only succeeds when the candidate sheet is **unique**. If one product ID maps to
   several different sheets, skip with `映射歧义` and list the candidates in the log instead
   of guessing.

Never re-tighten this back to a strict `(store, product_id)`-only lookup: that is exactly what
produced 0 output records on a flat `data/` directory.

`商品名` is an optional mapping column, used only when `商品ID` is empty or the `xxx` placeholder.

### Resolved issue: 0 syncable records

Root cause was the mapping lookup, **not** the paths:

- `data/` was flat, so `discover_store_dirs()` returned `[]` and the store became the
  placeholder `商品每日转化统计`.
- `mapping.get(("商品每日转化统计", product_id))` never matched the real store names, so all
  743 source rows were skipped as `未配置映射` and the output was 0 records.
- After the fix: dry-run parses **32** records; real mode writes
  `sum_daily_conversion/result/商品每日转化统计_每天转化数据.xlsx`.

Remaining work is data, not code:

- The mapping covers only 26 unique product IDs (27 rows) while the two source files hold 384
  unique product IDs. 16 IDs match → 32 records; the other **711** rows still need mapping rows.
- Use `sum_daily_conversion/result/logs/mapping_diagnostics.csv` to fill the table. Status
  values: `已匹配` / `未出现在源数据` / `歧义（同ID多sheet）`.
- `1736121773627573268` maps to both `日本本土店=蜻蜓挂件` and `日本跨境店=鼻毛刀`. It is not
  in the source data today, but in flat mode it would be reported as `映射歧义` and skipped.
- Missing `发品状态` only raises a WARN; that column stays empty in the output.

## Packaging

Current EXE/package name:

```text
TikTokShopDataTool_v<version>.exe
TikTokShopDataTool_v<version>_windows.zip
```

Manual build command:

```bash
pyinstaller --clean --onefile --name TikTokShopDataTool_v1.1.0 --hidden-import sum_daily_conversion.config.cal_product_daily_conversion main.py
```

Update both files when changing packaging:

- `.github/workflows/build-windows-exe.yml`
- `.workflow/流水线-202607222147.yml`

Release package should look like:

```text
TikTokShopDataTool_v<version>/
├── TikTokShopDataTool_v<version>.exe
├── TikTokShopDataTool_使用说明.txt
├── VERSION
├── sum_daily_order/
│   ├── config/
│   ├── data/
│   └── result/
└── sum_daily_conversion/
    ├── config/
    ├── data/
    └── result/
```

## Working Rules

- Check `git status --short` before editing.
- Do not revert user changes unless explicitly asked.
- Do not commit `.env`, generated result workbooks, cache folders, or local data unless the user explicitly wants that.
- Prefer targeted changes to one module; do not refactor JP/VN/MX cost logic together unless needed.
- After code changes, run at least:

```bash
python -m py_compile main.py \
  sum_daily_order/config/cal_cost_return_jp_daily.py \
  sum_daily_order/config/cal_cost_return_vn_daily.py \
  sum_daily_order/config/cal_cost_return_mx_daily.py \
  sum_daily_conversion/config/cal_product_daily_conversion.py
```

For conversion changes, also run:

```bash
python main.py --task conversion --dry-run
```

For launcher changes, test menu routing with:

```bash
printf '2\n' | python main.py
```

## Local Run Notes

`main.py` imports `pandas`, so a bare `python3` may fail with
`ModuleNotFoundError: No module named 'pandas'`. In this environment use the managed venv,
which already has pandas + openpyxl:

```bash
/Users/macbook/.workbuddy/binaries/python/envs/default/bin/python main.py --task conversion --dry-run
```

Regression test for the mapping fallback (store-directory mode must still match exactly):

1. Create a temp root with `config/config.json`, `config/product_sheet_mapping.csv`,
   `data/日本直邮一店/0903-0909商品数据.xlsx` and `result/logs/`.
2. Run with `TIKTOK_CONVERSION_ROOT=<temp root>` set *before* importing the module and assert
   16 records are parsed with match mode `精确(店铺+商品ID)` in `mapping_diagnostics.csv`.

## Packaging Sync

Both pipeline files already carry the required flag and the `TikTokShopDataTool_v<version>`
naming — verify, do not duplicate:

- `.github/workflows/build-windows-exe.yml`
- `.workflow/流水线-202607222147.yml`

If you ever change packaging, keep `--hidden-import sum_daily_conversion.config.cal_product_daily_conversion`
(and the `TikTokShopDataTool_${tagVersion}` / `_windows` names) identical in both files.

## First Files To Read

For handoff or unfamiliar work:

1. `progress.md`
2. `README.md`
3. `TikTokShopDataTool_使用说明.txt`
4. `main.py`
5. The relevant module script under `sum_daily_order/config/` or `sum_daily_conversion/config/`
