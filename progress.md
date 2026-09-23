# TikTokShopDataTool Progress

最后更新：2026-09-17

## 1. 项目定位

这个仓库当前维护的是一个 TikTok Shop 多功能数据工具，面向运营同事本地运行和 Windows EXE 交付。

当前主入口：

- `main.py`

当前工具名：

- `TikTokShopDataTool`

当前主要功能：

- 订单成本日报/周期报表：`sum_daily_order`
- 商品每日转化统计：`sum_daily_conversion`

原来的 `TikTokDailyReport` 命名正在逐步替换为 `TikTokShopDataTool`。后续新增功能时，不要再把工具定位成“只做日报”。

## 2. 当前目录结构

目标结构是根目录只放入口、说明、打包配置；业务模块各自维护自己的 `config/data/result`。

```text
tiktok_data_process/
├── main.py
├── README.md
├── TikTokShopDataTool_使用说明.txt
├── VERSION
├── progress.md
├── docs/
│   └── skills/
│       └── tiktok_daily_report/
│           └── SKILL.md
├── sum_daily_order/
│   ├── config/
│   │   ├── app_config.xlsx
│   │   ├── .env.example
│   │   └── cache/
│   ├── data/
│   │   ├── data_JP/
│   │   ├── data_VN/
│   │   └── data_MX/
│   └── result/
└── sum_daily_conversion/
    ├── config/
    │   ├── cal_product_daily_conversion.py
    │   ├── config.json
    │   └── product_sheet_mapping.csv
    ├── data/
    └── result/
```

重要约定：

- 订单日报的配置、数据、结果都在 `sum_daily_order/` 下面。
- 商品每日转化统计的配置、数据、结果都在 `sum_daily_conversion/` 下面。
- 不要再把业务配置放回根目录 `config/`。

## 3. 主程序运行逻辑

启动 `main.py` 后先选择功能：

```text
1. 订单成本日报
2. 商品每日转化统计
3. 全部运行
```

如果选择订单成本日报，再继续选择站点：

```text
1. 日本 JP
2. 越南 VN
3. 墨西哥 MX
4. 全部 all
```

如果选择商品每日转化统计，不再选择站点，直接读取 `sum_daily_conversion/data`。

命令行常用方式：

```bash
python main.py --task order --site JP
python main.py --task order --site VN
python main.py --task order --site MX
python main.py --task conversion --dry-run
python main.py --task all --site all
python main.py --list
```

为了兼容旧用法，`python main.py --site JP` 会默认跑订单日报。

## 4. 订单日报当前状态

主要脚本：

- `sum_daily_order/config/cal_cost_return_jp_daily.py`
- `sum_daily_order/config/cal_cost_return_vn_daily.py`
- `sum_daily_order/config/cal_cost_return_mx_daily.py`

已接入站点和店铺：

- 日本：本土店、跨境店、直邮老店、直邮新店
- 越南：本土店、跨境店
- 墨西哥：本土店、直邮老店、直邮新店

店铺配置来自：

- `sum_daily_order/config/app_config.xlsx`

SKU/成本配置来自：

- 飞书在线读取
- 或 `sum_daily_order/config/cache/` 本地缓存

订单日报输出目录：

- `sum_daily_order/result/`

## 5. 商品每日转化统计当前状态

主要脚本：

- `sum_daily_conversion/config/cal_product_daily_conversion.py`

配置文件：

- `sum_daily_conversion/config/config.json`
- `sum_daily_conversion/config/product_sheet_mapping.csv`

输入目录：

- `sum_daily_conversion/data/`

输出目录：

- `sum_daily_conversion/result/`

日志/排查目录：

- `sum_daily_conversion/result/logs/`

`logs` 目录建议保留。它用于保存 dry-run 预览、未匹配商品清单和映射表体检表，方便排查为什么没有出结果。它不是主要业务结果目录，不需要运营同事日常查看。

当前已验证（2026-09-17 修复后）：

- `python main.py --task conversion --dry-run` 能运行到新目录，输出 **32 条**（修复前是 0 条）。
- 真实模式 `python main.py --task conversion` 能写出 `sum_daily_conversion/result/商品每日转化统计_每天转化数据.xlsx`，含 1 个「新视频数汇总」+ 16 个商品 sheet。
- 语法检查和 dry-run 均通过。

### 5.1 0 条输出的根因（已修复）

根因是**映射匹配逻辑**，不是路径问题：

- `product_sheet_mapping.csv` 的主键是 `(店铺, 商品ID)`，店铺取值为 `日本本土店 / 日本跨境店 / 日本直邮一店 / 日本直邮二店`。
- 但 `sum_daily_conversion/data/` 当前是**扁平目录**（两个 xlsx 直接放在 data 下，没有店铺子目录）。
- `discover_store_dirs()` 返回空 → 店铺退化成占位名 `商品每日转化统计`（`DEFAULT_FLAT_STORE_NAME`）。
- 于是 `mapping.get(("商品每日转化统计", 商品ID))` 永远命中不了 → 743 条全部按「未配置映射」跳过 → 输出 0 条。

修复方式（`ProductMapping.resolve`）：

1. 先按 `(店铺, 商品ID)` 精确匹配，命中即用。
2. 如果店铺是映射表里**已知的真实店铺**但该商品没配，判定为真正的「未配置映射」，**不做跨店铺回退**（避免把 A 店商品写进 B 店 sheet）。
3. 如果店铺是扁平占位名/未知店铺，退化为按 **商品ID** 匹配，商品ID 无结果时再按 **商品名** 匹配。
4. 回退时只有候选 sheet **唯一**才算匹配成功；同一商品ID 对应多个不同 sheet 时按「映射歧义」跳过，并在日志里列出候选明细，避免写错 sheet。

### 5.2 映射表匹配规则

- 精确优先：分店铺目录（`data/日本本土店/xxx.xlsx`）走精确匹配，且优先于任何回退。
- 扁平回退：`data/xxx.xlsx` 无店铺信息时才回退到 商品ID / 商品名。
- 歧义安全：回退时同 ID 多 sheet 一律跳过报错，不猜。
- `product_sheet_mapping.csv` 可选列 `商品名`：当商品ID 为空或写成 `xxx` 占位时，用商品名兜底。

### 5.3 当前待处理

- **需要运营补齐映射表**：映射表只有 26 个唯一商品ID（27 行），而源数据两期合计 384 个唯一商品ID。修复后命中 16 个 ID → 32 条记录，剩余 **711 条**仍是「未配置映射」。
- 补齐依据：`sum_daily_conversion/result/logs/mapping_diagnostics.csv`，标注每条映射是「已匹配 / 未出现在源数据 / 歧义（同ID多sheet）」；未匹配商品明细在 `skipped_unmapped_products.csv`。
- `1736121773627573268` 在映射表里同时指向 `日本本土店=蜻蜓挂件` 和 `日本跨境店=鼻毛刀`，属于同 ID 多 sheet。该 ID 目前不在源数据里，暂未触发；若以后出现，扁平模式下会被判定为「映射歧义」跳过，需要拆数据目录或调整映射表。
- 源文件提示缺少 `发品状态` 字段（每个文件一行 WARN）。已核实：**不是错误，也不影响任何计算**。
  - 机制：`STANDARD_COLUMNS` 声明了 41 个目标列，脚本逐列去源表表头里找；源表 175 列里没有 `发品状态` → 收进 `missing_fields` 打印 WARN。其余 40 列全部命中。
  - 后果：`get_standard_value()` 取不到列索引时返回 `""`，所以输出里表头保留 `发品状态`、每行值为空（实际验证：`0903 | 1735933257201321438 | '' | 61,216円`）。
  - 影响面：`发品状态` 在整个代码库里只出现在 `cal_product_daily_conversion.py:24` 的列声明处，**没有任何逻辑读取它**，不参与筛选、计算或飞书写入判断。
  - 附注：`config.json` 里的 `gmv_max_rules`（target_roi / base_test_cost / impression_threshold 等）**从未被任何代码引用**，属死配置；`发品状态` 可能是当初那套未实现的「GMV 上限/投放建议」功能留下的字段。
  - 要清理的话有三种方向（尚未执行，需业务确认）：① 保持现状；② 从 `STANDARD_COLUMNS` 移除该列（告警消失，但会改变列数 → `upsert_daily_row` 的表头比对会触发重写飞书表头）；③ 真正需要该列 → 导出时带上，或在代码里加字段别名兼容。
- 长期方案（可选）：把 data 改成 `data/<店铺名>/xxx.xlsx` 分店铺摆放，即可走精确匹配、彻底消除歧义。


## 6. 关键业务规则

### 6.1 订单日报通用规则

- 日期归类主要基于 `Paid Time`。
- 产品级汇总使用飞书配置中的 `产品大类`。
- SKU 级明细保留到 `SKU明细` 或相关矩阵里。
- 报表里的利润/利润率只是预估毛利，不是最终财务净利润。
- 当前未扣除完整达人佣金、广告费、平台服务费、退款售后、仓储人工和税费等。

### 6.2 除运费外销售额

所有站点订单表里的“除运费外销售额”口径已统一为：

```text
SKU Platform Discount + SKU Subtotal After Discount
```

不要再使用旧的 `P列折后价` 命名。

### 6.3 飞书合并单元格

飞书配置表里的 `产品大类` 可能是合并单元格。读取配置后要向下填充，否则同一产品系列会被拆成带颜色/规格的多个产品大类。

日本、越南、墨西哥脚本里都已加过 `ffill()` 处理。

### 6.4 墨西哥本土店和直邮店区别

墨西哥本土店：

- 没有正常订单 IVA。
- 没有寄样 IVA。
- 销售额不应包含 `Original Shipping Fee`。
- 物流成本 = 飞书配置的 `头程物流成本(元) + 尾程物流成本(元)`。
- 样品支出 = 产品成本 + 物流成本。

墨西哥直邮店：

- 使用墨西哥直邮运费规则。
- 物流成本由计费重匹配价卡后换算人民币。
- 正常订单 IVA 和寄样 IVA 需要换算成人民币。
- 如果飞书配置有真实 `进口IVA` 且为负数，取绝对值后按汇率换算；如果为 0，则按销售额估算。
- 如果缺实重或长宽高，运费按 0 处理，并提示补飞书配置。

### 6.5 日本直邮店

- 日本直邮老店和新店分目录处理。
- 日本直邮使用特货/敏货/普货价卡。
- 只要订单内有特货或敏货，就按特货/敏货价卡。
- 全部普货才按普货价卡。
- 订单物流成本可按重量分摊回 SKU。

## 7. 打包状态

GitHub Actions 文件：

- `.github/workflows/build-windows-exe.yml`

备用流水线：

- `.workflow/流水线-202607222147.yml`

当前打包名称：

- `TikTokShopDataTool_v<版本号>.exe`
- `TikTokShopDataTool_v<版本号>_windows.zip`

手动打包命令（现在使用 `onedir`，不要只复制单个 EXE）：

```bash
pyinstaller --clean --onedir --name TikTokShopDataTool_v1.1.0 --hidden-import sum_daily_conversion.config.cal_product_daily_conversion main.py
```

打包后目标结构：

```text
TikTokShopDataTool_v<版本号>/
├── TikTokShopDataTool_v<版本号>.exe
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

不要把订单日报的 `config/data/result` 放到 EXE 根目录。

## 8. 当前注意事项和风险

- 不要提交真实 `.env`。
- 不要删除 `sum_daily_conversion/result/logs/`，除非用户明确只想清理输出。
- 不要把 `sum_daily_conversion/config/config.json` 和 `product_sheet_mapping.csv` 移回根目录。
- 不要把墨西哥本土店和直邮店逻辑混在一起改。
- 不要把日本/越南/墨西哥的成本逻辑抽成一个公共算法，除非有明确测试覆盖。
- 修改打包逻辑时，要同步 GitHub Actions 和备用 `.workflow`。
- 工作区可能有用户或上一次任务留下的未提交改动，改代码前先看 `git status`。

## 9. 建议 WorkBuddy 接手顺序

1. 先看本文件 `progress.md`。
2. 再看 `README.md` 和 `TikTokShopDataTool_使用说明.txt`。
3. 进入 `main.py`，理解功能分流。
4. 订单日报问题再看对应国家脚本。
5. 每日转化问题先看 `sum_daily_conversion/config/cal_product_daily_conversion.py` 和 `product_sheet_mapping.csv`。
6. 当前最应该优先处理的是转化统计映射表不匹配导致 0 条输出的问题。

## 10. 最近验证命令

```bash
python -m py_compile main.py \
  sum_daily_order/config/cal_cost_return_jp_daily.py \
  sum_daily_order/config/cal_cost_return_vn_daily.py \
  sum_daily_order/config/cal_cost_return_mx_daily.py \
  sum_daily_conversion/config/cal_product_daily_conversion.py

python main.py --task conversion --dry-run
python main.py --list
```

最近验证结果：

- 语法检查通过。
- 转化统计能运行到 `sum_daily_conversion` 目录。
- 转化统计修复前 0 条（映射未匹配，非路径错误）；修复后 dry-run 输出 32 条，真实模式能写出 `result/商品每日转化统计_每天转化数据.xlsx`。
- 分店铺目录模式回归测试同样通过（16 条，命中方式为「精确(店铺+商品ID)」）。
