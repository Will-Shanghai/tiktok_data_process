# TikTok Data Process

这个项目用于生成 TikTok Shop 订单成本日报/周期报表。当前重点维护：

- 日本站日报：`sum_daily_order/config/cal_cost_return_jp_daily.py`
- 越南站日报：`sum_daily_order/config/cal_cost_return_vn_daily.py`
- 打包友好的统一入口：`main.py`

## 成本与利润口径说明

报表里的“利润”“利润率”不是财务净利润，只能理解为“预估毛利”或“扣除已配置成本后的余额”。

当前程序主要计算：

- 销售额 / 汇率后金额
- 产品成本
- 物流成本
- 寄样成本 / 寄样支出

成本计算口径：

- 产品成本：按销量计算，即 `销量 × 产品成本(元)`。
- 物流成本：按订单计算，不再简单按销量重复计算。同一订单内会先根据配置表 `每单物流承载数量` 计算各产品大类的候选物流成本，然后整单只计一次物流成本；混合产品订单会按候选物流成本比例分摊到产品大类。
- 寄样成本：按样品单数量计算，即 `寄样数 × 寄样成本(元)`。

当前程序没有扣除：

- 达人佣金
- 广告费
- 平台佣金 / 平台服务费
- 支付手续费
- 退款、售后、赔付
- 仓储、人工、税费等其他经营成本

所以报表中的“利润”适合用于快速判断产品是否有毛利空间、哪个文件或产品表现更好；不适合作为最终利润、财务利润或结算利润。如果要做真实利润核算，需要在导出的报表基础上继续扣除达人、广告、平台费等费用。

## 目录结构

源码运行时主要使用：

```text
sum_daily_order/
├── config/
│   ├── app_config.xlsx          # 控制启用国家和店铺
│   ├── .env                     # 飞书应用凭证，可选
│   └── cache/                   # 飞书配置缓存
├── data/
│   ├── data_JP/
│   │   ├── local/
│   │   ├── cross-border/
│   │   ├── direct_old/
│   │   └── direct_new/
│   └── data_VN/
│       ├── local/
│       └── cross_border/
└── result/                      # 报表输出目录
```

Windows 打包后，建议给同事的目录结构是：

```text
TikTokDailyReport/
├── TikTokDailyReport_v<版本号>.exe
├── config/
│   ├── app_config.xlsx
│   ├── .env                     # 需要在线读取飞书时才放
│   └── cache/
├── data/
│   ├── data_JP/
│   │   ├── local/
│   │   ├── cross-border/
│   │   ├── direct_old/
│   │   └── direct_new/
│   └── data_VN/
│       ├── local/
│       └── cross_border/
└── result/
```

## app_config.xlsx

配置文件位置：

```text
sum_daily_order/config/app_config.xlsx
```

打包后放在：

```text
TikTokDailyReport/config/app_config.xlsx
```

`Stores` Sheet 字段说明：

```text
enabled       1=启用，0=跳过
country_code  JP=日本，VN=越南，MX=墨西哥
country_name  中文国家名
store_key     输出文件名使用
store_name    中文店铺名
store_dir     订单 CSV 所在文件夹名
sheet_name    飞书配置表 Sheet 名称
```

如果暂时不跑某个店铺，把 `enabled` 改成 `0` 即可。

## 本地运行

安装依赖：

```bash
pip install pandas requests python-dotenv openpyxl xlsxwriter
```

列出启用店铺：

```bash
python main.py --list
```

运行日本：

```bash
python main.py --site JP
```

运行越南：

```bash
python main.py --site VN
```

运行全部：

```bash
python main.py --site all
```

不带参数运行时，会出现菜单让你选择站点：

```bash
python main.py
```

## 数据放置

把 TikTok Shop 导出的订单 CSV 放进对应目录。

日本直邮老店示例：

```text
sum_daily_order/data/data_JP/direct_old/
```

日本直邮新店示例：

```text
sum_daily_order/data/data_JP/direct_new/
```

越南跨境店示例：

```text
sum_daily_order/data/data_VN/cross_border/
```

一个目录下可以放多个 CSV。程序会：

- 按 CSV 文件和日期生成日报汇总，避免多个文件日期重叠时重复混算
- 按 CSV 文件生成周期对比
- 在 `Daily Product Detail` 中按配置表 `产品大类` 生成产品级 GMV、成本、预估毛利汇总
- 在 `SKU Detail` 中保留 SKU/规格级明细，方便排查具体规格表现
- 在 `Product Quantity Matrix` 顶部生成按文件+产品的销量汇总
- 在 `Product Quantity Matrix` 下方生成 `文件名 + 产品名称 + 日期` 的每日销量矩阵
- 在 `Sample Statistics` 中生成 `文件名 + 产品名称 + 日期` 的每日样品单矩阵

## 飞书配置

程序会优先尝试使用本地缓存：

```text
config/cache/
```

GitHub Actions 自动打包时，默认只负责生成 exe 和交付包，不会把真实 `.env` 打进去：

```text
config/cache/
```

所以普通同事下载 Release 包后，只要把 `.env` 放到 `config/.env`，就可以在线读取飞书配置。

如果你自己需要在本机运行时在线刷新飞书配置，请在 `config/.env` 中配置飞书应用凭证：

```text
FEISHU_APP_ID=你的飞书应用ID
FEISHU_APP_SECRET=你的飞书应用Secret
```

如果没有 `.env`，但 `config/cache/` 中已有对应配置缓存，程序仍可以运行。

`.env` 是密钥文件，不要提交到 Git，不要写进代码，不要放进公开 Release，也不要发到公开群。Release 包默认不带真实 `.env`。

运行时规则：

- 有 `config/.env` 时，程序会优先从飞书刷新配置，并更新本地 `config/cache/`。
- 没有 `config/.env` 时，程序会使用 `config/cache/` 离线运行。

## Windows 打包

建议在 Windows 电脑上打包。PyInstaller 不是跨平台编译器，所以 Windows exe 最好在 Windows 上生成。

## 版本号管理

项目根目录有一个单独的版本文件：

```text
VERSION
```

当前内容示例：

```text
1.1.0
```

版本规则：

- 这里写 `1.1.0` 即可，不需要自己手动加 `v`
- 程序运行时会显示成 `v1.1.0`
- GitHub Actions 打包时会自动生成：
  - `TikTokDailyReport_v1.1.0.exe`
  - `TikTokDailyReport_v1.1.0_windows.zip`
  - GitHub Release 附件 `TikTokDailyReport_v1.1.0_windows.zip`

如果你要发新版本，最简单的方式就是先改这个文件，再提交代码。

## GitHub Actions 自动打包与上传

仓库已经预留了 GitHub Actions 工作流：

```text
.github/workflows/build-windows-exe.yml
```

它会在以下情况自动运行：

- 推送到 `master`
- 在 GitHub Actions 页面手动点击运行

自动完成的事情：

1. 校验 `main.py`、日本脚本、越南脚本语法
2. 在 Windows 环境打包 exe
3. 自动组装只包含日报工具的交付目录
4. 自动压缩成 zip 包
5. 上传到 Actions Artifact
6. 创建或更新 GitHub Release
7. 把 zip 包挂到 GitHub Release 的 Assets 附件里
8. 通过飞书机器人发送成功/失败通知

### Release 下载说明

自动打包后的交付包会出现在两个地方：

- GitHub Actions 的 `Artifacts`
- GitHub Release 的 `Assets`

交付包里只保留日报工具本身需要的文件，不再附带广告/商品卡/订单分析等旧模块目录。
交付包会包含 `config/.env.example`，但不会包含真实 `.env` 密钥文件。

建议平时给同事发飞书通知里的 Release 链接。打开 Release 页面后：

1. 找到页面下方的 `Assets`
2. 下载 `TikTokDailyReport_v<版本号>_windows.zip`
3. 不要下载 `Source code (zip)` 或 `Source code (tar.gz)`

`Source code` 是 GitHub 自动生成的源码包，不是给同事使用的工具包。

如果你希望在线读取飞书配置，请把 `.env` 放到 `config/.env`，内容格式参考 `config/.env.example`。

### 飞书机器人说明

飞书机器人链接不要写死在代码或工作流文件里。  
正确做法是放到 GitHub 仓库的 Actions Secrets 里，例如：

```text
FEISHU_BOT_WEBHOOK
```

如果你希望同事本地运行时在线读取飞书配置，就把飞书应用密钥放在本地的 `config/.env`，不要提交到 Git。

### 手动发版

如果你希望手动指定一个版本号再打包，也可以在 GitHub 的 Actions 页面里手动运行工作流，并填写：

```text
release_version
```

例如填写：

```text
1.1.1
```

或者：

```text
v1.1.1
```

都可以。

打包电脑需要先安装 Python，推荐安装：

```text
Python 3.11.x 64-bit
```

下载地址：

```text
https://www.python.org/downloads/windows/
```

安装时请勾选：

```text
Add python.exe to PATH
```

安装完成后，打开新的命令行窗口，检查：

```bat
py -3.11 --version
```

如果能看到类似下面的输出，就说明 Python 3.11 可用：

```text
Python 3.11.9
```

注意：只有“打包 exe 的电脑”需要安装 Python。最终拿到 `TikTokDailyReport_v<版本号>.exe` 的普通同事不需要安装 Python。

下面说的“项目文件夹根目录”，指的是你解压/拉取代码后的这个文件夹，例如：

```text
C:\Users\MAC\Desktop\tiktok_data_process
```

不是 `C:\`，也不是电脑系统根目录。

先在命令行进入项目文件夹根目录：

```bat
cd /d C:\Users\MAC\Desktop\tiktok_data_process
```

然后运行：

```bat
py -3.11 -m venv .venv
.venv\Scripts\activate
pip install pandas requests python-dotenv openpyxl xlsxwriter pyinstaller
pyinstaller --onefile --name TikTokDailyReport_v<版本号> main.py
```

打包完成后，exe 通常在：

```text
dist/TikTokDailyReport_v<版本号>.exe
```

然后在 `dist` 里补齐外部目录：

```text
dist/
├── TikTokDailyReport_v<版本号>.exe
├── config/
├── data/
└── result/
```

程序第一次启动时，如果发现没有：

```text
dist/config/app_config.xlsx
```

会自动生成一个默认配置文件。所以 `app_config.xlsx` 不一定必须手动复制，但建议你确认里面的 `enabled` 和店铺目录是否符合实际。

GitHub Actions 自动打包出来的 Release 包会自带：

```text
config/.env.example
```

普通同事如果要在线刷新飞书配置，需要自己准备 `config/.env`。

如果你自己希望本机或某个内部版本每次运行都能在线读取飞书配置，也可以单独准备：

```text
config/.env
```

内容格式：

```text
FEISHU_APP_ID=你的飞书应用ID
FEISHU_APP_SECRET=你的飞书应用Secret
```

注意：`.env` 里是飞书应用密钥，不建议放进 Release 包，也不要提交到 Git。普通同事如果没有 `.env`，就只能依赖 Release 包里已有的 `config/cache/`。

注意：`app_config.xlsx` 只是“运行哪些国家/店铺”的配置；SKU 成本、物流成本、寄样成本、产品大类来自飞书配置表或本地缓存。  
如果 exe 目录下没有 `config/.env`，就必须准备：

```text
dist/config/cache/
```

否则程序会提示无法读取 SKU 成本配置。

订单 CSV 不建议随 exe 打包，日常使用时直接放到 `dist/data/` 下对应店铺目录。

最终同事只需要：

1. 把订单 CSV 放进对应 `data` 子目录。
2. 双击 `TikTokDailyReport_v<版本号>.exe`。
3. 按菜单选择日本、越南或全部。
4. 到 `result` 目录查看生成的 Excel。

## macOS 快速部署

macOS 同事可以不打包，直接运行源码：

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install pandas requests python-dotenv openpyxl xlsxwriter
python main.py --site all
```

如果也要打包 macOS 可执行文件，在 Mac 上运行：

```bash
pip install pyinstaller
pyinstaller --onefile --name TikTokDailyReport_v<版本号> main.py
```

macOS 打包出来的是 macOS 可执行文件，不能给 Windows 直接使用。

## 常见问题

### 双击后提示找不到 app_config.xlsx

新版程序会自动生成默认配置。如果你仍看到这个提示，说明你运行的是旧 exe，请重新执行：

```bat
pyinstaller --onefile --name TikTokDailyReport_v<版本号> main.py
```

然后重新打开：

```text
dist/TikTokDailyReport_v<版本号>.exe
```

### 提示某个目录下未找到文件

确认订单 CSV 放在对应店铺目录。例如日本直邮老店和新店需要分开放：

```text
data/data_JP/direct_old/
data/data_JP/direct_new/
```

### 提示无法从飞书刷新配置

如果你希望在线读取飞书，请检查：

```text
config/.env
```

如果只是给普通同事使用，建议提前准备好 `config/.env`，让程序在线读取飞书。

打包后对应路径是：

```text
dist/config/.env
```

也就是把源码里的：

```text
config/.env
```

整个复制到：

```text
dist/config/.env
```

### pip install 最后出现一大段 Traceback

如果前面已经显示：

```text
Successfully installed ...
```

后面的 Traceback 很可能只是 pip 在检查自身更新版本时网络卡住，手动 `Ctrl+C` 中断了。通常不影响已经安装好的依赖。

为了减少这个提示，可以安装时加上：

```bat
pip install pandas requests python-dotenv openpyxl xlsxwriter pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple --disable-pip-version-check
```

### PyInstaller 打包时有 WARNING

例如：

```text
WARNING: Hidden import "jinja2" not found!
WARNING: Library not found: could not resolve 'VERSION.dll'
```

这类 warning 不一定代表失败。只要最后生成了：

```text
dist/TikTokDailyReport_v1.1.0.exe
```

并且双击能启动，就可以继续测试。真正需要处理的是运行 exe 后出现的业务报错，例如配置文件、数据目录、飞书缓存缺失等。
