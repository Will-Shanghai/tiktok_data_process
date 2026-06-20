# TikTok Data Process

这个项目用于生成 TikTok Shop 订单成本日报/周期报表。当前重点维护：

- 日本站日报：`sum_daily_order/config/cal_cost_return_jp_daily.py`
- 越南站日报：`sum_daily_order/config/cal_cost_return_vn_daily.py`
- 打包友好的统一入口：`main.py`

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
│   │   └── direct/
│   └── data_VN/
│       ├── local/
│       └── cross_border/
└── result/                      # 报表输出目录
```

Windows 打包后，建议给同事的目录结构是：

```text
TikTokDailyReport/
├── TikTokDailyReport.exe
├── config/
│   ├── app_config.xlsx
│   ├── .env                     # 需要在线读取飞书时才放
│   └── cache/
├── data/
│   ├── data_JP/
│   │   ├── local/
│   │   ├── cross-border/
│   │   └── direct/
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
country_code  JP=日本，VN=越南
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

日本直邮店示例：

```text
sum_daily_order/data/data_JP/direct/
```

越南跨境店示例：

```text
sum_daily_order/data/data_VN/cross_border/
```

一个目录下可以放多个 CSV。程序会：

- 汇总所有 CSV 生成总览日报
- 按 CSV 文件生成周期对比
- 在 `Daily Product Detail` 顶部生成按文件+产品的 GMV、成本、利润汇总
- 在 `Product Quantity Matrix` 顶部生成按文件+产品的销量汇总

## 飞书配置

程序会优先尝试使用本地缓存：

```text
config/cache/
```

如果需要在线刷新飞书配置，请在 `config/.env` 中配置：

```text
FEISHU_APP_ID=你的飞书应用ID
FEISHU_APP_SECRET=你的飞书应用Secret
```

如果没有 `.env`，但 `config/cache/` 中已有对应配置缓存，程序仍可以运行。

## Windows 打包

建议在 Windows 电脑上打包。PyInstaller 不是跨平台编译器，所以 Windows exe 最好在 Windows 上生成。

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

注意：只有“打包 exe 的电脑”需要安装 Python。最终拿到 `TikTokDailyReport.exe` 的普通同事不需要安装 Python。

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
pyinstaller --onefile --name TikTokDailyReport main.py
```

打包完成后，exe 通常在：

```text
dist/TikTokDailyReport.exe
```

然后在 `dist` 里补齐外部目录：

```text
dist/
├── TikTokDailyReport.exe
├── config/
├── data/
└── result/
```

程序第一次启动时，如果发现没有：

```text
dist/config/app_config.xlsx
```

会自动生成一个默认配置文件。所以 `app_config.xlsx` 不一定必须手动复制，但建议你确认里面的 `enabled` 和店铺目录是否符合实际。

如果同事需要离线使用飞书配置缓存，仍建议把源码里的缓存复制过去：

```text
sum_daily_order/config/cache/
```

订单 CSV 不建议随 exe 打包，日常使用时直接放到 `dist/data/` 下对应店铺目录。

最终同事只需要：

1. 把订单 CSV 放进对应 `data` 子目录。
2. 双击 `TikTokDailyReport.exe`。
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
pyinstaller --onefile --name TikTokDailyReport main.py
```

macOS 打包出来的是 macOS 可执行文件，不能给 Windows 直接使用。

## 常见问题

### 双击后提示找不到 app_config.xlsx

新版程序会自动生成默认配置。如果你仍看到这个提示，说明你运行的是旧 exe，请重新执行：

```bat
pyinstaller --onefile --name TikTokDailyReport main.py
```

然后重新打开：

```text
dist/TikTokDailyReport.exe
```

### 提示某个目录下未找到文件

确认订单 CSV 放在对应店铺目录。例如日本直邮店：

```text
data/data_JP/direct/
```

### 提示无法从飞书刷新配置

如果你希望在线读取飞书，请检查：

```text
config/.env
```

如果只是给普通同事使用，可以提前准备好 `config/cache/`，让程序使用本地缓存。

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
dist/TikTokDailyReport.exe
```

并且双击能启动，就可以继续测试。真正需要处理的是运行 exe 后出现的业务报错，例如配置文件、数据目录、飞书缓存缺失等。
