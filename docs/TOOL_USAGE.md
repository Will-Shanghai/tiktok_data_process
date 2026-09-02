# TikTok 数据处理工具使用说明

## 新目录结构

```text
tiktok_data_process/
  app.py                         # 统一入口，后续打包 EXE 优先使用它
  config/
    tool_tasks.csv               # 功能任务开关和脚本路径配置
  tiktok_data_tool/
    cli.py                       # 菜单和命令行参数
    tasks.py                     # 读取配置、切换目录、执行脚本
  cal_ads_data/                  # 计算广告数据
  cal_product_data/              # 计算商品卡数据
  click_order_data/              # 计算达人点击订单数据
  sum_daily_order/               # 汇总每日订单
```

这次调整没有移动原来的数据和结果目录，旧脚本仍然可以单独运行。

## 运行方式

在项目根目录执行：

```bash
python app.py
```

也可以直接执行某个任务：

```bash
python app.py ads
python app.py product_card
python app.py click_order
python app.py daily_order
python app.py all
```

查看当前启用的任务：

```bash
python app.py --list
```

## 配置任务

任务配置在 `config/tool_tasks.csv`：

```csv
enabled,task_id,task_name,script_path,workdir
1,ads,计算广告数据,cal_ads_data/config/ads_statistics.py,cal_ads_data/config
```

- `enabled`: `1` 表示启用；改成 `0` 就不会出现在菜单里。
- `task_id`: 命令行使用的英文短名称，比如 `daily_order`。
- `task_name`: 菜单里展示的中文名称。
- `script_path`: 要执行的脚本路径，从项目根目录开始写。
- `workdir`: 执行脚本前切换到哪个目录。旧脚本如果用了 `../data/`，这里要填脚本所在的 `config` 目录。

## 后续打包建议

后续打包 EXE 时优先打包 `app.py`。建议先保持 `data/`、`result/`、`.env`、CSV 配置文件在外部目录，这样同事替换数据和配置时不用重新打包。
