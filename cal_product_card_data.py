import pandas as pd
import os

# 文件目录路径
directory_path = './product_data/'

# 获取目录下所有文件
file_list = [f for f in os.listdir(directory_path) if f.endswith('.xlsx')]

# 遍历文件列表，循环执行处理
for file_name in file_list:
    # 拼接文件路径
    file_path = os.path.join(directory_path, file_name)

    # 读取文件，跳过前两行，将第三行作为列名
    df = pd.read_excel(file_path, skiprows=2)

    # 打印第一行第一列的值
    print(f"文件: {file_name}，数据时间起点:", df.iloc[0, 0])

    # 清理并转换为数值类型，防止数据类型错误
    df['曝光次数'] = pd.to_numeric(df['曝光次数'], errors='coerce')
    df['SKU订单数'] = pd.to_numeric(df['SKU订单数'], errors='coerce')
    df['GMV (₫)'] = pd.to_numeric(df['GMV (₫)'], errors='coerce')

    # 处理百分比列：去掉百分号并转换为浮动数值
    df['曝光到点击转化率'] = df['曝光到点击转化率'].replace('%', '', regex=True).astype(float) / 100
    df['点击到成交转化率'] = df['点击到成交转化率'].replace('%', '', regex=True).astype(float) / 100

    # 如果有NaN值，填充为0
    df.fillna(0, inplace=True)

    # 求和相关列，并进行必要的转换
    exposure_sum = df['曝光次数'].sum()  # B列：曝光次数
    sku_order_sum = df['SKU订单数'].sum()  # I列：SKU订单数
    gmv_sum = df['GMV (₫)'].sum()  # J列：GMV (₫)
    exposure_to_click_conversion_sum = df['曝光到点击转化率'].sum() / 7  # M列：曝光到点击转化率（除以7）
    click_to_deal_conversion_sum = df['点击到成交转化率'].sum() / 7  # O列：点击到成交转化率（除以7）

    # 汇总并打印结果
    summary = {
        '商品卡订单数': int(sku_order_sum),
        '商品卡销售额': int(gmv_sum),
        '商品卡曝光': int(exposure_sum),
        '商品卡点击率': f"{exposure_to_click_conversion_sum * 100:.2f}%",
        '商品卡转化率': f"{click_to_deal_conversion_sum * 100:.2f}%"
    }

    print(summary, '\n')
