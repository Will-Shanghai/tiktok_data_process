import pandas as pd
import os

# 文件目录路径
directory_path = './video_data/'

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

    # 计算点击率（视频）_sum/7 和 点击成交转化率（视频）_sum/7
    click_rate_per_7 = df['点击率（视频）'].sum() / 7
    conversion_rate_per_7 = df['点击成交转化率（视频）'].sum() / 7

    # 计算视频转化率
    video_conversion_rate = (conversion_rate_per_7 / click_rate_per_7) * 100

    # 汇总相关列，并转换为普通的数字类型
    summary = {
        '视频订单': int(df['SKU 订单数'].sum()),
        '视频销售额': int(df['商品交易总额（视频） (₫)'].sum()),
        '视频曝光': int(df['视频vv'].sum()),
        '视频点击率': f"{click_rate_per_7 * 100:.2f}%",
        '视频转化率': f"{video_conversion_rate:.2f}%"
    }

    print(summary, '\n')
