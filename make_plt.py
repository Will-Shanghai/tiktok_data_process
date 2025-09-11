import matplotlib.pyplot as plt

# 数据
weeks = range(1, 14)  # 周次1~13
profit_rates = [-42.92, -30.80, -46.56, -40.44, -36.05, -28.08, -22.03, -20.29, -18.54, -29.71, -26.47, -29.85, -34.76]

# 绘制散点图
plt.scatter(weeks, profit_rates, color='blue', alpha=0.7)  # 散点样式
plt.xlabel('周次（从2025/06/02开始的第n周）')  # x轴标签
plt.ylabel('利润率（%）')  # y轴标签
plt.title('产品利润率变化散点图')  # 图表标题
plt.grid(True, linestyle='--', alpha=0.5)  # 显示网格线
plt.show()