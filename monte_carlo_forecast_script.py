import pandas as pd
import numpy as np
from scipy.stats import norm
import openpyxl
from datetime import datetime
import os


def monte_carlo_forecast(
    file_path, num_simulations, distribution_type, years_to_predict=10
):
    df = pd.read_excel(file_path)
    last_year = df.iloc[-1, 0]

    # 创建一个仅包含预测年份的DataFrame
    predicted_df = pd.DataFrame(
        {"Year": [last_year + i for i in range(1, years_to_predict + 1)]}
    )

    # 遍历每个财务指标列
    for column in df.columns[1:]:  # 跳过第一列（年份）
        simulated_data = {last_year + i: [] for i in range(1, years_to_predict + 1)}

        # 对每列执行蒙特卡洛模拟
        for _ in range(num_simulations):
            if distribution_type == "normal":
                mean, std = df[column].mean(), df[column].std()
                simulated_values = norm.rvs(loc=mean, scale=std, size=years_to_predict)
            elif distribution_type == "uniform":
                min_val, max_val = df[column].min(), df[column].max()
                simulated_values = np.random.uniform(
                    low=min_val, high=max_val, size=years_to_predict
                )

            # 将模拟值添加到对应的年份
            for year, value in zip(simulated_data.keys(), simulated_values):
                simulated_data[year].append(value)

        # 对每个年份找到最接近平均值的数据点
        closest_to_mean_data = [
            min(values, key=lambda x: abs(x - np.mean(values)))
            for values in simulated_data.values()
        ]

        # 将最接近平均值的数据添加到predicted_df中
        predicted_df[column] = closest_to_mean_data

    # 获取当前日期时间
    current_time = datetime.now().strftime("%Y%m%d%H%M%S")

    # 将原始数据和预测数据合并，并保持年份列在前
    final_df = pd.concat([df, predicted_df], ignore_index=True)

    # 获取脚本的当前路径
    script_directory = os.path.dirname(os.path.realpath(__file__))

    # 创建目标文件夹的路径
    target_directory = os.path.join(
        script_directory, "forecast", "monte_carlo_forecast"
    )

    # 检查目标文件夹是否存在，如果不存在，则创建它
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)

    # 获取当前日期时间
    current_time = datetime.now().strftime("%Y%m%d%H%M%S")

    # 设置输出文件的路径
    output_file = os.path.join(
        target_directory,
        f"@{num_simulations}T@{distribution_type}@{years_to_predict}Y-{current_time}.xlsx",
    )
    # 将结果保存到带有当前日期时间的新Excel文件
    final_df.to_excel(output_file, index=False)

    return output_file


# 用户输入
print("若采用默认值请直接回车")
file_path = input("请输入源数据路径（默认值：data.xlsx）: \n") or "data.xlsx"
num_simulations = (input("请输入模拟次数(默认值：2000次): \n")) or "2000"
num_simulations = int(num_simulations)
distribution_type = (
    input("请输入采用的数据分布 (输入normal表示正态分布 ， 输入uniform表示均匀分布，默认值为normal): \n") or "normal"
)
years_to_predict = (input("请输入需要预测的年数 (默认值：10年): ")) or "10"
years_to_predict = int(years_to_predict)
output_file = monte_carlo_forecast(
    file_path, num_simulations, distribution_type, years_to_predict
)

print(f"蒙特卡洛预测完成，数据已保存在 {output_file}")
