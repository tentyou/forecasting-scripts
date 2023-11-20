import pandas as pd
from sklearn.linear_model import LinearRegression
from datetime import datetime
import numpy as np
import os


def linear_regression_forecast(file_path, years_to_predict):
    df = pd.read_excel(file_path)
    last_year = df.iloc[-1, 0]

    # 创建一个仅包含预测年份的DataFrame
    predicted_years = np.arange(last_year + 1, last_year + years_to_predict + 1)
    predicted_df = pd.DataFrame(predicted_years, columns=["Year"])  # 确保列名与训练数据一致

    # 对每个财务指标进行线性回归
    for column in df.columns[1:]:  # 跳过第一列（年份）
        X = df[["Year"]]  # Predictor variable, 使用DataFrame格式
        y = df[column]  # Dependent variable

        # 创建并训练线性回归模型
        model = LinearRegression()
        model.fit(X, y)

        # 进行预测
        predicted_values = model.predict(predicted_df[["Year"]])  # 确保预测数据的格式与训练数据一致
        predicted_df[column] = predicted_values

    # 获取当前日期时间
    current_time = datetime.now().strftime("%Y%m%d%H%M%S")

    # 将原始数据和预测数据合并，并保持年份列在前
    final_df = pd.concat([df, predicted_df], ignore_index=True)

    # 获取脚本的当前路径
    script_directory = os.path.dirname(os.path.realpath(__file__))

    # 创建目标文件夹的路径
    target_directory = os.path.join(
        script_directory, "forecast", "linear_regression_forecast"
    )

    # 检查目标文件夹是否存在，如果不存在，则创建它
    if not os.path.exists(target_directory):
        os.makedirs(target_directory)

    # 获取当前日期时间
    current_time = datetime.now().strftime("%Y%m%d%H%M%S")

    # 设置输出文件的路径
    output_file = os.path.join(
        target_directory, f"@{years_to_predict}Y-{current_time}.xlsx"
    )
    # 将结果保存到带有当前日期时间的新Excel文件
    final_df.to_excel(output_file, index=False)
    return output_file


# 获取用户输入
print("若采用默认值请直接回车")
file_path = input("请输入源数据路径（默认值：data.xlsx）: \n") or "data.xlsx"
years_to_predict = (input("请输入需要预测的年数 (默认值：10年): ")) or "10"
years_to_predict = int(years_to_predict)

# 执行预测
output_file = linear_regression_forecast(file_path, years_to_predict)
print(f"线性回归预测完成。数据已保存在 {output_file}。")
