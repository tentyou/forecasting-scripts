# forecasting-scripts
该项目运用多种方式对未来数据进行评估，可以通过命令行窗口接收用户数据并生成预测数据表，方便用户使用。
该项目包括了三种预测方法：
+ 蒙特卡洛模拟
+ 简单线性回归
+ 多项式回归
## 使用教程
1. 下载[Source Code](https://github.com/tentyou/forecasting-scripts/releases)，将压缩包解压至合适位置
2. 安装Python，注意在安装时选择将Python加入系统环境中（Add Python to environment variables）
3. 右键当前目录，选择在当前目录打开终端/CMD/Powershell，输入：```pip install -r forecasting_requirements.txt```
4. 安装完依赖库后，即可使用脚本
	在当前路径下打开终端：输入```python <脚本名称>```

例如：
	```python .\linear_regression_forecast_script.py```

按照命令行提示提供参数，随后脚本会自动在路径下生成预测文件


## 注意
+ 建议将预测数据放置至脚本当前目录下，并命名为```data.xlsx```

+ 格式参照```data_exaple.xlsx```，第一行为标题，第一列为年份，之后为需要预测的数据
+ 请保证数据均为正确填写的**数字**

**※※※A1单元格务必填"Year"，否则程序报错※※※**

**※※※A1单元格务必填"Year"，否则程序报错※※※**

**※※※A1单元格务必填"Year"，否则程序报错※※※**
