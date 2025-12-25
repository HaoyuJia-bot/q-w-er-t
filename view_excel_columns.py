import pandas as pd
import os

# 检查文件是否存在
file_path = '两版合并后的年报数据_完整版.xlsx'
if os.path.exists(file_path):
    print(f"文件存在: {file_path}")
    
    # 读取Excel文件的前10行数据
    df = pd.read_excel(file_path, nrows=10)
    
    # 显示所有列名
    print("\nExcel文件的列名:")
    for i, col in enumerate(df.columns, 1):
        print(f"{i}. {col}")
    
    # 显示数据类型
    print("\n各列的数据类型:")
    print(df.dtypes)
    
    # 显示前5行数据
    print("\n前5行数据预览:")
    print(df.head())
    
    # 显示基本统计信息
    print("\n数据基本统计信息:")
    print(df.describe(include='all'))
    
    # 检查是否有'企业名称'列
    if '企业名称' in df.columns:
        print("\n企业数量:", df['企业名称'].nunique())
    
    # 检查是否有'年份'列
    if '年份' in df.columns:
        print("\n年份范围:", df['年份'].min(), "-", df['年份'].max())
    
    # 检查是否有包含'数字化'、'转型'或'指数'的列
    index_columns = [col for col in df.columns if any(keyword in col for keyword in ['数字化', '转型', '指数'])]
    if index_columns:
        print("\n数字化转型相关列:", index_columns)
    
    # 检查是否有行业相关列
    industry_columns = [col for col in df.columns if any(keyword in col for keyword in ['行业', '产业', '所属行业'])]
    if industry_columns:
        print("\n行业相关列:", industry_columns)
        print("行业类别:", df[industry_columns[0]].unique())
    
    # 检查是否有地区相关列
    region_columns = [col for col in df.columns if any(keyword in col for keyword in ['地区', '省份', '城市', '地域', '区域'])]
    if region_columns:
        print("\n地区相关列:", region_columns)
        print("地区类别:", df[region_columns[0]].unique()[:10])  # 只显示前10个
else:
    print(f"文件不存在: {file_path}")
    print("当前工作目录:", os.getcwd())
    print("当前目录下的文件:", os.listdir('.'))