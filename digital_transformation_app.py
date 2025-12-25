import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
import matplotlib.font_manager as fm
import numpy as np
import seaborn as sns

# 设置字体以避免中文渲染问题
import warnings
warnings.filterwarnings('ignore')

# 使用安全的字体设置
try:
    # 在macOS上尝试使用系统字体
    plt.rcParams['font.family'] = ['DejaVu Sans', 'Arial', 'sans-serif']
    plt.rcParams['axes.unicode_minus'] = False
    sns.set_style("whitegrid")
except:
    # 如果设置失败，使用默认字体
    plt.rcParams['font.family'] = ['sans-serif']
    plt.rcParams['axes.unicode_minus'] = False

# 设置白色主题
st.set_page_config(
    page_title='企业数字化转型指数查询系统',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='expanded'
)

# 加载Excel数据
def load_data():
    try:
        # 检查文件是否存在
        file_path = '两版合并后的年报数据_完整版.xlsx'
        if not os.path.exists(file_path):
            st.error(f"文件不存在: {file_path}")
            st.write("当前工作目录:", os.getcwd())
            st.write("当前目录下的文件:", os.listdir('.'))
            return None
        
        # 尝试使用openpyxl引擎读取Excel文件
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
        except Exception as openpyxl_error:
            # 如果openpyxl失败，尝试xlrd引擎
            try:
                df = pd.read_excel(file_path, engine='xlrd')
            except Exception as xlrd_error:
                # 如果都失败，尝试不指定引擎（使用默认）
                df = pd.read_excel(file_path)
        
        return df
    except Exception as e:
        st.error(f"加载数据失败: {e}")
        st.write("当前工作目录:", os.getcwd())
        st.write("当前目录下的文件:", os.listdir('.'))
        return None

# 加载数据
df = load_data()

if df is not None:
    # 检查必要列是否存在
    required_columns = ['股票代码', '年份', '企业名称']
    index_columns = [col for col in df.columns if '数字化' in col or '转型' in col or '指数' in col]
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        st.error(f"缺少必要的列: {', '.join(missing_columns)}")
        st.stop()
    
    if not index_columns:
        st.warning("未找到包含'数字化'、'转型'或'指数'的列，请检查数据")
        st.stop()
    
    # 获取唯一的股票代码和年份
    stock_codes = df['股票代码'].unique().tolist()
    years = df['年份'].unique().tolist()
    
    # 排序
    years.sort()
    
    # 页面标题
    st.title('企业数字化转型指数查询系统')
    
    # 使用两列布局：左边查询面板，右边内容
    query_col, content_col = st.columns([1, 4])
    
    # 左边查询面板
    with query_col:
        st.header('查询面板')
        st.write('请选择以下参数进行查询')
        
        selected_stock = st.selectbox('股票代码', stock_codes, key='stock_select')
        selected_year = st.selectbox('年份', years, key='year_select')
        
        # 查询按钮
        search_button = st.button('查询', key='search_button', help='点击查询数据')
        
        # 显示选中股票的基本信息
        stock_data = df[df['股票代码'] == selected_stock]
        if not stock_data.empty:
            company_name = stock_data['企业名称'].iloc[0]
            st.markdown('---')
            st.subheader('企业信息')
            st.write(f"**企业名称**: {company_name}")
            st.write(f"**股票代码**: {selected_stock}")
            st.write(f"**数据年份**: {', '.join(map(str, sorted(stock_data['年份'].unique())))}")
    
    # 计算统计指标（在主内容区域外）
    index_col = index_columns[0]
    avg_index = df[index_col].mean() if index_col in df.columns else 0
    max_index = df[index_col].max() if index_col in df.columns else 0
    min_index = df[index_col].min() if index_col in df.columns else 0
    median_index = df[index_col].median() if index_col in df.columns else 0
    std_index = df[index_col].std() if index_col in df.columns else 0
    
    # 1. 统计概览（在数据概览之前）
    st.subheader('统计概览')
    
    # 使用两列布局分两排显示统计信息
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric(label="总记录数", value=len(df))
        st.metric(label="企业数量", value=len(df['股票代码'].unique()))
        st.metric(label="年份范围", value=f"{min(years)}-{max(years)}")
        st.metric(label="平均指数", value=f"{avg_index:.2f}")
    
    with col2:
        st.metric(label="最高指数", value=f"{max_index:.2f}")
        st.metric(label="最低指数", value=f"{min_index:.2f}")
        st.metric(label="中位数指数", value=f"{median_index:.2f}")
        st.metric(label="指数标准差", value=f"{std_index:.2f}")
        st.metric(label="数据年份数", value=len(years))
    
    # 右边主内容区域
    with content_col:
        
        st.markdown('---')
        
        # 2. 数据概览
        st.subheader('数据概览')
        st.dataframe(df.sample(10))
        st.write("**数据结构**")
        st.write(f"行数: {df.shape[0]}")
        st.write(f"列数: {df.shape[1]}")
        st.write(f"\n**主要列名**")
        st.write("\n".join(df.columns[:10]))
        if len(df.columns) > 10:
            st.write(f"... 等 {len(df.columns)} 列")
        
        st.markdown('---')
        
        # 3. 数字化转型指数分布
        st.subheader('数字化转型指数分布')
        if index_col in df.columns:
            # 根据选择的股票代码筛选数据
            if search_button:
                filtered_df = df[df['股票代码'] == selected_stock]
                if not filtered_df.empty:
                    company_name = filtered_df['企业名称'].iloc[0]
                    display_df = filtered_df
                    title_suffix = f'- {company_name}({selected_stock})'
                else:
                    display_df = df
                    title_suffix = ''
            else:
                display_df = df
                title_suffix = ''
            
            # 直方图和密度图
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
            
            # 直方图
            ax1.hist(display_df[index_col], bins=20, alpha=0.7, color='#1f77b4')
            ax1.set_title(f'{index_col}分布直方图{title_suffix}')
            ax1.set_xlabel(index_col)
            ax1.set_ylabel('企业数量')
            ax1.grid(True, alpha=0.3)
            
            # 密度图
            sns.kdeplot(display_df[index_col], ax=ax2, fill=True, color='#ff7f0e', alpha=0.7)
            ax2.set_title(f'{index_col}分布密度图{title_suffix}')
            ax2.set_xlabel(index_col)
            ax2.set_ylabel('密度')
            ax2.grid(True, alpha=0.3)
            
            plt.tight_layout()
            st.pyplot(fig)
        else:
            st.info(f"未找到{index_col}列，无法生成指数分布")
        
        st.markdown('---')
        
        # 第四个：数字化转型指数详细统计
        st.subheader('数字化转型指数详细统计')
        if index_col in df.columns:
            index_stats = {
                '平均值': df[index_col].mean(),
                '中位数': df[index_col].median(),
                '标准差': df[index_col].std(),
                '最小值': df[index_col].min(),
                '最大值': df[index_col].max(),
                '25%分位数': df[index_col].quantile(0.25),
                '75%分位数': df[index_col].quantile(0.75)
            }
            
            for stat_name, value in index_stats.items():
                st.info(f"**{stat_name}**\n{value:.4f}")
        
        st.markdown('---')
        
        # 第五个：数字化转型的折线图
        st.subheader('数字化转型指数折线图')
        if index_col in stock_data.columns and not stock_data.empty:
            # 按年份排序
            stock_data_sorted = stock_data.sort_values('年份')
            
            plt.figure(figsize=(12, 6))
            plt.plot(stock_data_sorted['年份'], stock_data_sorted[index_col], 
                    marker='o', linewidth=2, markersize=8, color='#2E86C1')
            plt.title(f'{company_name}({selected_stock})历年数字化转型指数变化', fontsize=14)
            plt.xlabel('年份')
            plt.ylabel(index_col)
            plt.grid(True, alpha=0.3)
            plt.xticks(stock_data_sorted['年份'])
            
            # 添加数值标签
            for x, y in zip(stock_data_sorted['年份'], stock_data_sorted[index_col]):
                plt.annotate(f'{y:.2f}', (x, y), textcoords="offset points", 
                           xytext=(0,10), ha='center', fontsize=10)
            
            plt.tight_layout()
            st.pyplot(plt)
        else:
            st.info("暂无足够数据生成折线图")
        
        st.markdown('---')
        
        # 第六个：整体数字化转型趋势折线图
        st.subheader('整体数字化转型趋势折线图')
        
        if index_col in df.columns:
            # 计算每年的平均指数
            yearly_avg = df.groupby('年份')[index_col].mean().reset_index()
            
            plt.figure(figsize=(14, 8))
            plt.plot(yearly_avg['年份'], yearly_avg[index_col], 
                    marker='o', linewidth=3, markersize=10, color='#E74C3C')
            
            plt.title('历年企业数字化转型指数整体趋势', fontsize=16, fontweight='bold')
            plt.xlabel('年份', fontsize=14)
            plt.ylabel(f'平均{index_col}', fontsize=14)
            plt.grid(True, alpha=0.3)
            plt.xticks(yearly_avg['年份'])
            
            # 添加数值标签
            for x, y in zip(yearly_avg['年份'], yearly_avg[index_col]):
                plt.annotate(f'{y:.2f}', (x, y), textcoords="offset points", 
                           xytext=(0,15), ha='center', fontsize=12, fontweight='bold')
            
            # 添加趋势线
            z = np.polyfit(yearly_avg['年份'], yearly_avg[index_col], 1)
            p = np.poly1d(z)
            plt.plot(yearly_avg['年份'], p(yearly_avg['年份']), "--", alpha=0.7, color='#2C3E50', linewidth=2)
            
            plt.tight_layout()
            st.pyplot(plt)
            
            # 显示趋势分析
            st.write("**趋势分析**")
            if z[0] > 0:
                st.success(f"📈 整体呈上升趋势，年均增长约 {z[0]:.3f} 个单位")
            elif z[0] < 0:
                st.warning(f"📉 整体呈下降趋势，年均下降约 {abs(z[0]):.3f} 个单位")
            else:
                st.info("📊 整体趋势相对平稳")
        else:
            st.info(f"未找到{index_col}列，无法生成整体趋势图")
        

    

    # 查询结果
    if search_button or True:  # 默认显示所有数据
        st.markdown('---')
        st.header('查询结果')
        
        # 按股票代码过滤
        stock_data = df[df['股票代码'] == selected_stock]
        
        # 显示该股票的基本信息
        if not stock_data.empty:
            # 获取企业名称
            company_name = stock_data['企业名称'].iloc[0]
            
            # 公司信息卡片
            with st.container():
                st.subheader('公司基本信息')
                st.info(f"**企业名称**\n{company_name}")
                st.info(f"**股票代码**\n{selected_stock}")
                st.info(f"**数据年份**\n{', '.join(map(str, sorted(stock_data['年份'].unique())))}")
            
            # 获取指定年份的数据
            year_data = stock_data[stock_data['年份'] == selected_year]
            if not year_data.empty:
                # 使用动态检测到的索引列
                index_col = index_columns[0]
                if index_col in year_data.columns:
                    index_value = year_data[index_col].iloc[0]
                    
                    # 指数展示卡片
                    with st.container():
                        st.subheader(f'{selected_year}年数字化转型指数')
                        st.metric(
                            label=f"{selected_year}年{index_col}", 
                            value=f"{index_value:.2f}" if isinstance(index_value, (int, float)) else index_value,
                            delta=None
                        )
                else:
                    st.warning(f"未找到{index_col}列")
            else:
                st.warning(f"未找到{selected_stock}在{selected_year}年的数据")
        else:
            st.warning(f"未找到股票代码{selected_stock}的数据")
        
        # 可视化部分
        st.markdown('---')
        st.header('数据可视化')
        
        # 使用动态检测到的索引列
        index_col = index_columns[0]
        
        if index_col in stock_data.columns:
            # 按年份排序
            stock_data_sorted = stock_data.sort_values('年份')
            
            # 折线图和柱状图并排显示
            st.subheader('历年指数折线图')
            plt.figure(figsize=(10, 6))
            plt.plot(stock_data_sorted['年份'], stock_data_sorted[index_col], marker='o', linestyle='-', color='#1f77b4')
            plt.title(f'{company_name}({selected_stock})历年{index_col}')
            plt.xlabel('年份')
            plt.ylabel(index_col)
            plt.grid(True, alpha=0.3)
            plt.tight_layout()
            st.pyplot(plt)
            
            st.subheader('历年指数柱状图')
            plt.figure(figsize=(10, 6))
            bars = plt.bar(stock_data_sorted['年份'], stock_data_sorted[index_col], color='#ff7f0e', alpha=0.8)
            plt.title(f'{company_name}({selected_stock})历年{index_col}')
            plt.xlabel('年份')
            plt.ylabel(index_col)
            plt.grid(True, alpha=0.3, axis='y')
            
            # 在柱状图上显示数值
            for bar in bars:
                height = bar.get_height()
                plt.text(bar.get_x() + bar.get_width()/2., height + 0.01 * max(stock_data_sorted[index_col]),
                        f'{height:.2f}', ha='center', va='bottom')
            
            plt.tight_layout()
            st.pyplot(plt)
        else:
            st.warning(f"未找到{index_col}列，无法生成趋势图")
        
        # 数据表格
        st.markdown('---')
        st.header('详细数据')
        st.dataframe(stock_data)
        
        # 提供下载功能
        st.markdown('---')
        st.header('数据下载')
        csv = stock_data.to_csv(index=False)
        st.download_button(
            label="下载当前股票数据 (CSV)",
            data=csv,
            file_name=f"{company_name}_{selected_stock}_数字化转型数据.csv",
            mime="text/csv"
        )

else:
    st.error("数据加载失败，请检查Excel文件")