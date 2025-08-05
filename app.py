import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import yfinance as yf
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# 页面配置
st.set_page_config(
    page_title="AGIX Fund Monitor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #1f77b4;
    }
    .positive {
        color: #28a745;
    }
    .negative {
        color: #dc3545;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    """加载所有数据文件"""
    try:
        # 设置pandas显示选项
        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)
        # 加载价格数据
        stockprice = pd.read_excel(
            "ciq reference/data/StockPriceValue.xlsx", 
            header=0, 
            index_col=0, 
            sheet_name="Price", 
            parse_dates=['Date']
        )
        
        # 确保索引是日期类型
        if stockprice.index.name == 'Date':
            stockprice.index = pd.to_datetime(stockprice.index, errors='coerce')
        
        stockprice = stockprice.replace(0, np.nan)
        filled_pri = stockprice.bfill()
        
        # 确保所有数值列都是数值类型
        for col in filled_pri.columns:
            if filled_pri[col].dtype == 'object':
                filled_pri[col] = pd.to_numeric(filled_pri[col], errors='coerce')
        
        # 加载持仓数据
        shares = pd.read_excel(
            "ciq reference/data/Shares.xlsx", 
            header=0, 
            index_col=0, 
            sheet_name="shares"
        )
        
        # 确保所有数值列都是数值类型
        for col in shares.columns:
            if shares[col].dtype == 'object':
                shares[col] = pd.to_numeric(shares[col], errors='coerce')
        
        # 加载每日监控数据 - 跳过前3行无用信息，并过滤掉底部的日期行
        daily_monitor = pd.read_excel(
            "ciq reference/data/每日数据监控.xlsx",
            sheet_name="chart",
            header=0,  # 第1行作为列标题
            skiprows=[1, 2],  # 跳过第2-3行无用信息
            index_col=0,
            parse_dates=True
        )
        
        # 加载raw1表数据 - A1:I9范围
        try:
            raw1_data = pd.read_excel(
                "ciq reference/data/每日数据监控.xlsx",
                sheet_name="raw1",
                header=0,  # 第1行作为列名
                usecols="A:I",  # 只读取A-I列
                nrows=9  # 只读取前9行
            )
            
            # 过滤掉Name是"Heng Sheng Tech Index"的行
            if 'Name' in raw1_data.columns:
                raw1_data = raw1_data[raw1_data['Name'] != 'Heng Sheng Tech Index']
            
            # 过滤掉Name是"Nasdaq"或包含"Nasdaq"的行
            if 'Name' in raw1_data.columns:
                raw1_data = raw1_data[~raw1_data['Name'].str.contains('Nasdaq', case=False, na=False)]
            
            # 删除Ticker列
            if 'Ticker' in raw1_data.columns:
                raw1_data = raw1_data.drop(columns=['Ticker'])
            
        except Exception as e:
            st.warning(f"无法加载raw1表数据: {e}")
            raw1_data = None
        
        # 确保索引是日期类型
        if daily_monitor.index.name == 'Date':
            daily_monitor.index = pd.to_datetime(daily_monitor.index, errors='coerce')
        
        # 过滤掉底部的日期行（这些行的所有数据列都是None）
        # 找到第一个所有列都是None的行，然后只保留该行之前的数据
        none_mask = daily_monitor.isna().all(axis=1)
        if none_mask.any():
            first_none_row = none_mask.idxmax()
            daily_monitor = daily_monitor.loc[:first_none_row].iloc[:-1]  # 排除第一个全None的行
        
        # 确保所有数值列都是数值类型（排除非数值列如Capital IQ Ticker）
        numeric_columns = ['Fund Asset(MLN USD)', 'Volume(MLN)', 'Price Change', 'Daily Flow', 
                          'YTD Flow', 'Flow since Jan 2023', 'Expense Ratio', 'Return since 2024', 
                          'Return since 2025', '30D Vol', 'Holdings']
        
        for col in daily_monitor.columns:
            if col in numeric_columns and daily_monitor[col].dtype == 'object':
                daily_monitor[col] = pd.to_numeric(daily_monitor[col], errors='coerce')
        
        # 加载持仓信息
        holdings = pd.read_excel(
            "ciq reference/data/Shares.xlsx",
            sheet_name="07_30_2025_agix_holdings",
            skiprows=1
        )
        
        # 加载每日监控的holdings表数据 - A1:O47范围
        try:
            daily_holdings = pd.read_excel(
                "ciq reference/data/每日数据监控.xlsx",
                sheet_name="holdings",
                header=0,  # 第1行作为列名
                skiprows=[1, 2],  # 跳过第2-3行无用信息
                usecols="A:O",  # 读取A到O列
                nrows=47  # 读取47行数据
            )
        except Exception as e:
            st.warning(f"无法加载每日监控holdings表数据: {e}")
            daily_holdings = None
        
        return filled_pri, shares, daily_monitor, holdings, raw1_data, daily_holdings
    except Exception as e:
        st.error(f"数据加载错误: {e}")
        return None, None, None, None, None, None

def calculate_returns(prices, periods=[1, 5, 30, 90, 252]):
    """计算不同期间的收益率"""
    returns = {}
    for period in periods:
        if len(prices) > period:
            returns[f'{period}d'] = (prices.iloc[0] - prices.iloc[period]) / prices.iloc[period]
    return returns

def calculate_risk_metrics(prices, returns):
    """计算风险指标"""
    daily_returns = prices.pct_change().dropna()
    
    # 年化波动率
    volatility = daily_returns.std() * np.sqrt(252)
    
    # 最大回撤
    cumulative_returns = (1 + daily_returns).cumprod()
    rolling_max = cumulative_returns.expanding().max()
    drawdown = (cumulative_returns - rolling_max) / rolling_max
    max_drawdown = drawdown.min()
    
    # VaR (95% 置信水平)
    var_95 = daily_returns.quantile(0.05)
    
    # Sharpe比率 (假设无风险利率为2%)
    excess_returns = daily_returns - 0.02/252
    sharpe_ratio = excess_returns.mean() / daily_returns.std() * np.sqrt(252)
    
    return {
        'volatility': volatility,
        'max_drawdown': max_drawdown,
        'var_95': var_95,
        'sharpe_ratio': sharpe_ratio
    }

def calculate_contribution(prices, shares):
    """计算持仓贡献度分析"""
    try:
        # 计算价格变化
        price_diff = prices.iloc[0] - prices.iloc[1]
        price_delta = price_diff / prices.iloc[1]
        
        # 计算价格影响 - 使用正确的列索引
        if len(shares.columns) > 1:
            shares_col = shares.iloc[:, 1]  # 第2列
        else:
            shares_col = shares.iloc[:, 0]  # 如果只有1列，使用第1列
        
        price_impact = price_diff * shares_col
        
        # 计算总市值
        yesterday_value = prices.iloc[1] * shares_col
        total_value = yesterday_value.sum()
        
        # 计算贡献度
        contribution = price_impact / total_value
        
        return price_delta, price_impact, contribution
    except Exception as e:
        st.error(f"贡献度计算错误: {e}")
        return pd.Series(), pd.Series(), pd.Series()

def calculate_cumulative_returns(prices, periods=[30, 90, 180, 252]):
    """计算不同期间的累计收益率"""
    try:
        cumulative_returns = {}
        
        for period in periods:
            if len(prices) > period:
                # 计算从period天前到现在的累计收益率
                start_prices = prices.iloc[period]
                end_prices = prices.iloc[0]
                returns = (end_prices - start_prices) / start_prices * 100
                cumulative_returns[f'{period}d'] = returns
        
        return cumulative_returns
    except Exception as e:
        st.error(f"累计收益率计算错误: {e}")
        return {}

def calculate_specific_period_returns(prices):
    """计算特定期间的累计收益率（since 2024, since 2025）"""
    try:
        specific_returns = {}
        
        # 确保索引是日期类型并按时间排序
        prices = prices.sort_index()
        
        # 计算since 2024的累计收益率
        start_date_2024 = pd.Timestamp('2024-01-01')
        available_dates = prices.index.sort_values()
        
        # 找到2024年或之后的第一天
        start_idx_2024 = None
        for date in available_dates:
            if date >= start_date_2024:
                start_idx_2024 = date
                break
        
        if start_idx_2024 is not None:
            prices_since_2024 = prices.loc[start_idx_2024:].copy()
            if len(prices_since_2024) > 1:
                base_prices_2024 = prices_since_2024.iloc[0]
                current_prices = prices_since_2024.iloc[-1]
                
                for ticker in prices.columns:
                    if ticker in base_prices_2024.index and ticker in current_prices.index:
                        base_price = base_prices_2024[ticker]
                        current_price = current_prices[ticker]
                        if pd.notna(base_price) and base_price != 0:
                            returns_2024 = (current_price - base_price) / base_price * 100
                            specific_returns[f'{ticker}_since2024'] = returns_2024
        
        # 计算since 2025的累计收益率
        start_date_2025 = pd.Timestamp('2025-01-01')
        start_idx_2025 = None
        for date in available_dates:
            if date >= start_date_2025:
                start_idx_2025 = date
                break
        
        if start_idx_2025 is not None:
            prices_since_2025 = prices.loc[start_idx_2025:].copy()
            if len(prices_since_2025) > 1:
                base_prices_2025 = prices_since_2025.iloc[0]
                current_prices = prices_since_2025.iloc[-1]
                
                for ticker in prices.columns:
                    if ticker in base_prices_2025.index and ticker in current_prices.index:
                        base_price = base_prices_2025[ticker]
                        current_price = current_prices[ticker]
                        if pd.notna(base_price) and base_price != 0:
                            returns_2025 = (current_price - base_price) / base_price * 100
                            specific_returns[f'{ticker}_since2025'] = returns_2025
        
        return specific_returns
    except Exception as e:
        st.error(f"特定期间收益率计算错误: {e}")
        return {}

def plot_cumulative_returns(prices, selected_tickers=None, periods=[30, 90]):
    """绘制累计收益率图表"""
    try:
        if prices.empty or len(prices) < 30:
            st.warning("价格数据不足，无法计算累计收益率")
            return go.Figure()
        
        # 计算不同期间的累计收益率
        cumulative_returns = calculate_cumulative_returns(prices, periods)
        
        if not cumulative_returns:
            return go.Figure()
        
        # 如果没有选择Ticker，使用所有可用的Ticker
        if selected_tickers is None or len(selected_tickers) == 0:
            available_tickers = list(prices.columns)
            selected_tickers = available_tickers[:10]  # 默认显示前10个
        
        # 过滤出用户选择的Ticker
        available_tickers = [ticker for ticker in selected_tickers if ticker in prices.columns]
        
        if not available_tickers:
            st.warning("所选Ticker在价格数据中不存在")
            return go.Figure()
        
        # 创建图表
        fig = go.Figure()
        
        # 为每个期间添加一个条形图
        colors = ['#1f77b4', '#ff7f0e']
        
        for i, (period, returns) in enumerate(cumulative_returns.items()):
            if period in [f'{p}d' for p in periods]:
                # 只显示用户选择的Ticker的数据
                period_returns = returns.loc[available_tickers]
                
                fig.add_trace(go.Bar(
                    name=f'{period}累计收益率',
                    x=period_returns.index,
                    y=period_returns.values,
                    marker_color=colors[i % len(colors)],
                    text=[f'{x:.1f}%' for x in period_returns.values],
                    textposition='auto',
                    opacity=0.8
                ))
        
        fig.update_layout(
            title=f"选定持仓累计收益率对比",
            xaxis_title="股票代码",
            yaxis_title="累计收益率 (%)",
            barmode='group',  # 分组显示
            height=600,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        
        # 添加零线
        fig.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1)
        
        return fig
    except Exception as e:
        st.error(f"累计收益率图表生成错误: {e}")
        return go.Figure()



def plot_cumulative_returns_since_2025(prices, selected_tickers=None):
    """绘制自2025年初的累计收益率对比图（平滑曲线图）"""
    try:
        if prices.empty:
            st.warning("价格数据为空")
            return go.Figure()
        
        # 确保索引是日期类型并按时间排序
        prices = prices.sort_index()
        
        # 找到2025年1月1日或之后的第一天
        start_date = pd.Timestamp('2025-01-01')
        available_dates = prices.index.sort_values()
        
        # 找到2025年或之后的第一天
        start_idx = None
        for date in available_dates:
            if date >= start_date:
                start_idx = date
                break
        
        if start_idx is None:
            st.warning("没有找到2025年的数据，使用最新数据")
            # 如果没有2025年数据，使用最近30天的数据
            start_idx = available_dates[-30] if len(available_dates) >= 30 else available_dates[0]
        
        # 获取从起始日期开始的数据
        prices_since_start = prices.loc[start_idx:].copy()
        
        if len(prices_since_start) < 2:
            st.warning("数据不足")
            return go.Figure()
        
        # 如果没有选择Ticker，使用所有可用的Ticker
        if selected_tickers is None or len(selected_tickers) == 0:
            available_tickers = list(prices.columns)
            selected_tickers = available_tickers[:10]  # 默认显示前10个
        
        # 过滤出用户选择的Ticker
        available_tickers = [ticker for ticker in selected_tickers if ticker in prices.columns]
        
        if not available_tickers:
            st.warning("所选Ticker在价格数据中不存在")
            return go.Figure()
        
        # 计算累计收益率（以起始日期为基准）
        base_prices = prices_since_start.iloc[0]
        cumulative_returns = {}
        
        for ticker in available_tickers:
            if ticker in prices_since_start.columns:
                price_series = prices_since_start[ticker]
                base_price = base_prices[ticker]
                if pd.notna(base_price) and base_price != 0:
                    # 计算每日累计收益率
                    returns = (price_series - base_price) / base_price * 100
                    cumulative_returns[ticker] = returns
        
        if not cumulative_returns:
            st.warning("无法计算累计收益率")
            return go.Figure()
        
        # 创建图表
        fig = go.Figure()
        
        # 颜色列表
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', 
                 '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
        
        for i, (ticker, returns) in enumerate(cumulative_returns.items()):
            # 移除NaN值，但保留有效的数据点
            valid_data = returns.dropna()
            if len(valid_data) > 1:  # 至少需要2个点才能画线
                fig.add_trace(go.Scatter(
                    x=valid_data.index,
                    y=valid_data.values,
                    mode='lines',
                    name=ticker,
                    line=dict(color=colors[i % len(colors)], width=2),
                    hovertemplate=f'{ticker}<br>日期: %{{x}}<br>累计收益率: %{{y:.2f}}%<extra></extra>'
                ))
        
        # 设置图表标题
        if start_idx >= pd.Timestamp('2025-01-01'):
            title = "自2025年初累计收益率对比"
        else:
            title = f"自{start_idx.strftime('%Y-%m-%d')}累计收益率对比"
        
        fig.update_layout(
            title=title,
            xaxis_title="日期",
            yaxis_title="累计收益率 (%)",
            height=500,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            ),
            hovermode='x unified'
        )
        
        # 添加零线
        fig.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1)
        
        return fig
    except Exception as e:
        st.error(f"累计收益率对比图生成错误: {e}")
        return go.Figure()


def plot_contribution_analysis(contribution, top_n=10):
    """绘制贡献度分析图表"""
    try:
        if contribution.empty:
            st.warning("贡献度数据为空")
            return go.Figure()
        
        # 获取贡献度绝对值最大的前N个持仓，但保留原始符号
        top_contributors_abs = contribution.abs().nlargest(top_n)
        top_contributors = contribution.loc[top_contributors_abs.index]
        
        fig = go.Figure()
        
        colors = ['red' if x < 0 else 'green' for x in top_contributors]
        
        fig.add_trace(go.Bar(
            x=top_contributors.index,
            y=top_contributors.values * 100,
            marker_color=colors,
            text=[f'{x:.2f}%' for x in top_contributors.values * 100],
            textposition='auto'
        ))
        
        fig.update_layout(
            title=f"Top {top_n} Holdings Contribution to Fund Performance",
            xaxis_title="Holdings",
            yaxis_title="Contribution (%)",
            height=500
        )
        
        return fig
    except Exception as e:
        st.error(f"贡献度图表生成错误: {e}")
        return go.Figure()

def plot_sector_allocation(holdings):
    """绘制行业配置图表"""
    try:
        if holdings is None or 'Sector' not in holdings.columns:
            return None
        
        sector_counts = holdings['Sector'].value_counts()
        
        fig = px.pie(
            values=sector_counts.values,
            names=sector_counts.index,
            title="Sector Allocation by Company Count"
        )
        
        fig.update_layout(height=400)
        return fig
    except Exception as e:
        st.error(f"行业配置图表生成错误: {e}")
        return None

def plot_sector_market_cap(holdings):
    """绘制按市值的行业配置图表"""
    try:
        if holdings is None or 'Sector' not in holdings.columns:
            return None
        
        # 查找市值列（可能的列名）
        market_cap_col = None
        possible_market_cap_cols = ['Market Cap', 'MarketCap', 'Market_Cap', 'MarketCap(MLN)', 'Market Cap(MLN)', '市值', 'Market Value']
        
        for col in holdings.columns:
            if any(mc in str(col) for mc in possible_market_cap_cols):
                market_cap_col = col
                break
        
        if market_cap_col is None:
            st.warning("未找到市值列，无法生成按市值的行业配置图")
            return None
        
        # 按行业汇总市值
        sector_market_cap = holdings.groupby('Sector')[market_cap_col].sum()
        
        fig = px.pie(
            values=sector_market_cap.values,
            names=sector_market_cap.index,
            title=f"Sector Allocation by Market Cap ({market_cap_col})"
        )
        
        fig.update_layout(height=400)
        return fig
    except Exception as e:
        st.error(f"按市值的行业配置图表生成错误: {e}")
        return None

def calculate_sector_contribution(holdings):
    """计算行业贡献表格"""
    try:
        if holdings is None or 'Sector' not in holdings.columns:
            return None
        
        # 查找贡献度列（可能的列名）
        contribution_col = None
        possible_contribution_cols = ['Contribute']
        
        for col in holdings.columns:
            if any(contrib in str(col) for contrib in possible_contribution_cols):
                contribution_col = col
                break
        
        if contribution_col is None:
            st.warning("未找到贡献度列，无法生成行业贡献表格")
            return None
        
        # 按行业汇总贡献度
        sector_contribution = holdings.groupby('Sector')[contribution_col].sum()
        
        # 创建贡献表格
        contribution_df = pd.DataFrame({
            '行业': sector_contribution.index,
            '贡献度总和': sector_contribution.values
        })
        
        # 按贡献度排序
        contribution_df = contribution_df.sort_values('贡献度总和', ascending=False)
        
        return contribution_df
    except Exception as e:
        st.error(f"行业贡献计算错误: {e}")
        return None

def main():
    # 主标题
    st.markdown('<h1 class="main-header">📊 AGIX Fund Monitor</h1>', unsafe_allow_html=True)
    
    # 加载数据
    with st.spinner("正在加载数据..."):
        filled_pri, shares, daily_monitor, holdings, raw1_data, daily_holdings = load_data()
    
    if filled_pri is None:
        st.error("无法加载数据，请检查文件路径")
        return
    
    # 侧边栏
    st.sidebar.header("📈 监控设置")
    
    # 占位符 - 未来可添加监控设置功能
    st.sidebar.info("监控设置功能正在开发中...")
    st.sidebar.write("未来将支持：")
    st.sidebar.write("• 自定义时间范围")
    st.sidebar.write("• 风险预警设置")
    st.sidebar.write("• 收益率目标设置")
    st.sidebar.write("• 自动报告生成")
    

    
    # 主页面标签
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 概览", 
        "📊 基金对比",
        "🎯 贡献度分析", 
        "📈 持仓收益率分析", 
        "⚠️ 风险分析", 
        "📋 持仓详情"
    ])
    
    # 概览标签页
    with tab1:
        st.header("📊 基金概览")
        
        col1, col2, col3, col4 = st.columns(4)
        
        # 计算关键指标
        if len(filled_pri) > 1:
            try:
                current_value = filled_pri.iloc[0].sum()
                prev_value = filled_pri.iloc[1].sum()
                
                # 计算日收益率
                if pd.notna(prev_value) and prev_value != 0:
                    daily_return = (current_value - prev_value) / prev_value
                    daily_return_str = f"{daily_return:.2%}"
                else:
                    daily_return_str = "N/A"
                
                with col1:
                    if pd.notna(current_value):
                        st.metric(
                            "当前净值",
                            f"${current_value:,.2f}",
                            daily_return_str,
                            delta_color="normal"
                        )
                    else:
                        st.metric("当前净值", "N/A", "数据无效")
                
                with col2:
                    st.metric(
                        "持仓数量",
                        len(shares),
                        "",
                        delta_color="normal"
                    )
                
                with col3:
                    # 计算年化收益率
                    try:
                        if len(filled_pri) > 252:
                            # 如果有超过一年的数据，使用一年前的数据
                            past_value = filled_pri.iloc[252].sum()
                            if pd.notna(past_value) and past_value != 0:
                                annual_return = (current_value / past_value - 1)
                                st.metric(
                                    "年化收益率",
                                    f"{annual_return:.2%}",
                                    "",
                                    delta_color="normal"
                                )
                            else:
                                st.metric("年化收益率", "N/A", "数据无效")
                        elif len(filled_pri) > 30:
                            # 如果数据不足一年但超过30天，使用年化计算
                            days_available = len(filled_pri) - 1
                            past_value = filled_pri.iloc[-1].sum()
                            if pd.notna(past_value) and past_value != 0:
                                total_return = (current_value / past_value - 1)
                                annual_return = (1 + total_return) ** (252 / days_available) - 1
                                st.metric(
                                    "年化收益率",
                                    f"{annual_return:.2%}",
                                    f"(基于{days_available}天数据)",
                                    delta_color="normal"
                                )
                            else:
                                st.metric("年化收益率", "N/A", "数据无效")
                        else:
                            st.metric("年化收益率", "N/A", "数据不足")
                    except Exception as e:
                        st.metric("年化收益率", "N/A", f"计算错误")
                
                with col4:
                    # 计算波动率
                    try:
                        daily_returns = filled_pri.pct_change().dropna()
                        if not daily_returns.empty:
                            volatility = daily_returns.std().mean() * np.sqrt(252)
                            if pd.notna(volatility):
                                st.metric(
                                    "年化波动率",
                                    f"{volatility:.2%}",
                                    "",
                                    delta_color="normal"
                                )
                            else:
                                st.metric("年化波动率", "N/A", "数据无效")
                        else:
                            st.metric("年化波动率", "N/A", "数据不足")
                    except Exception as e:
                        st.metric("年化波动率", "N/A", "计算错误")
            except Exception as e:
                st.error(f"指标计算错误: {e}")
        
        # 基金概览信息
        st.subheader("📊 基金基本信息")
        if daily_monitor is not None and len(daily_monitor) > 0:
            try:
                # 显示最新的基金资产信息
                latest_data = daily_monitor.iloc[0] if len(daily_monitor) > 0 else None
                if latest_data is not None and 'Fund Asset(MLN USD)' in latest_data:
                    fund_asset = latest_data['Fund Asset(MLN USD)']
                    st.metric("基金资产规模", f"${fund_asset:,.2f}M")
            except Exception as e:
                st.error(f"显示基金信息时出错: {e}")
        else:
            st.warning("无法加载基金基本信息")
        
        # AGIX表现总结
        if raw1_data is not None:
            st.subheader("🎯 AGIX表现总结")
            agix_data = raw1_data[raw1_data['Name'] == 'ETNA']
            if not agix_data.empty:
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    if 'DTD' in agix_data.columns:
                        dtd_val = pd.to_numeric(agix_data['DTD'].iloc[0], errors='coerce') * 100
                        st.metric("日收益率", f"{dtd_val:.2f}%" if pd.notna(dtd_val) else "N/A")
                
                with col2:
                    if 'WTD' in agix_data.columns:
                        wtd_val = pd.to_numeric(agix_data['WTD'].iloc[0], errors='coerce') * 100
                        st.metric("周收益率", f"{wtd_val:.2f}%" if pd.notna(wtd_val) else "N/A")
                
                with col3:
                    if 'YTD' in agix_data.columns:
                        ytd_val = pd.to_numeric(agix_data['YTD'].iloc[0], errors='coerce') * 100
                        st.metric("年收益率", f"{ytd_val:.2f}%" if pd.notna(ytd_val) else "N/A")
                
                with col4:
                    if 'Return since 2024' in agix_data.columns:
                        ret_2024_val = pd.to_numeric(agix_data['Return since 2024'].iloc[0], errors='coerce') * 100
                        st.metric("2024年收益", f"{ret_2024_val:.2f}%" if pd.notna(ret_2024_val) else "N/A")
        
 
    
    # 基金对比标签页
    with tab2:
        st.header("📊 基金对比")
        
        # 第一部分：AGIX与主要基金对比表格
        if raw1_data is not None:
            st.subheader("📋 AGIX与主要基金对比数据")
            st.write("**基金收益率对比表格:**")
            
            # 创建格式化后的数据框用于显示
            display_raw1 = raw1_data.copy()
            
            # 定义需要以百分比形式显示的列
            percentage_columns = ['DTD', 'WTD', 'YTD', 'Return since 2024', 'Return since launch']
            
            # 应用百分比格式化
            for col in percentage_columns:
                if col in display_raw1.columns:
                    # 确保数据是数值类型
                    display_raw1[col] = pd.to_numeric(display_raw1[col], errors='coerce')
                    # 转换为百分比格式
                    display_raw1[col] = display_raw1[col].apply(lambda x: f"{x:.2f}%" if pd.notna(x) else "")
            
            # 确保所有列都是字符串类型，避免Arrow序列化问题
            for col in display_raw1.columns:
                display_raw1[col] = display_raw1[col].astype(str)
            
            st.dataframe(display_raw1, use_container_width=True)
            
            # 第二部分：收益率对比图表
            st.subheader("📈 收益率对比图表")
            
            # 选择要展示的指数/基金
            # 创建基金名称到ticker的映射
            fund_mapping = dict(zip(raw1_data['Ticker.1'], raw1_data['Name']))
            available_tickers = raw1_data['Ticker.1'].tolist()
            
            selected_tickers = st.multiselect(
                "选择要对比的指数/基金:",
                available_tickers,
                default=available_tickers  # 默认选择所有基金
            )
            
            # 将选中的ticker转换为对应的基金名称用于数据过滤
            selected_funds = [fund_mapping[ticker] for ticker in selected_tickers if ticker in fund_mapping]
            
            if selected_funds:
                # 创建分组柱状图
                fig = go.Figure()
                
                # 定义指标颜色方案 - 每个指标一个固定颜色
                metric_colors = {
                    'DTD': '#1f77b4',      # 蓝色
                    'WTD': '#ff7f0e',      # 橙色
                    'YTD': '#2ca02c',      # 绿色
                    'Return since 2024': '#d62728'  # 红色
                }
                
                # 过滤选中的基金数据
                filtered_data = raw1_data[raw1_data['Name'].isin(selected_funds)]
                
                # 为每个指标创建柱状图
                for i, metric in enumerate(['DTD', 'WTD', 'YTD', 'Return since 2024']):
                    if metric in filtered_data.columns:
                        # 获取数值数据（不包含百分比符号）
                        values = pd.to_numeric(filtered_data[metric], errors='coerce') * 100  # 转换为百分比
                        
                        # 为每个基金设置颜色，AGIX突出显示
                        fund_colors = []
                        for name in filtered_data['Name']:
                            if name == 'ETNA':
                                fund_colors.append(metric_colors[metric])  # AGIX使用指标颜色
                            else:
                                fund_colors.append(metric_colors[metric])  # 其他基金也使用相同颜色
                        
                        fig.add_trace(go.Bar(
                            name=metric,
                            x=filtered_data['Ticker.1'],  # 使用Ticker.1列作为X轴标签
                            y=values,
                            marker_color=fund_colors,
                            text=[f'{val:.2f}%' if pd.notna(val) else '' for val in values],
                            textposition='auto',
                            textfont=dict(size=10),
                            offsetgroup=i,
                            width=0.15,
                            opacity=0.8
                        ))
                
                # 更新布局
                fig.update_layout(
                    title=dict(
                        text="AGIX与主要基金收益率对比",
                        x=0.5,
                        font=dict(size=16, color='white')
                    ),
                    xaxis_title="基金/指数",
                    yaxis_title="收益率 (%)",
                    barmode='group',
                    height=600,
                    showlegend=True,
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1,
                        bgcolor='rgba(0,0,0,0)',
                        bordercolor='rgba(0,0,0,0)'
                    ),
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(color='white'),
                    margin=dict(t=80, b=80, l=80, r=80)
                )
                
                # 更新x轴样式
                fig.update_xaxes(
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='rgba(128,128,128,0.2)',
                    tickangle=45
                )
                
                # 更新y轴样式
                fig.update_yaxes(
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='rgba(128,128,128,0.2)',
                    zeroline=True,
                    zerolinecolor='rgba(128,128,128,0.5)',
                    zerolinewidth=1
                )
                
                # 添加水平零线
                fig.add_hline(y=0, line_dash="dash", line_color="rgba(128,128,128,0.8)", line_width=1)
                
                st.plotly_chart(fig, use_container_width=True)
         
        # 第三部分：详细基金对比数据
        if daily_monitor is not None:
            st.subheader("📊 详细基金对比数据")
            st.write("**基金与基准指数详细对比:**")
            
            try:
                # 确保数据类型正确（排除非数值列如Capital IQ Ticker）
                display_monitor = daily_monitor.copy()
                numeric_columns = ['Fund Asset(MLN USD)', 'Volume(MLN)', 'Price Change', 'Daily Flow', 
                                  'YTD Flow', 'Flow since Jan 2023', 'Expense Ratio', 'Return since 2024', 
                                  'Return since 2025', '30D Vol', 'Holdings']
                
                for col in display_monitor.columns:
                    if col in numeric_columns and display_monitor[col].dtype == 'object':
                        display_monitor[col] = pd.to_numeric(display_monitor[col], errors='coerce')
                
                # 确保所有列都是数值或字符串类型，避免Arrow序列化问题
                for col in display_monitor.columns:
                    if display_monitor[col].dtype == 'object':
                        display_monitor[col] = display_monitor[col].astype(str)
                
                st.dataframe(display_monitor, use_container_width=True)
                
            except Exception as e:
                st.error(f"显示详细基金对比数据时出错: {e}")
                st.write("无法显示详细基金对比数据")
        else:
            st.warning("无法加载详细基金对比数据")
    
    # 持仓收益率分析标签页
    with tab4:
        st.header("📈 持仓收益率分析")
        
        try:
            # 计算不同期间的收益率
            returns = calculate_returns(filled_pri)
            
            # 准备收益率数据用于热力图
            returns_df = pd.DataFrame(returns).T
            returns_df = returns_df * 100  # 转换为百分比
            
            # 确保数据类型兼容
            returns_df = returns_df.astype(float)
            
            # 收益率热力图
            st.subheader("收益率热力图")
            fig = px.imshow(
                returns_df,
                aspect="auto",
                title="Returns Heatmap by Period"
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # 累计收益率分析
            st.subheader("📈 累计收益率分析")
            
            # 获取所有可用的Ticker
            available_tickers = list(filled_pri.columns)
            
            # 计算30天收益率来获取Top 10 Holdings
            cumulative_returns_30d = calculate_cumulative_returns(filled_pri, [30])
            if '30d' in cumulative_returns_30d:
                top_10_holdings = cumulative_returns_30d['30d'].nlargest(10).index.tolist()
            else:
                top_10_holdings = available_tickers[:10]
            
            # 创建多选器
            selected_tickers = st.multiselect(
                "选择要显示的股票代码:",
                options=available_tickers,
                default=top_10_holdings,
                help="默认显示30天收益率最高的前10个持仓"
            )
            
            # 绘制累计收益率图表
            if selected_tickers:
                cum_returns_fig = plot_cumulative_returns(filled_pri, selected_tickers, [30, 90])
                st.plotly_chart(cum_returns_fig, use_container_width=True)
            else:
                st.warning("请至少选择一个股票代码")
            
            # 累计收益率详细数据表格
            st.subheader("📊 累计收益率详细数据")
            
            # 计算累计收益率数据 - 添加更多期间
            cumulative_returns_data = calculate_cumulative_returns(filled_pri, [1, 5, 30, 90])
            
            if cumulative_returns_data:
                # 创建累计收益率表格
                returns_df_cum = pd.DataFrame(cumulative_returns_data)
                
                # 计算特定期间的收益率
                specific_returns = calculate_specific_period_returns(filled_pri)
                
                # 添加since2024和since2025数据
                if specific_returns:
                    # 将特定期间数据转换为DataFrame格式
                    since2024_data = {}
                    since2025_data = {}
                    
                    for key, value in specific_returns.items():
                        if key.endswith('_since2024'):
                            ticker = key.replace('_since2024', '')
                            since2024_data[ticker] = value
                        elif key.endswith('_since2025'):
                            ticker = key.replace('_since2025', '')
                            since2025_data[ticker] = value
                    
                    # 添加到主表格
                    if since2024_data:
                        returns_df_cum['since2024'] = pd.Series(since2024_data)
                    if since2025_data:
                        returns_df_cum['since2025'] = pd.Series(since2025_data)
                
                # 添加持仓权重信息
                if len(shares.columns) > 1:
                    shares_col = shares.iloc[:, 1]  # 第2列
                else:
                    shares_col = shares.iloc[:, 0]  # 如果只有1列，使用第1列
                
                yesterday_value = filled_pri.iloc[1] * shares_col
                total_value = yesterday_value.sum()
                weight = yesterday_value / total_value * 100
                
                returns_df_cum['Weight(%)'] = weight
                
                # 重新排序列 - 按时间顺序和重要性排序
                column_order = ['1d', '5d', '30d', '90d', 'since2024', 'since2025', 'Weight(%)']
                available_columns = [col for col in column_order if col in returns_df_cum.columns]
                returns_df_cum = returns_df_cum[available_columns]
                
                # 格式化显示 - 只对收益率列应用颜色，权重列保持白色
                def color_returns_only(df):
                    """只对收益率列应用颜色，权重列保持白色"""
                    styled_df = df.copy()
                    for col in df.columns:
                        if col in ['1d', '5d', '30d', '90d', 'since2024', 'since2025']:
                            styled_df[col] = df[col].apply(lambda x: 'color: red' if x < 0 else 'color: green' if x > 0 else '')
                        else:
                            styled_df[col] = ''  # 权重列保持白色
                    return styled_df
                
                # 准备格式化字典
                format_dict = {}
                for col in returns_df_cum.columns:
                    if col == 'Weight(%)':
                        format_dict[col] = '{:.2f}%'
                    else:
                        format_dict[col] = '{:.2f}%'
                
                # 显示表格
                st.dataframe(returns_df_cum.style.format(format_dict).apply(color_returns_only, axis=None), use_container_width=True)
            
            # 自2025年初累计收益率对比图
            st.subheader("📈 自2025年初累计收益率对比")
            
            # 获取所有可用的Ticker
            available_tickers_2025 = list(filled_pri.columns)
            
            # 计算30天收益率来获取Top 10 Holdings（用于默认选择）
            cumulative_returns_30d_for_2025 = calculate_cumulative_returns(filled_pri, [30])
            if '30d' in cumulative_returns_30d_for_2025:
                top_10_holdings_2025 = cumulative_returns_30d_for_2025['30d'].nlargest(10).index.tolist()
            else:
                top_10_holdings_2025 = available_tickers_2025[:10]
            
            # 创建独立的多选器
            selected_tickers_2025 = st.multiselect(
                "选择要显示的股票代码 (累计收益率对比):",
                options=available_tickers_2025,
                default=top_10_holdings_2025,
                help="默认显示30天收益率最高的前10个持仓"
            )
            
            # 绘制累计收益率对比图
            if selected_tickers_2025:
                cum_returns_2025_fig = plot_cumulative_returns_since_2025(filled_pri, selected_tickers_2025)
                st.plotly_chart(cum_returns_2025_fig, use_container_width=True)
            else:
                st.warning("请至少选择一个股票代码")
                
        except Exception as e:
            st.error(f"持仓收益率分析错误: {e}")
    
    # 风险分析标签页
    with tab5:
        st.header("⚠️ 风险分析")
        
        try:
            # 计算风险指标
            risk_metrics = calculate_risk_metrics(filled_pri, None)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("风险指标")
                
                # 波动率
                volatility_avg = risk_metrics['volatility'].mean()
                st.metric(
                    "年化波动率",
                    f"{volatility_avg:.2%}",
                    "",
                    delta_color="normal"
                )
                
                # 最大回撤
                max_dd = risk_metrics['max_drawdown'].min()
                st.metric(
                    "最大回撤",
                    f"{max_dd:.2%}",
                    "",
                    delta_color="inverse"
                )
            
            with col2:
                # VaR
                var_avg = risk_metrics['var_95'].mean()
                st.metric(
                    "VaR (95%)",
                    f"{var_avg:.2%}",
                    "",
                    delta_color="inverse"
                )
                
                # Sharpe比率
                sharpe_avg = risk_metrics['sharpe_ratio'].mean()
                st.metric(
                    "Sharpe比率",
                    f"{sharpe_avg:.2f}",
                    "",
                    delta_color="normal"
                )
            
            # 回撤图表
            st.subheader("回撤分析")
            daily_returns = filled_pri.pct_change().dropna()
            cumulative_returns = (1 + daily_returns).cumprod()
            rolling_max = cumulative_returns.expanding().max()
            drawdown = (cumulative_returns - rolling_max) / rolling_max
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=drawdown.index,
                y=drawdown.iloc[:, 0] * 100,
                mode='lines',
                fill='tonexty',
                name='Drawdown',
                line=dict(color='red')
            ))
            
            fig.update_layout(
                title="Fund Drawdown",
                xaxis_title="Date",
                yaxis_title="Drawdown (%)",
                height=400
            )
            
            st.plotly_chart(fig, use_container_width=True)
        except Exception as e:
            st.error(f"风险分析错误: {e}")
    
    # 贡献度分析标签页
    with tab3:
        st.header("🎯 贡献度分析")
        
        try:
            # 计算贡献度
            price_delta, price_impact, contribution = calculate_contribution(filled_pri, shares)
            
            if not price_delta.empty:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("价格变化分析")
                    
                    # 显示涨幅最大的股票（只显示正涨幅）
                    positive_changes = price_delta[price_delta > 0]
                    if len(positive_changes) > 0:
                        top_gainers = positive_changes.nlargest(5) * 100
                        st.write("**📈 涨幅最大的股票:**")
                        for ticker, change in top_gainers.items():
                            st.markdown(f'<span class="positive">{ticker}: {change:.2f}%</span>', unsafe_allow_html=True)
                    else:
                        st.write("**📈 涨幅最大的股票:**")
                        st.write("今日无上涨股票")
                
                with col2:
                    st.subheader("跌幅最大的股票")
                    
                    # 显示跌幅最大的股票（只显示负跌幅）
                    negative_changes = price_delta[price_delta < 0]
                    if len(negative_changes) > 0:
                        top_losers = negative_changes.nsmallest(5) * 100
                        st.write("**📉 跌幅最大的股票:**")
                        for ticker, change in top_losers.items():
                            st.markdown(f'<span class="negative">{ticker}: {change:.2f}%</span>', unsafe_allow_html=True)
                    else:
                        st.write("**📉 跌幅最大的股票:**")
                        st.write("今日无下跌股票")
                
                # 贡献度图表
                st.subheader("持仓贡献度分析")
                fig = plot_contribution_analysis(contribution)
                st.plotly_chart(fig, use_container_width=True)
                
                # 贡献度表格
                st.subheader("昨日回报详细贡献度数据")
                
                # 计算持仓权重
                if len(shares.columns) > 1:
                    shares_col = shares.iloc[:, 1]  # 第2列
                else:
                    shares_col = shares.iloc[:, 0]  # 如果只有1列，使用第1列
                
                yesterday_value = filled_pri.iloc[1] * shares_col
                total_value = yesterday_value.sum()
                weight = yesterday_value / total_value * 100
                
                contribution_df = pd.DataFrame({
                    'Price_Change(%)': price_delta * 100,
                    'Weight(%)': weight,
                    'Price_Impact': price_impact,
                    'Contribution(%)': contribution * 100
                }).sort_values('Contribution(%)', ascending=False)
                
                # 确保数据类型兼容
                for col in contribution_df.columns:
                    contribution_df[col] = pd.to_numeric(contribution_df[col], errors='coerce')
                
                # 添加颜色样式 - 只对价格变化和贡献度列应用颜色
                def color_negative_red(val):
                    if isinstance(val, (int, float)):
                        if val < 0:
                            return 'color: red'
                        elif val > 0:
                            return 'color: green'
                    return ''
                
                def color_contribution_only(df):
                    """只对价格变化和贡献度列应用颜色，权重列保持白色"""
                    styled_df = df.copy()
                    for col in df.columns:
                        if col in ['Price_Change(%)', 'Contribution(%)']:
                            styled_df[col] = df[col].apply(lambda x: 'color: red' if x < 0 else 'color: green' if x > 0 else '')
                        else:
                            styled_df[col] = ''  # 权重列保持白色
                    return styled_df
                
                # 添加搜索和过滤功能
                col1, col2 = st.columns(2)
                
                with col1:
                    # 搜索特定股票
                    search_term = st.text_input("🔍 搜索股票代码或名称:", "")
                
                with col2:
                    # 过滤选项
                    filter_option = st.selectbox(
                        "📊 过滤选项:",
                        ["全部持仓", "正贡献", "负贡献", "权重前10", "贡献度前10"]
                    )
                
                # 应用过滤
                filtered_df = contribution_df.copy()
                
                if search_term:
                    filtered_df = filtered_df[filtered_df.index.str.contains(search_term, case=False, na=False)]
                
                if filter_option == "正贡献":
                    filtered_df = filtered_df[filtered_df['Contribution(%)'] > 0]
                elif filter_option == "负贡献":
                    filtered_df = filtered_df[filtered_df['Contribution(%)'] < 0]
                elif filter_option == "权重前10":
                    filtered_df = filtered_df.nlargest(10, 'Weight(%)')
                elif filter_option == "贡献度前10":
                    filtered_df = filtered_df.nlargest(10, 'Contribution(%)')
                
                # 显示过滤后的数据
                st.dataframe(filtered_df.style.format({
                    'Price_Change(%)': '{:.2f}%',
                    'Weight(%)': '{:.2f}%',
                    'Price_Impact': '{:.2f}',
                    'Contribution(%)': '{:.2f}%'
                }).apply(color_contribution_only, axis=None), use_container_width=True)
                
                # 添加统计信息
                st.subheader("📊 贡献度统计")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    positive_contrib = len(contribution_df[contribution_df['Contribution(%)'] > 0])
                    st.metric("正贡献持仓数", positive_contrib)
                
                with col2:
                    negative_contrib = len(contribution_df[contribution_df['Contribution(%)'] < 0])
                    st.metric("负贡献持仓数", negative_contrib)
                
                with col3:
                    total_contrib = len(contribution_df)
                    st.metric("总持仓数", total_contrib)
                
                with col4:
                    avg_contrib = contribution_df['Contribution(%)'].mean()
                    st.metric("平均贡献度", f"{avg_contrib:.2f}%")
                
                # 添加详细分析报告
                st.subheader("📊 详细分析报告")
                
                # 计算总影响
                total_impact = sum(price_impact)
                
                # 价格下降分析
                down_delta = price_delta.sort_values() * 100
                down_stocks = down_delta[down_delta <= -2]
                
                if len(down_stocks) > 0:
                    st.write("**📉 价格下降幅度较大的股票 (>2%):**")
                    down_text = ""
                    for ticker, change in down_stocks.items():
                        # 股票名称映射
                        display_name = ticker
                        if ticker == "KOSE:A000660":
                            display_name = "SK Hynix"
                        elif ticker == "TASE:NICE":
                            display_name = "NICE"
                        elif ticker == "TSE:3110":
                            display_name = "Nitto Boseki"
                        elif ticker == "TWSE:2330":
                            display_name = "TSM"
                        elif ticker == "TWSE:2454":
                            display_name = "MediaTek"
                        
                        down_text += f"{display_name}({change:.2f}%)，"
                    
                    st.markdown(f'<span class="negative">{down_text}价格下降幅度较大；</span>', unsafe_allow_html=True)
                else:
                    st.write("**📉 价格下降幅度较大的股票 (>2%):**")
                    st.write("今日无跌幅超过2%的股票")
                
                # 价格上涨分析
                up_delta = price_delta.sort_values(ascending=False) * 100
                up_stocks = up_delta[up_delta >= 2]
                
                if len(up_stocks) > 0:
                    st.write("**📈 价格上涨幅度较大的股票 (>2%):**")
                    up_text = ""
                    for ticker, change in up_stocks.items():
                        # 股票名称映射
                        display_name = ticker
                        if ticker == "KOSE:A000660":
                            display_name = "SK Hynix"
                        elif ticker == "TASE:NICE":
                            display_name = "NICE"
                        elif ticker == "TSE:3110":
                            display_name = "Nitto Boseki"
                        elif ticker == "TWSE:2330":
                            display_name = "TSM"
                        elif ticker == "TWSE:2454":
                            display_name = "MediaTek"
                        
                        up_text += f"{display_name}({change:.2f}%)，"
                    
                    st.markdown(f'<span class="positive">{up_text}价格上涨幅度较大；</span>', unsafe_allow_html=True)
                else:
                    st.write("**📈 价格上涨幅度较大的股票 (>2%):**")
                    st.write("今日无涨幅超过2%的股票")
                
                # 贡献度影响分析
                st.write("**🎯 对基金表现影响最大的持仓:**")
                
                if total_impact > 0:
                    # 总影响为正，显示贡献最大的股票
                    top_contributors = price_impact.sort_values(ascending=False)[:5]
                    impact_text = ""
                    for ticker, impact in top_contributors.items():
                        # 股票名称映射
                        display_name = ticker
                        if ticker == "KOSE:A000660":
                            display_name = "SK Hynix"
                        elif ticker == "TASE:NICE":
                            display_name = "NICE"
                        elif ticker == "TSE:3110":
                            display_name = "Nitto Boseki"
                        elif ticker == "TWSE:2330":
                            display_name = "TSM"
                        elif ticker == "TWSE:2454":
                            display_name = "MediaTek"
                        
                        # 获取对应的价格变化
                        price_change = price_delta[ticker] * 100
                        impact_text += f"{display_name}({price_change:.2f}%)，"
                    
                    st.markdown(f'<span class="positive">{impact_text}为对指数表现上涨影响最大的前五持仓；</span>', unsafe_allow_html=True)
                else:
                    # 总影响为负，显示拖累最大的股票
                    bottom_contributors = price_impact.sort_values()[:5]
                    impact_text = ""
                    for ticker, impact in bottom_contributors.items():
                        # 股票名称映射
                        display_name = ticker
                        if ticker == "KOSE:A000660":
                            display_name = "SK Hynix"
                        elif ticker == "TASE:NICE":
                            display_name = "NICE"
                        elif ticker == "TSE:3110":
                            display_name = "Nitto Boseki"
                        elif ticker == "TWSE:2330":
                            display_name = "TSM"
                        elif ticker == "TWSE:2454":
                            display_name = "MediaTek"
                        
                        # 获取对应的价格变化
                        price_change = price_delta[ticker] * 100
                        impact_text += f"{display_name}({price_change:.2f}%)，"
                    
                    st.markdown(f'<span class="negative">{impact_text}为对指数表现下降影响最大的前五持仓；</span>', unsafe_allow_html=True)
                

            else:
                st.warning("无法计算贡献度，请检查数据")
        except Exception as e:
            st.error(f"贡献度分析错误: {e}")
    
    # 持仓详情标签页
    with tab6:
        st.header("📋 持仓详情")
        
        if daily_holdings is not None:
            # 显示持仓表格
            # 确保数据类型兼容，避免Arrow序列化问题
            display_holdings = daily_holdings.copy()
            
            # 过滤掉ticker是nan的行
            ticker_col = None
            possible_ticker_cols = ['Ticker', 'Symbol', '股票代码', '代码']
            for col in display_holdings.columns:
                if any(ticker in str(col) for ticker in possible_ticker_cols):
                    ticker_col = col
                    break
            
            if ticker_col is not None:
                # 过滤掉ticker是nan的行
                display_holdings = display_holdings.dropna(subset=[ticker_col])
                # 过滤掉ticker是'nan'字符串的行
                display_holdings = display_holdings[display_holdings[ticker_col] != 'nan']
            
            # 过滤掉日期时间列（包含时间戳的列）
            date_columns = []
            for col in display_holdings.columns:
                if isinstance(col, str) and ('2025-' in col or '2024-' in col):
                    date_columns.append(col)
                elif hasattr(col, 'strftime'):  # 检查是否是datetime对象
                    date_columns.append(col)
            
            # 删除日期时间列
            if date_columns:
                display_holdings = display_holdings.drop(columns=date_columns)
            
            # 确保所有列都是字符串类型
            for col in display_holdings.columns:
                if display_holdings[col].dtype == 'object':
                    display_holdings[col] = display_holdings[col].astype(str)
            
            st.dataframe(display_holdings, use_container_width=True)
            
            # 如果有行业信息，显示行业配置图表
            if 'Sector' in daily_holdings.columns:
                st.subheader("行业配置")
                
                # 创建两列布局
                col1, col2 = st.columns(2)
                
                with col1:
                    # 按公司数量的行业配置
                    sector_fig = plot_sector_allocation(daily_holdings)
                    if sector_fig:
                        st.plotly_chart(sector_fig, use_container_width=True)
                    else:
                        st.write("无法生成按公司数量的行业配置图")
                
                with col2:
                    # 按市值的行业配置
                    market_cap_fig = plot_sector_market_cap(daily_holdings)
                    if market_cap_fig:
                        st.plotly_chart(market_cap_fig, use_container_width=True)
                    else:
                        st.write("无法生成按市值的行业配置图")
                
                # 行业贡献表格
                st.subheader("DTD行业贡献")
                contribution_df = calculate_sector_contribution(daily_holdings)
                if contribution_df is not None:
                    st.dataframe(contribution_df, use_container_width=True)
                else:
                    st.write("无法生成行业贡献表格")
        else:
            st.warning("无法加载每日监控持仓详情数据")
            st.write("请检查 'ciq reference/data/每日数据监控.xlsx' 文件中的 'holdings' 工作表是否存在")
        

if __name__ == "__main__":
    main() 