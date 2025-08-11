# 📊 AGIX Fund Monitor

一个基于Streamlit的AGIX基金监控应用，提供全面的基金表现分析、风险监控和持仓贡献度分析功能。

## 🌐 在线访问

**🎯 立即体验**: [https://funds-monitor-ciq-v3.streamlit.app/](https://funds-monitor-ciq-v3.streamlit.app/)

> 💡 无需安装，直接在浏览器中访问即可使用所有功能！

## 🚀 功能特性

### 📊 概览页面
- **关键指标展示**: 当前净值、持仓数量、年化收益率、年化波动率
- **基金基本信息**: 基金资产规模等基本信息
- **AGIX表现总结**: 日收益率、周收益率、年收益率、2024年收益
- **智能计算**: 支持短期数据年化计算，自动处理数据不足情况

### 📊 基金对比
- **基金收益率对比表格**: AGIX与主要基金的详细对比数据
- **收益率对比图表**: 交互式图表展示不同基金的表现对比
- **多选器支持**: 用户可选择要对比的指数/基金
- **Capital IQ Ticker**: 完整的股票代码信息显示

### 🎯 贡献度分析
- **累计收益率分析**: 30天和90天累计收益率条形图对比
- **累计收益率详细数据**: 包含权重信息的完整数据表格
- **自2025年初累计收益率对比**: 平滑曲线图展示长期表现趋势
- **持仓贡献度分析**: 各持仓对基金表现的贡献度可视化
- **昨日回报详细贡献度数据**: 价格变化、权重、价格影响、贡献度百分比
- **多选器功能**: 独立的多选器支持不同图表选择不同股票组合
- **价格变化分析**: 涨幅和跌幅最大的股票排名
- **详细分析报告**: 对基金表现影响最大的持仓分析

### 📈 收益率分析
- **多期间收益率**: 1天、5天、30天、90天、1年收益率计算
- **收益率热力图**: 直观展示不同期间和持仓的收益率分布
- **数据表格**: 详细的收益率数据展示

### ⚠️ 风险分析
- **风险指标**: 年化波动率、最大回撤、VaR(95%)、Sharpe比率
- **回撤分析**: 基金回撤趋势图表
- **风险监控**: 实时风险指标更新

### 📋 持仓详情
- **当前持仓**: 完整的持仓信息表格
- **行业配置**: 行业分布饼图
- **持仓权重**: 详细的持仓权重数据

## 🛠️ 安装和运行

### 🌐 在线使用（推荐）
**无需安装，直接访问**: [https://funds-monitor-ciq-v3.streamlit.app/](https://funds-monitor-ciq-v3.streamlit.app/)

### 💻 本地运行

#### 1. 安装依赖
```bash
pip install -r requirements.txt
```

#### 2. 运行应用
```bash
streamlit run app.py
```

#### 3. 访问应用
打开浏览器访问: `http://localhost:8501`

## 📁 数据文件结构

应用需要以下数据文件，请确保文件路径正确：

```
ciq reference/
├── data/
│   ├── 每日数据监控.xlsx      # 包含chart、raw1、holdings工作表
│   ├── StockPriceValue.xlsx   # 股票价格数据
│   └── Shares.xlsx           # 包含shares和07_30_2025_agix_holdings工作表
```

### 数据文件说明

1. **每日数据监控.xlsx**
   - `chart` 工作表: AGIX与基准指数的对比数据，包含Capital IQ Ticker
   - `raw1` 工作表: 基金基本信息数据
   - `holdings` 工作表: 持仓详细信息

2. **StockPriceValue.xlsx**
   - 包含所有持仓股票的历史价格数据
   - 支持累计收益率计算

3. **Shares.xlsx**
   - `shares` 工作表: 持仓权重和数量信息
   - `07_30_2025_agix_holdings` 工作表: 当前持仓详细信息

## 🎨 界面特色

- **响应式设计**: 适配不同屏幕尺寸
- **交互式图表**: 基于Plotly的可交互图表
- **实时数据**: 数据缓存和实时更新
- **中文界面**: 完全中文化的用户界面
- **专业配色**: 金融数据可视化专业配色方案
- **多选器支持**: 灵活选择要分析的股票组合
- **智能默认**: 自动选择表现最好的持仓作为默认显示

## 🔧 技术栈

- **前端框架**: Streamlit
- **数据处理**: Pandas, NumPy
- **数据可视化**: Plotly
- **数据源**: Excel文件 (通过openpyxl和xlrd)

## 📊 核心算法

### 收益率计算
```python
returns = (current_price - historical_price) / historical_price
```

### 累计收益率计算
```python
cumulative_returns = (price_series - base_price) / base_price * 100
```

### 年化收益率计算
```python
# 短期数据年化
annual_return = (1 + total_return) ** (252 / days_available) - 1
```

### 风险指标计算
- **年化波动率**: `std(returns) * sqrt(252)`
- **最大回撤**: `min((cumulative_returns - rolling_max) / rolling_max)`
- **VaR(95%)**: `quantile(returns, 0.05)`
- **Sharpe比率**: `(mean_returns - risk_free_rate) / std_returns * sqrt(252)`

### 贡献度分析
```python
price_impact = price_change * shares_held
contribution = price_impact / total_portfolio_value
```

## 🚀 使用指南

### 🌐 在线使用（推荐）
1. **访问应用**: 打开 [https://funds-monitor-ciq-v3.streamlit.app/](https://funds-monitor-ciq-v3.streamlit.app/)
2. **开始使用**: 无需安装，直接体验所有功能

### 💻 本地使用
1. **启动应用**: 运行 `streamlit run app.py`
2. **访问应用**: 打开浏览器访问 `http://localhost:8501`

### 📊 功能使用
1. **概览页面**: 查看关键指标和AGIX表现总结
2. **基金对比**: 选择要对比的基金，查看对比图表
3. **贡献度分析**: 
   - 使用多选器选择要分析的股票
   - 查看累计收益率分析
   - 查看自2025年初的长期表现趋势
   - 分析持仓贡献度
4. **收益率分析**: 查看多期间收益率和热力图
5. **风险分析**: 查看风险指标和回撤分析
6. **持仓详情**: 查看完整的持仓信息

## 🔄 数据更新

应用会自动读取Excel文件中的数据，以下是详细的数据更新流程：

### 📅 每日更新流程

#### 1. 更新每日数据监控文件
- **raw1 sheet**: 更新上个交易日，n-1日期，week，month等，记得保存文件


#### 2. 更新StockPriceValue文件
- **插入新行**: 在文件最上面插入一行新的日期
- **公式填充**: 拉动下面所有公式进行填充
- **重要**: 必须保存文件，否则代码无法读取新数据

#### 3. 更新Shares文件
- **更改日期刷新**: 更新标黄日期，记得保存文件

## 🆙 项目接管指南 - 成为新的Admin

如果您需要接管这个项目并成为新的管理员，请按照以下步骤操作：

### 📥 第一步：下载项目代码

#### 方法1：从GitHub下载（推荐）
1. 访问 [https://github.com/Simondjy/funds-monitor-ciq](https://github.com/Simondjy/funds-monitor-ciq)
2. 点击绿色的 "Code" 按钮
3. 选择 "Download ZIP"
4. 解压到您的本地文件夹

#### 方法2：使用Git克隆（如果您熟悉Git）
```bash
git clone https://github.com/Simondjy/funds-monitor-ciq.git
cd funds-monitor-ciq
```

### 🔐 第二步：创建自己的GitHub仓库

1. 登录您的GitHub账户
2. 点击右上角 "+" 号，选择 "New repository"
3. 输入仓库名称（例如：`my-funds-monitor`）
4. 选择 "Public" 或 "Private"
5. 不要勾选 "Add a README file"（因为我们已经有了）
6. 点击 "Create repository"

### 📤 第三步：上传代码到您的仓库

#### 方法1：使用Git（推荐）
```bash
# 在项目文件夹中初始化Git
git init
git add .
git commit -m "Initial commit: AGIX Fund Monitor"

# 添加您的远程仓库
git remote add origin https://github.com/您的用户名/您的仓库名.git
git branch -M main
git push -u origin main
```

#### 方法2：直接上传文件
1. 在您的GitHub仓库页面，点击 "uploading an existing file"
2. 拖拽所有项目文件到上传区域
3. 添加提交信息，点击 "Commit changes"

### ☁️ 第四步：创建自己的Streamlit Cloud应用

1. 访问 [https://share.streamlit.io/](https://share.streamlit.io/)
2. 使用GitHub账户登录
3. 点击 "New app"
4. 填写应用信息：
   - **Repository**: 选择您刚创建的仓库
   - **Branch**: 选择 `main`
   - **Main file path**: 输入 `app.py`
   - **App URL**: 输入您想要的URL（例如：`my-funds-monitor`）
5. 点击 "Deploy!"

### 🔄 第五步：设置自动更新

#### 每日数据更新流程：
1. **准备数据文件**：
   - 更新 `ciq reference/data/` 文件夹中的Excel文件
   - 确保文件名和格式与原始项目一致

2. **上传到GitHub**：
   ```bash
   git add .
   git commit -m "Update daily data: $(date)"
   git push origin main
   ```

3. **Streamlit自动部署**：
   - Streamlit Cloud会自动检测到代码更新
   - 应用会在几分钟内自动重新部署
   - 新数据会立即生效

### 📋 第六步：自定义配置


#### 修改数据文件路径：
如果需要，可以调整数据文件的路径结构：
```python
# 在app.py中搜索这些路径并修改
data_path = "ciq reference/data/"
```

5. **测试部署**：在正式使用前，先在测试环境中验证功能


### 📞 技术支持

如果在接管过程中遇到问题：
1. 查看Streamlit官方文档
2. 检查GitHub Issues页面
3. 联系原项目维护者

---

## 📝 注意事项

- 确保Excel文件格式正确，包含所需的工作表
- 数据文件路径必须正确配置
- 建议定期备份原始数据文件
- 应用支持的数据格式: .xlsx, .xls
- 累计收益率分析需要足够的历史数据
- 年化收益率计算会自动处理数据不足的情况

## 🆕 最新更新

- ✅ **在线部署**: 应用已成功部署到Streamlit Cloud，支持在线访问
- ✅ 修复了Capital IQ Ticker列显示为空的问题
- ✅ 新增累计收益率分析功能
- ✅ 新增自2025年初累计收益率对比图表
- ✅ 添加了多选器支持，用户可灵活选择股票
- ✅ 优化了年化收益率计算，支持短期数据
- ✅ 改进了错误处理和数据验证
- ✅ 统一了表格颜色显示规则
- ✅ 优化了资金流量列的红绿显示逻辑
- ✅ **新增项目接管指南**: 详细说明如何成为新的Admin并创建自己的应用

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个应用！

## �� 许可证

MIT License 