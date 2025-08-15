#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
更新Shares.xlsx文件中指定工作表的脚本
下载AGIX Holdings数据并更新到Shares文件中
"""

import pandas as pd
from pathlib import Path
import openpyxl
from openpyxl import Workbook
import requests
from datetime import datetime, timedelta
import pytz
import win32com.client
import time

# 导入公共日期处理模块
from date_utils import calculate_n_and_n_minus_1

# 公司名称到股票代码的映射（请在这里添加您的MAPPING）
COMPANY_TO_TICKER_ADD = {
    "XAI HOLDINGS CORP": "NA",
    "ANTHROPIC, PBC": "NA",
    "ALPHABET INC-CL A": "NasdaqGS:GOOGL",
    "TSMC": "TWSE:2330",
    "SK HYNIX INC": "KOSE:A000660",
    "ASML HOLDING NV": "NasdaqGS:ASML",
    "ARISTA NETWORKS INC": "ANET",
    "MEDIATEK INC": "TWSE:2454",
    # 可继续补充
}

def is_trading_day(date):
    """判断是否为交易日（周一至周五）"""
    return date.weekday() < 5

def get_previous_trading_day(date):
    """获取指定日期的前一个交易日"""
    current = date
    while True:
        current = current - timedelta(days=1)
        if is_trading_day(current):
            return current

def get_us_market_date():
    """
    获取美国市场对应的日期
    考虑时区差异：北京时间比美东时间快12-13小时
    应该使用美东时间的日期作为基准
    """
    # 获取当前北京时间
    beijing_tz = pytz.timezone('Asia/Shanghai')
    beijing_time = datetime.now(beijing_tz)
    
    # 获取当前美东时间
    us_eastern_tz = pytz.timezone('America/New_York')
    us_eastern_time = datetime.now(us_eastern_tz)
    
    print(f"北京时间: {beijing_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"美东时间: {us_eastern_time.strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 使用美东时间的日期作为基准
    us_date = us_eastern_time.date()
    print(f"使用美东日期: {us_date}")
    return us_date

def is_market_closed_for_today():
    """
    判断今天是否已经收市
    如果当前时间已经过了今天的交易时间，返回True
    这样就能确保在收市后使用当天的数据（因为CIQ数据已经有了）
    """
    us_eastern_tz = pytz.timezone('America/New_York')
    us_eastern_time = datetime.now(us_eastern_tz)
    
    # 检查是否为工作日
    if us_eastern_time.weekday() >= 5:
        return True  # 周末直接认为已收市
    
    # 检查是否已经过了今天的交易时间
    current_time = us_eastern_time.time()
    market_close = datetime.strptime('16:00', '%H:%M').time()
    
    return current_time > market_close

def try_download(csv_name, save_path):
    """尝试从kraneshares官网按指定文件名下载持仓csv文件，下载成功返回True，否则返回False"""
    url = f'https://kraneshares.com/csv/{csv_name}'
    print(f'[INFO] 尝试下载URL: {url}')
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(url, timeout=10, headers=headers)
        response.raise_for_status()
        with open(save_path, 'wb') as f:
            f.write(response.content)
        print(f'[INFO] 下载成功: {save_path}')
        return True
    except Exception as e:
        print(f'[WARN] 下载失败: {url}, 错误: {e}')
        return False

def replace_tickers_in_holdings_file(holdings_csv):
    """读取并修改CSV文件中的Tickers"""
    # 使用更灵活的CSV读取参数，确保所有列都被读取
    df = pd.read_csv(holdings_csv, skiprows=1, encoding='utf-8', on_bad_lines='skip')
    
    print(f"[DEBUG] CSV读取后列数: {len(df.columns)}")
    print(f"[DEBUG] 列名: {list(df.columns)}")
    
    for company, ticker in COMPANY_TO_TICKER_ADD.items():
        df.loc[df['Company Name'] == company, 'Ticker'] = ticker
    
    # 重新保存，保留原有的表头格式
    with open(holdings_csv, 'r', encoding='utf-8') as f:
        header = f.readline()
    
    # 保存时确保所有列都被写入
    df.to_csv(holdings_csv, index=False, mode='w', encoding='utf-8')
    
    # 保留原始的第一行表头
    with open(holdings_csv, 'r+', encoding='utf-8') as f:
        content = f.read()
        f.seek(0, 0)
        f.write(header + content)
    
    print(f"[DEBUG] 处理后的列数: {len(df.columns)}")

def download_agix_holdings(n_date):
    """下载AGIX Holdings数据"""
    # 构建文件名格式：MM_DD_YYYY_agix_holdings.csv
    # 将日期格式从 YYYY-MM-DD 转换为 MM_DD_YYYY
    date_obj = datetime.strptime(n_date, '%Y-%m-%d')
    csv_name = f'{date_obj.strftime("%m_%d_%Y")}_agix_holdings.csv'
    save_path = Path(f'agix_holdings/{csv_name}')
    
    print(f'[INFO] 检查AGIX Holdings数据，日期: {n_date}')
    
    # 检查文件是否已存在
    if save_path.exists():
        print(f'[INFO] 文件已存在，跳过下载: {save_path}')
        
        # 如果文件存在，仍然需要处理Tickers映射
        if COMPANY_TO_TICKER_ADD:
            print('[INFO] 开始处理Tickers映射...')
            replace_tickers_in_holdings_file(save_path)
            print('[INFO] Tickers映射处理完成')
        
        return save_path
    
    print(f'[INFO] 文件不存在，开始下载...')
    
    # 确保data目录存在
    save_path.parent.mkdir(exist_ok=True)
    
    # 尝试下载
    if try_download(csv_name, save_path):
        print(f'[INFO] AGIX Holdings数据下载成功: {save_path}')
        
        # 处理CSV文件中的Tickers
        if COMPANY_TO_TICKER_ADD:
            print('[INFO] 开始处理Tickers映射...')
            replace_tickers_in_holdings_file(save_path)
            print('[INFO] Tickers映射处理完成')
        
        return save_path
    else:
        print(f'[ERROR] AGIX Holdings数据下载失败')
        return None

def update_shares_sheet_with_holdings_data(holdings_csv_path, dates):
    """使用持仓数据更新Shares.xlsx文件（采用另存为方式）"""
    
    # 文件路径
    source_file = Path("process_data/Shares.xlsx").resolve()
    target_file = Path("process_data/Shares_complete.xlsx").resolve()
    
    # 检查源文件是否存在
    if not source_file.exists():
        print(f"❌ 错误: 源文件不存在 - {source_file}")
        return False
    
    print(f"源文件: {source_file}")
    print(f"目标文件: {target_file}")
    
    excel_app = None
    workbook = None
    
    try:
        print("正在启动Excel应用程序...")
        
        # 创建Excel应用程序实例
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False  # 隐藏Excel窗口
        excel_app.DisplayAlerts = False  # 禁用警告对话框
        excel_app.ScreenUpdating = False  # 禁用屏幕更新
        excel_app.EnableEvents = False  # 禁用事件
        excel_app.Interactive = False  # 禁用交互
        
        print("正在打开源文件...")
        
        # 打开源文件
        workbook = excel_app.Workbooks.Open(str(source_file))
        print(f"✅ 成功打开源文件: {source_file}")
        
        # 读取持仓CSV数据 - 使用更灵活的读取方式
        print(f'[INFO] 读取持仓数据: {holdings_csv_path}')
        
        # 先尝试直接读取，看看原始结构
        try:
            # 读取原始CSV文件的前几行来了解结构
            with open(holdings_csv_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                print(f"[DEBUG] CSV文件总行数: {len(lines)}")
                print(f"[DEBUG] 前3行内容:")
                for i, line in enumerate(lines[:3]):
                    print(f"  行{i+1}: {line.strip()}")
        except Exception as e:
            print(f"[WARN] 无法读取原始CSV文件: {e}")
        
        # 使用更灵活的CSV读取参数
        df_holdings = pd.read_csv(
            holdings_csv_path, 
            skiprows=1,  # 跳过第一行（通常是描述性文字）
            encoding='utf-8',
            on_bad_lines='skip',  # 跳过有问题的行
            engine='python'  # 使用Python引擎，更灵活
        )
        
        print(f'[INFO] 持仓数据包含 {len(df_holdings)} 行')
        print(f'[INFO] 持仓数据包含 {len(df_holdings.columns)} 列')
        print(f'[INFO] 列名: {list(df_holdings.columns)}')
        
        # 检查是否有空列或问题列
        for i, col in enumerate(df_holdings.columns):
            if pd.isna(col) or col == '':
                print(f"[WARN] 发现空列名，列索引: {i}")
            print(f"列{i}: '{col}' - 数据类型: {df_holdings[col].dtype}")
        
        # 获取或创建工作表
        sheet_name = "agix_holdings_raw"
        try:
            worksheet = workbook.Worksheets(sheet_name)
            print(f"找到工作表：{sheet_name}")
            
            # 清除所有现有内容
            last_row = worksheet.UsedRange.Rows.Count
            if last_row > 0:
                worksheet.Range(f"A1:Z{last_row}").ClearContents()
                print(f"已清除工作表内容")
            
        except:
            print(f"工作表 {sheet_name} 不存在，正在创建...")
            worksheet = workbook.Worksheets.Add()
            worksheet.Name = sheet_name
        
        # 准备数据写入 - 排除Identifier列
        print(f"[DEBUG] 准备写入数据，行数: {len(df_holdings)}, 列数: {len(df_holdings.columns)}")
        
        # 获取列名（表头），排除Identifier列
        headers = [col for col in df_holdings.columns if col != 'Identifier']
        print(f"[DEBUG] 过滤后的表头: {headers}")
        
        # 获取数据值，排除Identifier列
        data_values = df_holdings[headers].values.tolist()
        print(f"[DEBUG] 数据行数: {len(data_values)}")
        
        # 组合表头和数据
        data = [headers] + data_values
        
        if len(data) > 0 and len(data[0]) > 0:
            nrows = len(data)
            ncols = len(data[0])
            
            print(f"[DEBUG] 最终数据矩阵: {nrows}行 x {ncols}列")
            
            # 计算Excel范围，确保包含所有列
            def get_excel_col(col_idx):
                col_str = ""
                while col_idx > 0:
                    col_idx, rem = divmod(col_idx - 1, 26)
                    col_str = chr(65 + rem) + col_str
                return col_str
            
            end_col = get_excel_col(ncols)
            end_row = nrows
            excel_range = f"A1:{end_col}{end_row}"
            
            print(f"[DEBUG] Excel写入范围: {excel_range}")
            
            # 写入数据
            worksheet.Range(excel_range).Value = data
            
            # 验证写入结果
            actual_range = worksheet.UsedRange
            actual_rows = actual_range.Rows.Count
            actual_cols = actual_range.Columns.Count
            
            print(f"[DEBUG] 实际写入结果: {actual_rows}行 x {actual_cols}列")
            
            if actual_cols != ncols:
                print(f"[WARN] 列数不匹配！期望: {ncols}, 实际: {actual_cols}")
            else:
                print(f"[INFO] 列数匹配，写入成功")
        
        print(f"已成功写入 {len(df_holdings)} 行持仓数据到工作表 {sheet_name}")
        
        # 更新shares工作表
        print("正在更新shares工作表...")
        try:
            shares_worksheet = workbook.Worksheets("shares")
            print("找到shares工作表，开始更新...")
            
            # 写入日期信息
            shares_worksheet.Range("E2").Value = dates['n']
            shares_worksheet.Range("E3").Value = dates['n-1']
            print(f"已更新E2为: {dates['n']}")
            print(f"已更新E3为: {dates['n-1']}")
            
            # 更新P2单元格
            shares_worksheet.Range("P2").Value = "1700002"
            print("已更新P2为: 1700002")
            
            print("✅ shares工作表更新完成")
            
        except Exception as e:
            print(f"⚠️ 更新shares工作表时出错: {e}")
            print("继续执行另存为操作...")
        
        # 另存为新文件（关键操作：另存为而不是保存）
        print(f"正在另存为新文件: {target_file}")
        workbook.SaveAs(str(target_file))
        print(f"✅ 文件已成功另存为: {target_file}")
        
        # 关闭工作簿和Excel
        workbook.Close()
        excel_app.Quit()
        
        # 重新打开文件让数据重新加载
        print("\n=== 重新打开文件让数据重新加载 ===")
        print("正在重新启动Excel应用程序...")
        
        # 创建新的Excel应用程序实例
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False
        excel_app.EnableEvents = False
        excel_app.Interactive = False
        
        print("正在重新打开保存的文件...")
        workbook = excel_app.Workbooks.Open(str(target_file))
        print(f"✅ 成功重新打开文件: {target_file}")
        
        # 获取工作表
        try:
            agix_worksheet = workbook.Worksheets("agix_holdings_raw")
            shares_worksheet = workbook.Worksheets("shares")
            print("✅ 成功获取所有工作表")
        except Exception as e:
            print(f"⚠️ 获取工作表时出错: {e}")
        
        # 等待数据加载完成
        print("等待数据加载完成...")
        time.sleep(5)  # 给数据加载一些时间
        
        # 保存并关闭
        print("正在保存并关闭文件...")
        try:
            workbook.Save()
            workbook.Close()
            excel_app.Quit()
            print("✅ 数据重新加载完成，文件已关闭")
        except:
            pass
        
        print("✅ Shares文件更新完成！")
        return True
        
    except Exception as e:
        print(f"❌ 更新过程中出错: {e}")
        import traceback
        traceback.print_exc()
        
        # 尝试关闭Excel
        try:
            if workbook:
                workbook.Close(SaveChanges=False)
            if excel_app:
                excel_app.Quit()
        except:
            pass
        
        return False
    
    finally:
        # 清理COM对象
        try:
            if excel_app:
                # 恢复Excel设置
                excel_app.ScreenUpdating = True
                excel_app.EnableEvents = True
                excel_app.Interactive = True
                excel_app.DisplayAlerts = True
            if workbook:
                workbook = None
            if excel_app:
                excel_app = None
        except:
            pass

def main():
    """主函数"""
    print("=== 开始更新AGIX Holdings数据 ===")
    
    # 1. 计算n和n-1日期
    print("\n1. 计算日期...")
    dates = calculate_n_and_n_minus_1()
    print(f"n: {dates['n']}")
    print(f"n-1: {dates['n-1']}")
    
    # 2. 下载AGIX Holdings数据
    print("\n2. 下载AGIX Holdings数据...")
    holdings_csv_path = download_agix_holdings(dates['n'])
    
    if holdings_csv_path is None:
        print("[ERROR] 下载失败，程序终止")
        return
    
    # 3. 更新Shares.xlsx文件
    print("\n3. 更新Shares.xlsx文件...")
    if update_shares_sheet_with_holdings_data(holdings_csv_path, dates):
        print("[SUCCESS] 所有操作完成！")
    else:
        print("[ERROR] 更新Shares文件失败")

if __name__ == "__main__":
    main()
