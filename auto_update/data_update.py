#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
统一数据更新脚本
整合所有数据更新功能：
1. 更新每日监控数据 (update_monitoring)
2. 更新Shares数据 (update_shares_sheet) 
3. 生成基金价格数据 (updatefundsvalue)
4. 生成股票价格数据 (updatestockprice)

使用统一的日期处理逻辑，确保所有数据使用相同的日期基准
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
import subprocess
import os

# 导入公共日期处理模块
from date_utils import (
    calculate_dates, 
    calculate_n_and_n_minus_1, 
    get_latest_trading_day, 
    generate_trading_dates
)

# 基金代码列表
FUND_TICKERS = [
    "NasdaqGM:QQQ", "NasdaqGM:AGIX", "NasdaqGM:SMH", "BATS:IGV", 
    "NasdaqGM:BOTZ", "NasdaqGM:AIQ", "ARCA:ARTY", "NasdaqGM:ROBT", "ARCA:IGPT", 
    "BATS:WTAI", "ARCA:THNQ", "NasdaqGM:FDTX", "ARCA:CHAT", "ARCA:LOUP", 
    "ARCA:LRNZ", "ARCA:AIS", "NasdaqGM:WISE", "LSE:RBOT", "XTRA:XAIX", 
    "BIT:WTAI", "LSE:AIAG", "ASX:RBTZ", "DB:XB0T"
]

# 股票代码列表
STOCK_TICKERS = [
    "KOSE:A000660", "TWSE:2330", "TWSE:2454",
    "AAPL", "ADBE", "AMZN", "ANET", "ARM", "NasdaqGS:ASML", 
    "AVGO", "CFLT", "CRM", "DDOG", "APP", "DUOL", "ESTC", "NasdaqGS:GOOGL", "GTLB", "IOT", "MDB", "META", "MRVL",
    "MSFT", "MU", "NBIS", "NET", "NOW", "NVDA", "ORCL", "PANW", "PLTR", "PSTG", "QCOM", "RBLX", "SAP", "SHOP", "SNOW", "SNPS", "TEAM", "TEM", "TSLA", "VRT", "WDAY", "ZS"
]

# 公司名称到股票代码的映射
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

def update_monitoring_data():
    """
    更新每日监控数据：
    - 源文件：每日数据监控.xlsx
    - 工作表：raw1
    - 自动计算并更新：A13=n, A14=n-1, A15=当周第一个交易日, A16=本月第一个交易日
    - 目标文件：每日监控数据_complete.xlsx
    """
    print("\n" + "="*60)
    print("1. 开始更新每日监控数据")
    print("="*60)
    
    # 文件路径
    source_file = Path("process_data/每日数据监控.xlsx").resolve()
    target_file = Path("process_data/每日监控数据_complete.xlsx").resolve()
    
    # 检查源文件是否存在
    if not source_file.exists():
        print(f"❌ 错误: 源文件不存在 - {source_file}")
        return False
    
    # 计算日期
    print("\n=== 计算日期 ===")
    dates = calculate_dates()
    print(f"n: {dates['n']}")
    print(f"n-1: {dates['n-1']}")
    print(f"当周第一个交易日: {dates['当周第一个交易日']}")
    print(f"本月第一个交易日: {dates['本月第一个交易日']}")
    
    excel_app = None
    workbook = None
    
    try:
        print("正在启动Excel应用程序...")
        
        # 创建Excel应用程序实例
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False
        excel_app.EnableEvents = False
        excel_app.Interactive = False
        
        print("正在打开Excel文件...")
        workbook = excel_app.Workbooks.Open(str(source_file))
        print(f"✅ 成功打开Excel文件: {source_file}")
        
        # 获取raw1工作表
        worksheet = workbook.Worksheets("raw1")
        
        # 更新单元格
        worksheet.Range("A13").Value = dates['n']
        worksheet.Range("A14").Value = dates['n-1']
        worksheet.Range("A15").Value = dates['当周第一个交易日']
        worksheet.Range("A16").Value = dates['本月第一个交易日']
        
        print("✅ 已更新所有日期单元格")
        
        # 刷新所有数据连接
        try:
            print("正在刷新数据连接...")
            for i, connection in enumerate(workbook.Connections, 1):
                print(f"  刷新连接 {i}/{len(workbook.Connections)}: {connection.Name}")
                connection.Refresh()
            print("✅ 已刷新所有数据连接")
        except Exception as conn_error:
            print(f"⚠️  数据连接刷新跳过: {conn_error}")
        
        # 等待数据连接完成
        print("等待数据连接完成...")
        time.sleep(5)
        
        # 另存为新文件
        print(f"正在保存为新文件: {target_file}")
        workbook.SaveAs(str(target_file))
        print(f"✅ 文件已成功保存为: {target_file}")
        
        # 关闭工作簿和Excel
        workbook.Close()
        excel_app.Quit()
        
        print("✅ 每日监控数据更新成功！")
        
        # 显式打开Excel文件进行最终确认
        print("\n=== 显式打开Excel文件进行最终确认 ===")
        
        try:
            print("正在使用系统默认程序打开Excel文件...")
            
            # 使用系统默认程序打开Excel文件
            import os
            import subprocess
            
            # 尝试使用系统默认程序打开
            try:
                os.startfile(str(target_file))
                print(f"✅ 已使用系统默认程序打开文件: {target_file}")
            except:
                # 如果os.startfile失败，尝试使用subprocess
                subprocess.Popen(['start', str(target_file)], shell=True)
                print(f"✅ 已使用subprocess打开文件: {target_file}")
            
            # 等待文件打开
            print("等待Excel文件打开...")
            time.sleep(3)
            
            # 使用pyautogui进行自动操作（如果可用）
            try:
                import pyautogui
                print("正在执行自动保存和关闭操作...")
                
                # 等待一下确保Excel完全加载
                time.sleep(30)
                
                # 按Ctrl+S保存
                pyautogui.hotkey('ctrl', 's')
                print("✅ 已执行保存操作")
                time.sleep(2)
                
                # 按Alt+F4关闭
                pyautogui.hotkey('alt', 'f4')
                print("✅ 已执行关闭操作")
                
            except ImportError:
                print("⚠️  pyautogui未安装，无法自动保存和关闭")
                print("请手动保存并关闭Excel文件")
                input("按回车键继续...")
            except Exception as auto_error:
                print(f"⚠️  自动操作失败: {auto_error}")
                print("请手动保存并关闭Excel文件")
                input("按回车键继续...")
                
        except Exception as final_error:
            print(f"⚠️  打开Excel文件时出现错误: {final_error}")
        
        return True
        
    except Exception as e:
        print(f"❌ 更新过程中出错: {e}")
        import traceback
        traceback.print_exc()
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

def try_download(csv_name, save_path):
    """尝试从kraneshares官网按指定文件名下载持仓csv文件"""
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
    df = pd.read_csv(holdings_csv, skiprows=1, encoding='utf-8', on_bad_lines='skip')
    
    print(f"[DEBUG] CSV读取后列数: {len(df.columns)}")
    print(f"[DEBUG] 列名: {list(df.columns)}")
    
    for company, ticker in COMPANY_TO_TICKER_ADD.items():
        df.loc[df['Company Name'] == company, 'Ticker'] = ticker
    
    # 重新保存，保留原有的表头格式
    with open(holdings_csv, 'r', encoding='utf-8') as f:
        header = f.readline()
    
    df.to_csv(holdings_csv, index=False, mode='w', encoding='utf-8')
    
    with open(holdings_csv, 'r+', encoding='utf-8') as f:
        content = f.read()
        f.seek(0, 0)
        f.write(header + content)
    
    print(f"[DEBUG] 处理后的列数: {len(df.columns)}")

def download_agix_holdings(n_date):
    """下载AGIX Holdings数据"""
    date_obj = datetime.strptime(n_date, '%Y-%m-%d')
    csv_name = f'{date_obj.strftime("%m_%d_%Y")}_agix_holdings.csv'
    save_path = Path(f'agix_holdings/{csv_name}')
    
    print(f'[INFO] 检查AGIX Holdings数据，日期: {n_date}')
    
    # 检查文件是否已存在
    if save_path.exists():
        print(f'[INFO] 文件已存在，跳过下载: {save_path}')
        
        if COMPANY_TO_TICKER_ADD:
            print('[INFO] 开始处理Tickers映射...')
            replace_tickers_in_holdings_file(save_path)
            print('[INFO] Tickers映射处理完成')
        
        return save_path
    
    print(f'[INFO] 文件不存在，开始下载...')
    
    # 确保目录存在
    save_path.parent.mkdir(exist_ok=True)
    
    # 尝试下载
    if try_download(csv_name, save_path):
        print(f'[INFO] AGIX Holdings数据下载成功: {save_path}')
        
        if COMPANY_TO_TICKER_ADD:
            print('[INFO] 开始处理Tickers映射...')
            replace_tickers_in_holdings_file(save_path)
            print('[INFO] Tickers映射处理完成')
        
        return save_path
    else:
        print(f'[ERROR] AGIX Holdings数据下载失败')
        return None

def update_shares_sheet_with_holdings_data(holdings_csv_path, dates):
    """使用持仓数据更新Shares.xlsx文件"""
    print("\n=== 更新Shares.xlsx文件 ===")
    
    # 文件路径
    source_file = Path("process_data/Shares.xlsx").resolve()
    target_file = Path("process_data/Shares_complete.xlsx").resolve()
    
    # 检查源文件是否存在
    if not source_file.exists():
        print(f"❌ 错误: 源文件不存在 - {source_file}")
        return False
    
    excel_app = None
    workbook = None
    
    try:
        print("正在启动Excel应用程序...")
        
        # 创建Excel应用程序实例
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        excel_app.ScreenUpdating = False
        excel_app.EnableEvents = False
        excel_app.Interactive = False
        
        print("正在打开源文件...")
        workbook = excel_app.Workbooks.Open(str(source_file))
        print(f"✅ 成功打开源文件: {source_file}")
        
        # 读取持仓CSV数据
        print(f'[INFO] 读取持仓数据: {holdings_csv_path}')
        
        df_holdings = pd.read_csv(
            holdings_csv_path, 
            skiprows=1,
            encoding='utf-8',
            on_bad_lines='skip',
            engine='python'
        )
        
        print(f'[INFO] 持仓数据包含 {len(df_holdings)} 行')
        print(f'[INFO] 持仓数据包含 {len(df_holdings.columns)} 列')
        
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
        headers = [col for col in df_holdings.columns if col != 'Identifier']
        data_values = df_holdings[headers].values.tolist()
        data = [headers] + data_values
        
        if len(data) > 0 and len(data[0]) > 0:
            nrows = len(data)
            ncols = len(data[0])
            
            # 计算Excel范围
            def get_excel_col(col_idx):
                col_str = ""
                while col_idx > 0:
                    col_idx, rem = divmod(col_idx - 1, 26)
                    col_str = chr(65 + rem) + col_str
                return col_str
            
            end_col = get_excel_col(ncols)
            end_row = nrows
            excel_range = f"A1:{end_col}{end_row}"
            
            # 写入数据
            worksheet.Range(excel_range).Value = data
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
        
        # 另存为新文件
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

def update_shares_data():
    """更新Shares数据"""
    print("\n" + "="*60)
    print("2. 开始更新Shares数据")
    print("="*60)
    
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
        return False
    
    # 3. 更新Shares.xlsx文件
    print("\n3. 更新Shares.xlsx文件...")
    if update_shares_sheet_with_holdings_data(holdings_csv_path, dates):
        print("[SUCCESS] Shares数据更新完成！")
        return True
    else:
        print("[ERROR] 更新Shares文件失败")
        return False

def generate_funds_value_data():
    """生成基金价格数据"""
    print("\n" + "="*60)
    print("3. 开始生成基金价格数据")
    print("="*60)
    
    output_file = Path("process_data/FundsValue_complete.xlsx")
    
    # 定义日期范围
    start_date = datetime(2024, 12, 20).date()
    end_date = get_latest_trading_day()
    
    print(f"正在生成从 {start_date} 到 {end_date} 的交易日数据...")
    
    # 生成所有交易日
    trading_dates = generate_trading_dates(start_date, end_date)
    print(f"总共找到 {len(trading_dates)} 个交易日")
    
    # 创建数据列表（倒序）
    data_rows = []
    
    for date in reversed(trading_dates):
        date_str = date.strftime('%Y-%m-%d')
        row_data = {'Date': date_str}
        
        # 为每个基金添加CIQ函数
        for ticker in FUND_TICKERS:
            row_data[ticker] = f'=@CIQ("{ticker}", "IQ_CLOSEPRICE", "{date_str}", "USD")'
        
        data_rows.append(row_data)
    
    # 创建DataFrame
    df = pd.DataFrame(data_rows)
    
    # 保存到文件
    df.to_excel(output_file, index=False, sheet_name="Price")
    
    print(f"成功创建文件：{output_file}")
    print(f"包含 {len(df)} 行数据，{len(df.columns)} 列")
    print(f"日期范围：{df['Date'].iloc[0]} 到 {df['Date'].iloc[-1]}")
    
    # 询问用户是否要自动打开Excel等待CIQ刷新
    print("\n" + "="*50)
    print("数据文件已生成完成！")
    print("现在将自动打开Excel文件等待CIQ插件刷新数据...")
    print("="*50)
    
    # 自动打开Excel并等待CIQ刷新
    excel_path = str(output_file.absolute())
    if open_excel_and_wait(excel_path, wait_time=60):
        print("Excel文件处理完成！CIQ数据已刷新。")
        return True
    else:
        print("Excel文件处理过程中出现错误，请手动检查。")
        return False

def generate_stock_price_data():
    """生成股票价格数据"""
    print("\n" + "="*60)
    print("4. 开始生成股票价格数据")
    print("="*60)
    
    output_file = Path("process_data/StockPriceValue_complete.xlsx")
    
    # 定义日期范围
    start_date = datetime(2024, 12, 20).date()
    end_date = get_latest_trading_day()
    
    print(f"正在生成从 {start_date} 到 {end_date} 的交易日数据...")
    
    # 生成所有交易日
    trading_dates = generate_trading_dates(start_date, end_date)
    print(f"总共找到 {len(trading_dates)} 个交易日")
    
    # 创建数据列表（倒序）
    data_rows = []
    
    for date in reversed(trading_dates):
        date_str = date.strftime('%Y-%m-%d')
        row_data = {'Date': date_str}
        
        # 为每个股票添加CIQ函数
        for ticker in STOCK_TICKERS:
            row_data[ticker] = f'=@CIQ("{ticker}", "IQ_CLOSEPRICE", "{date_str}", "USD")'
        
        data_rows.append(row_data)
    
    # 创建DataFrame
    df = pd.DataFrame(data_rows)
    
    # 保存到文件
    df.to_excel(output_file, index=False, sheet_name="Price")
    
    print(f"成功创建文件：{output_file}")
    print(f"包含 {len(df)} 行数据，{len(df.columns)} 列")
    print(f"日期范围：{df['Date'].iloc[0]} 到 {df['Date'].iloc[-1]}")
    
    # 询问用户是否要自动打开Excel等待CIQ刷新
    print("\n" + "="*50)
    print("数据文件已生成完成！")
    print("现在将自动打开Excel文件等待CIQ插件刷新数据...")
    print("="*50)
    
    # 自动打开Excel并等待CIQ刷新
    excel_path = str(output_file.absolute())
    if open_excel_and_wait(excel_path, wait_time=60):
        print("Excel文件处理完成！CIQ数据已刷新。")
        return True
    else:
        print("Excel文件处理过程中出现错误，请手动检查。")
        return False

def wait_for_ciq_refresh(excel_file_path, wait_time=60):
    """
    等待CIQ插件刷新数据
    Args:
        excel_file_path: Excel文件路径
        wait_time: 等待时间（秒），默认60秒
    """
    print(f"正在等待CIQ插件刷新数据，预计等待时间：{wait_time}秒...")
    
    # 等待指定时间让CIQ插件完成数据刷新
    for i in range(wait_time):
        if i % 10 == 0:  # 每10秒显示一次进度
            remaining = wait_time - i
            print(f"剩余等待时间：{remaining}秒...")
        time.sleep(1)
    
    print("CIQ数据刷新等待完成")

def open_excel_and_wait(file_path, wait_time=60):
    """
    打开Excel文件并等待CIQ插件刷新，然后自动保存数据
    Args:
        file_path: Excel文件路径
        wait_time: 等待时间（秒）
    """
    try:
        # 检查文件是否存在
        if not os.path.exists(file_path):
            print(f"错误：文件 {file_path} 不存在")
            return False
        
        print(f"正在打开Excel文件：{file_path}")
        print("注意：请确保Excel中的CIQ公式能够正常获取数据")
        
        # 使用默认程序打开Excel文件
        if os.name == 'nt':  # Windows系统
            os.startfile(file_path)
        else:  # Linux/Mac系统
            subprocess.run(['xdg-open', file_path])
        
        # 等待CIQ插件刷新
        wait_for_ciq_refresh(file_path, wait_time)
        
        print("CIQ数据刷新完成！正在尝试自动保存文件...")
        
        # 尝试自动保存文件
        save_success = False
        if os.name == 'nt':  # Windows系统
            try:
                import win32com.client
                print("正在通过COM接口自动保存Excel文件...")
                
                # 等待一下确保Excel完全加载
                time.sleep(3)
                
                # 连接到Excel应用程序
                excel = win32com.client.GetObject(Class="Excel.Application")
                if excel is None:
                    excel = win32com.client.Dispatch("Excel.Application")
                
                # 获取当前打开的工作簿
                workbook = excel.ActiveWorkbook
                if workbook is not None:
                    # 保存工作簿
                    workbook.Save()
                    print("✅ 文件已自动保存成功！")
                    save_success = True
                    
                    # 关闭工作簿
                    workbook.Close(SaveChanges=False)
                else:
                    print("⚠️ 未找到打开的工作簿，尝试其他保存方法...")
                
                # 退出Excel应用程序
                excel.Quit()
                
            except ImportError:
                print("❌ 未安装win32com库，无法自动保存")
            except Exception as e:
                print(f"❌ 自动保存失败：{e}")
        
        # 如果自动保存失败，提供备选方案
        if not save_success:
            print("\n" + "="*60)
            print("自动保存失败，请使用以下备选方案：")
            print("1. 在Excel中按 Ctrl+S 保存文件")
            print("2. 或者点击Excel的保存按钮")
            print("3. 确保看到文件已保存的提示")
            print("="*60)
            print("保存完成后，请按回车键继续...")
            input("按回车键继续...")
        
        print("正在关闭Excel文件...")
        
        # 关闭Excel进程（Windows系统）
        if os.name == 'nt':
            try:
                subprocess.run(['taskkill', '/f', '/im', 'excel.exe'], 
                             capture_output=True, check=False)
                print("Excel已关闭")
            except Exception as e:
                print(f"关闭Excel时出现警告：{e}")
        else:
            print("请手动关闭Excel文件")
        
        return True
        
    except Exception as e:
        print(f"打开或关闭Excel时出现错误：{e}")
        return False

def main():
    """主函数 - 执行所有数据更新任务"""
    print("🚀 统一数据更新脚本启动")
    print("="*80)
    print("本脚本将依次执行以下任务：")
    print("1. 更新每日监控数据")
    print("2. 更新Shares数据")
    print("3. 生成基金价格数据")
    print("4. 生成股票价格数据")
    print("="*80)
    
    # 记录开始时间
    start_time = time.time()
    
    # 执行所有任务
    results = []
    
    # 任务1: 更新每日监控数据
    results.append(("每日监控数据", update_monitoring_data()))
    
    # 任务2: 更新Shares数据
    results.append(("Shares数据", update_shares_data()))
    
    # 任务3: 生成基金价格数据
    results.append(("基金价格数据", generate_funds_value_data()))
    
    # 任务4: 生成股票价格数据
    results.append(("股票价格数据", generate_stock_price_data()))
    
    # 显示结果汇总
    print("\n" + "="*80)
    print("📊 任务执行结果汇总")
    print("="*80)
    
    success_count = 0
    for task_name, result in results:
        status = "✅ 成功" if result else "❌ 失败"
        print(f"{task_name}: {status}")
        if result:
            success_count += 1
    
    print(f"\n总计: {success_count}/{len(results)} 个任务成功完成")
    
    # 计算总耗时
    total_time = time.time() - start_time
    print(f"总耗时: {total_time:.2f} 秒")
    
    if success_count == len(results):
        print("\n🎉 所有任务都成功完成！")
    else:
        print(f"\n⚠️  有 {len(results) - success_count} 个任务失败，请检查错误信息")
    
    print("="*80)

if __name__ == "__main__":
    main()
