#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基金价格数据生成脚本
生成从2024/12/20到今天或今天前最近一个交易日的基金收盘价数据
注意：考虑北京时间与美国时间的时差，美国市场收盘后才能读取当天数据
"""

import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
import pytz
import time
import subprocess
import os

# 导入公共日期处理模块
from date_utils import get_latest_trading_day, generate_trading_dates

# 基金代码列表
tickers = [
    "NasdaqGM:QQQ", "NasdaqGM:AGIX", "NasdaqGM:SMH", "BATS:IGV", 
    "NasdaqGM:BOTZ", "NasdaqGM:AIQ", "ARCA:ARTY", "NasdaqGM:ROBT", "ARCA:IGPT", 
    "BATS:WTAI", "ARCA:THNQ", "NasdaqGM:FDTX", "ARCA:CHAT", "ARCA:LOUP", 
    "ARCA:LRNZ", "ARCA:AIS", "NasdaqGM:WISE", "LSE:RBOT", "XTRA:XAIX", 
    "BIT:WTAI", "LSE:AIAG", "ASX:RBTZ", "DB:XB0T"
]

def is_trading_day(date):
    """判断是否为交易日（周一至周五）"""
    return date.weekday() < 5

def is_us_market_closed():
    """判断美国市场是否已收盘"""
    # 获取北京时间
    beijing_tz = pytz.timezone('Asia/Shanghai')
    beijing_time = datetime.now(beijing_tz)
    
    # 获取美国东部时间（纽约时间）
    ny_tz = pytz.timezone('America/New_York')
    ny_time = beijing_time.astimezone(ny_tz)
    
    # 美国市场交易时间：9:30 AM - 4:00 PM ET（周一至周五）
    # 如果当前时间在美国市场收盘时间之后，则认为市场已收盘
    if ny_time.weekday() >= 5:  # 周末
        return True
    
    # 判断是否在交易时间内
    market_open = ny_time.replace(hour=9, minute=30, second=0, microsecond=0)
    market_close = ny_time.replace(hour=16, minute=0, second=0, microsecond=0)
    
    # 如果当前时间在收盘时间之后，市场已收盘
    if ny_time >= market_close:
        return True
    
    return False

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

def get_previous_trading_day(start_date):
    """获取指定日期前的最近一个交易日"""
    current_date = start_date - timedelta(days=1)
    while not is_trading_day(current_date):
        current_date -= timedelta(days=1)
    return current_date

def wait_for_ciq_refresh(excel_file_path, wait_time=40):
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

def generate_funds_value_data():
    """生成基金价格数据"""
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
        for ticker in tickers:
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
    else:
        print("Excel文件处理过程中出现错误，请手动检查。")

if __name__ == "__main__":
    generate_funds_value_data()
