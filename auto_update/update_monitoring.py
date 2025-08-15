#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
专门用于更新每日监控数据的脚本
修改 每日数据监控-0807.xlsx 文件中的特定单元格并另存为新文件
自动计算并更新日期：n, n-1, 当周第一个交易日, 本月第一个交易日
"""

import win32com.client
import time
from pathlib import Path
from datetime import datetime, timedelta
import pytz

# 导入公共日期处理模块
from date_utils import calculate_dates

def update_monitoring_data():
    """
    更新每日监控数据：
    - 源文件：每日数据监控.xlsx
    - 工作表：raw1
    - 自动计算并更新：A13=n, A14=n-1, A15=当周第一个交易日, A16=本月第一个交易日
    - 目标文件：每日监控数据_complete.xlsx
    """
    
    # 文件路径
    source_file = Path("process_data/每日数据监控.xlsx").resolve()
    target_file = Path("process_data/每日监控数据_complete.xlsx").resolve()
    
    #print("=== 开始更新每日监控数据 ===")
    #print(f"源文件: {source_file}")
    #print(f"目标文件: {target_file}")
    
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
        #print("正在启动Excel应用程序...")
        
        # 创建Excel应用程序实例
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False  # 隐藏Excel窗口
        excel_app.DisplayAlerts = False  # 禁用警告对话框
        excel_app.ScreenUpdating = False  # 禁用屏幕更新
        excel_app.EnableEvents = False  # 禁用事件
        excel_app.Interactive = False  # 禁用交互
        
        #print("正在打开Excel文件...")
        
        # 打开工作簿
        workbook = excel_app.Workbooks.Open(str(source_file))
        #print(f"✅ 成功打开Excel文件: {source_file}")
        
        # 获取raw1工作表
        #print("正在获取raw1工作表...")
        worksheet = workbook.Worksheets("raw1")
        
        # 更新A13单元格 - n
        #print("正在更新A13单元格...")
        worksheet.Range("A13").Value = dates['n']
        #print(f"✅ 已更新A13单元格为: {dates['n']}")
        
        # 更新A14单元格 - n-1
        #print("正在更新A14单元格...")
        worksheet.Range("A14").Value = dates['n-1']
        #print(f"✅ 已更新A14单元格为: {dates['n-1']}")
        
        # 更新A15单元格 - 当周第一个交易日
        #print("正在更新A15单元格...")
        worksheet.Range("A15").Value = dates['当周第一个交易日']
        #print(f"✅ 已更新A15单元格为: {dates['当周第一个交易日']}")
        
        # 更新A16单元格 - 本月第一个交易日
        #print("正在更新A16单元格...")
        worksheet.Range("A16").Value = dates['本月第一个交易日']
        #print(f"✅ 已更新A16单元格为: {dates['本月第一个交易日']}")
        
        # 刷新所有数据连接
        try:
            #print("正在刷新数据连接...")
            for i, connection in enumerate(workbook.Connections, 1):
                print(f"  刷新连接 {i}/{len(workbook.Connections)}: {connection.Name}")
                connection.Refresh()
            print("✅ 已刷新所有数据连接")
        except Exception as conn_error:
            print(f"⚠️  数据连接刷新跳过: {conn_error}")
        
        # 等待数据连接完成
        print("等待数据连接完成...")
        time.sleep(5)  # 增加等待时间到5秒
        
        # 另存为新文件
        #print(f"正在保存为新文件: {target_file}")
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

if __name__ == "__main__":
    #print("每日监控数据更新工具")
    #print("=" * 50)
    
    # 运行更新
    update_monitoring_data()
    
    #print("\n🎉 更新完成！")
    #print("文件已保存为: process_data/每日监控数据_complete.xlsx")
