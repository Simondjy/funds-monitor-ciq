#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日期处理工具模块
提供统一的交易日计算、时区处理、市场状态判断等功能
被多个脚本共同使用，确保日期处理逻辑的一致性
"""

from datetime import datetime, timedelta
import pytz

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

def get_next_trading_day(date):
    """获取指定日期的下一个交易日"""
    current = date
    while True:
        current = current + timedelta(days=1)
        if is_trading_day(current):
            return current

def get_week_first_trading_day(date):
    """获取当周第一个交易日"""
    monday = date - timedelta(days=date.weekday())
    return monday if is_trading_day(monday) else get_next_trading_day(monday)

def get_month_first_trading_day(date):
    """获取本月第一个交易日"""
    first_day = date.replace(day=1)
    return first_day if is_trading_day(first_day) else get_next_trading_day(first_day)

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

def is_us_market_open():
    """
    判断美国市场是否开盘
    美国市场时间：美东时间9:30-16:00 (周一至周五)
    """
    us_eastern_tz = pytz.timezone('America/New_York')
    us_eastern_time = datetime.now(us_eastern_tz)
    
    # 检查是否为工作日
    if us_eastern_time.weekday() >= 5:  # 周六、周日
        return False
    
    # 检查时间是否在交易时间内
    current_time = us_eastern_time.time()
    market_open = datetime.strptime('09:30', '%H:%M').time()
    market_close = datetime.strptime('16:00', '%H:%M').time()
    
    return market_open <= current_time <= market_close

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

def calculate_dates(target_date=None):
    """
    计算四个关键日期：n, n-1, 当周第一个交易日, 本月第一个交易日
    主要用于update_monitoring.py
    """
    if target_date is None:
        # 使用美国市场对应的日期
        target_date = get_us_market_date()
    elif isinstance(target_date, str):
        target_date = datetime.strptime(target_date, '%Y-%m-%d').date()
    
    # 检查今天是否已经收市
    market_closed_today = is_market_closed_for_today()
    print(f"今天是否已收市: {'是' if market_closed_today else '否'}")
    
    # 关键逻辑：如果今天已经收市，可以用今天；如果还没收市，必须用前一天
    if market_closed_today:
        # 今天已经收市，CIQ数据已经有了，可以用今天
        n = target_date
        print(f"今天已收市，使用当天日期: {n}")
    else:
        # 今天还没收市，CIQ数据还没出来，必须用前一天
        n = get_previous_trading_day(target_date)
        print(f"今天还没收市，使用前一个交易日: {n}")
    
    # 确保n是交易日，如果不是则调整为前一个交易日
    if not is_trading_day(n):
        print(f"注意：{n} 不是交易日，调整为前一个交易日")
        n = get_previous_trading_day(n)
    
    n_minus_1 = get_previous_trading_day(n)
    week_first = get_week_first_trading_day(n)
    month_first = get_month_first_trading_day(n)
    
    return {
        'n': n.strftime('%Y/%m/%d'),
        'n-1': n_minus_1.strftime('%Y/%m/%d'),
        '当周第一个交易日': week_first.strftime('%Y/%m/%d'),
        '本月第一个交易日': month_first.strftime('%Y/%m/%d')
    }

def calculate_n_and_n_minus_1():
    """
    计算n和n-1日期
    主要用于update_shares_sheet.py
    """
    target_date = get_us_market_date()
    
    # 检查今天是否已经收市
    market_closed_today = is_market_closed_for_today()
    print(f"今天是否已收市: {'是' if market_closed_today else '否'}")
    
    # 关键逻辑：如果今天已经收市，可以用今天；如果还没收市，必须用前一天
    if market_closed_today:
        # 今天已经收市，CIQ数据已经有了，可以用今天
        n = target_date
        print(f"今天已收市，使用当天日期: {n}")
    else:
        # 今天还没收市，CIQ数据还没出来，必须用前一天
        n = get_previous_trading_day(target_date)
        print(f"今天还没收市，使用前一个交易日: {n}")
    
    # 确保n是交易日，如果不是则调整为前一个交易日
    if not is_trading_day(n):
        print(f"注意：{n} 不是交易日，调整为前一个交易日")
        n = get_previous_trading_day(n)
    
    n_minus_1 = get_previous_trading_day(n)
    
    return {
        'n': n.strftime('%Y-%m-%d'),
        'n-1': n_minus_1.strftime('%Y-%m-%d')
    }

def get_latest_trading_day():
    """
    获取今天或今天前最近一个交易日，考虑美国市场收盘时间
    主要用于updatefundsvalue.py和updatestockprice.py
    """
    # 使用美东时间作为基准
    target_date = get_us_market_date()
    
    # 检查今天是否已经收市
    market_closed_today = is_market_closed_for_today()
    print(f"今天是否已收市: {'是' if market_closed_today else '否'}")
    
    # 关键逻辑：如果今天已经收市，可以用今天；如果还没收市，必须用前一天
    if market_closed_today:
        # 今天已经收市，CIQ数据已经有了，可以用今天
        n = target_date
        print(f"今天已收市，使用当天日期: {n}")
    else:
        # 今天还没收市，CIQ数据还没出来，必须用前一天
        n = get_previous_trading_day(target_date)
        print(f"今天还没收市，使用前一个交易日: {n}")
    
    # 确保n是交易日，如果不是则调整为前一个交易日
    if not is_trading_day(n):
        print(f"注意：{n} 不是交易日，调整为前一个交易日")
        n = get_previous_trading_day(n)
    
    return n

def generate_trading_dates(start_date, end_date):
    """生成指定日期范围内的所有交易日"""
    trading_dates = []
    current_date = start_date
    
    while current_date <= end_date:
        if is_trading_day(current_date):
            trading_dates.append(current_date)
        current_date += timedelta(days=1)
    
    return trading_dates
