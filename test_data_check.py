#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

def parse_date(date_value):
    """解析日期格式"""
    if pd.isna(date_value):
        return None
    
    date_str = str(date_value)
    
    # 尝试不同的日期格式
    formats = [
        '%Y%m%d',      # 20250403
        '%Y-%m-%d',    # 2025-04-03
        '%Y/%m/%d',    # 2025/04/03
        '%Y年%m月%d日', # 2025年04月03日
    ]
    
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    
    return None

def calculate_remaining_days(start_date, total_days):
    """计算剩余天数"""
    if start_date is None or total_days is None:
        return None
    
    try:
        start = parse_date(start_date)
        if start is None:
            return None
        
        total = int(total_days)
        today = datetime.now()
        
        # 计算已经过去的天数
        days_passed = (today - start).days
        
        # 计算剩余天数
        remaining = total - days_passed
        
        return max(0, remaining)  # 不能为负数
    except Exception as e:
        print(f"❌ 计算剩余天数失败: {e}")
        return None

def main():
    """主函数"""
    excel_file = os.getenv('EXCEL_FILE', 'yxc.xlsx')
    
    print(f"🔍 检查Excel文件: {excel_file}")
    
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file)
        print(f"✅ 成功读取Excel文件，共 {len(df)} 行数据")
        
        # 显示列名
        print(f"\n📋 列名: {list(df.columns)}")
        
        # 查找关键列
        columns = {}
        
        # 查找剩余天数列
        possible_remaining = ['剩余天数', '剩余时间', '到期天数', '过期天数', '天数', 'days', 'remaining_days', '剩余', '到期', '过期']
        for col in df.columns:
            col_str = str(col).lower()
            for name in possible_remaining:
                if name in col_str:
                    columns['remaining'] = col
                    break
            if 'remaining' in columns:
                break
        
        # 查找总天数列
        possible_total = ['总天数', '总时间', '总天', 'total_days', 'total', '天']
        for col in df.columns:
            col_str = str(col).lower()
            for name in possible_total:
                if name in col_str:
                    columns['total'] = col
                    break
            if 'total' in columns:
                break
        
        # 查找开始时间列
        possible_start = ['开始时间', '开始日期', 'start_date', 'start_time', '开始']
        for col in df.columns:
            col_str = str(col).lower()
            for name in possible_start:
                if name in col_str:
                    columns['start_date'] = col
                    break
            if 'start_date' in columns:
                break
        
        print(f"\n🎯 找到的列: {columns}")
        
        if not all(key in columns for key in ['remaining', 'total', 'start_date']):
            print("❌ 未找到所有必需的列")
            return
        
        # 检查前5行数据
        print(f"\n📊 前5行数据详情:")
        for idx in range(min(5, len(df))):
            row = df.iloc[idx]
            print(f"\n行 {idx + 1}:")
            print(f"  店铺名称: {row.get(' 店铺名称', 'N/A')}")
            print(f"  开始日期: {row[columns['start_date']]} (类型: {type(row[columns['start_date']])})")
            print(f"  总天数: {row[columns['total']]} (类型: {type(row[columns['total']])})")
            print(f"  当前剩余天数: {row[columns['remaining']]} (类型: {type(row[columns['remaining']])})")
            
            # 计算应该的剩余天数
            calculated_remaining = calculate_remaining_days(row[columns['start_date']], row[columns['total']])
            print(f"  计算出的剩余天数: {calculated_remaining}")
            
            if calculated_remaining is not None:
                current_remaining = row[columns['remaining']]
                if pd.notna(current_remaining):
                    current_remaining = int(current_remaining)
                    if current_remaining != calculated_remaining:
                        print(f"  ⚠️  不匹配！当前: {current_remaining}, 计算: {calculated_remaining}")
                    else:
                        print(f"  ✅ 匹配")
                else:
                    print(f"  ⚠️  当前剩余天数为空")
            else:
                print(f"  ❌ 无法计算剩余天数")
        
        # 检查所有行的剩余天数计算
        print(f"\n🔢 所有行的剩余天数计算:")
        updated_count = 0
        for idx, row in df.iterrows():
            start_date = row[columns['start_date']]
            total = row[columns['total']]
            current_remaining = row[columns['remaining']]
            
            calculated_remaining = calculate_remaining_days(start_date, total)
            
            if calculated_remaining is not None:
                if pd.notna(current_remaining):
                    current_remaining = int(current_remaining)
                    if current_remaining != calculated_remaining:
                        print(f"行 {idx + 1}: {row.get(' 店铺名称', 'N/A')} - 当前: {current_remaining} → 计算: {calculated_remaining}")
                        updated_count += 1
                else:
                    print(f"行 {idx + 1}: {row.get(' 店铺名称', 'N/A')} - 当前: 空 → 计算: {calculated_remaining}")
                    updated_count += 1
        
        print(f"\n📈 需要更新的行数: {updated_count}")
        
    except Exception as e:
        print(f"❌ 检查Excel文件时出错: {e}")

if __name__ == "__main__":
    main() 