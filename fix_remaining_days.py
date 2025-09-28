#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
修复剩余天数计算 - 根据开始时间和总天数重新计算
"""

import pandas as pd
from datetime import datetime
import shutil

def fix_remaining_days():
    """修复剩余天数计算"""
    print("=== 修复剩余天数计算 ===")
    
    # 备份原文件（如果可能的话）
    try:
        shutil.copy2('yxc.xlsx', 'yxc_backup_before_fix.xlsx')
        print("✅ 已备份原文件: yxc_backup_before_fix.xlsx")
    except Exception as e:
        print(f"⚠️ 备份失败，继续执行: {e}")
    
    # 读取Excel文件
    df = pd.read_excel('yxc.xlsx')
    print(f"✅ 读取Excel文件，共 {len(df)} 行数据")
    
    # 查找关键列
    columns = {
        'remaining': '剩余',
        'total': '总天', 
        'start_date': '开始时间'
    }
    
    current_date = datetime.now()
    print(f"📅 当前日期: {current_date.strftime('%Y%m%d')}")
    
    # 重新计算剩余天数
    fixed_count = 0
    for idx, row in df.iterrows():
        start_date_str = str(row[columns['start_date']])
        total_days = int(row[columns['total']])
        old_remaining = int(row[columns['remaining']])
        
        # 解析开始时间
        try:
            start_date = datetime.strptime(start_date_str, '%Y%m%d')
            
            # 计算已过去的天数
            days_passed = (current_date - start_date).days
            
            # 计算正确的剩余天数
            correct_remaining = max(0, total_days - days_passed)
            
            # 更新剩余天数
            df.at[idx, columns['remaining']] = correct_remaining
            
            store_name = row.get(' 店铺名称', f'行{idx+1}')
            print(f"📊 {store_name}: 开始{start_date_str}, 总{total_days}天, 已过{days_passed}天, {old_remaining}→{correct_remaining}")
            
            if old_remaining != correct_remaining:
                fixed_count += 1
                
        except Exception as e:
            print(f"❌ 处理行{idx+1}失败: {e}")
    
    print(f"\n✅ 修复了 {fixed_count} 个项目的剩余天数")
    
    # 保存修复后的文件
    df.to_excel('yxc.xlsx', index=False)
    print("💾 已保存修复后的Excel文件")
    
    # 显示修复后的状态
    print("\n📊 修复后的前10行数据:")
    print(df[[' 店铺名称', '总天', '剩余', '开始时间']].head(10))
    
    # 统计过期项目
    expired_count = len(df[df[columns['remaining']] == 0])
    print(f"\n🚨 发现 {expired_count} 个过期项目（剩余天数为0）")

if __name__ == "__main__":
    fix_remaining_days()
