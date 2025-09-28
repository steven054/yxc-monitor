#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
重置过期项目 - 将剩余天数为0的项目重置为总天数，并更新开始时间
"""

import pandas as pd
from datetime import datetime
import shutil

def reset_expired_items():
    """重置过期项目"""
    print("=== 重置过期项目 ===")
    
    # 备份原文件
    try:
        shutil.copy2('yxc.xlsx', 'yxc_backup_before_reset.xlsx')
        print("✅ 已备份原文件: yxc_backup_before_reset.xlsx")
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
    new_start_date = current_date.strftime('%Y%m%d')
    print(f"📅 当前日期: {current_date.strftime('%Y%m%d')}")
    print(f"🔄 新的开始时间: {new_start_date}")
    
    # 重置过期项目
    reset_count = 0
    for idx, row in df.iterrows():
        remaining = row[columns['remaining']]
        total = row[columns['total']]
        start_date = row[columns['start_date']]
        
        # 如果剩余天数为0，需要重置
        if pd.notna(remaining) and int(remaining) == 0:
            store_name = row.get(' 店铺名称', f'行{idx+1}')
            print(f"🔄 重置项目: {store_name}")
            print(f"   开始时间: {start_date} → {new_start_date}")
            print(f"   剩余天数: 0 → {total}")
            
            # 更新开始时间为今天
            if pd.api.types.is_integer_dtype(df[columns['start_date']]):
                df.at[idx, columns['start_date']] = int(new_start_date)
            else:
                df.at[idx, columns['start_date']] = str(new_start_date)
            
            # 重置剩余天数为总天数
            df.at[idx, columns['remaining']] = total
            
            reset_count += 1
    
    print(f"\n✅ 重置了 {reset_count} 个过期项目")
    
    # 保存重置后的文件
    df.to_excel('yxc.xlsx', index=False)
    print("💾 已保存重置后的Excel文件")
    
    # 显示重置后的状态
    print("\n📊 重置后的前10行数据:")
    print(df[[' 店铺名称', '总天', '剩余', '开始时间']].head(10))
    
    # 统计过期项目
    expired_count = len(df[df[columns['remaining']] == 0])
    print(f"\n🚨 重置后还有 {expired_count} 个过期项目（剩余天数为0）")

if __name__ == "__main__":
    reset_expired_items()

