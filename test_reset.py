#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试重置功能
"""

import pandas as pd
from datetime import datetime

def test_reset_function():
    """测试重置功能"""
    print("🧪 测试重置功能")
    print("=" * 50)
    
    # 读取Excel文件
    try:
        df = pd.read_excel("yxc.xlsx")
        print(f"✅ 成功读取Excel文件")
        print(f"📊 表格形状: {df.shape}")
        print(f"📋 列名: {list(df.columns)}")
        
        # 显示原始数据
        print("\n📄 原始数据:")
        print(df)
        
        # 查找关键列
        columns = {
            'remaining': '剩余',
            'total': '总天', 
            'start_date': '开始时间'
        }
        
        # 检查到期项目
        expired_count = 0
        current_date = datetime.now()
        new_start_date = current_date.strftime('%Y%m%d')
        
        print(f"\n🔄 开始重置到期项目...")
        print(f"📅 当前日期: {new_start_date}")
        
        for idx, row in df.iterrows():
            remaining = row[columns['remaining']]
            total = row[columns['total']]
            start_date = row[columns['start_date']]
            
            print(f"\n检查行 {idx + 1}: {row.get(' 店铺名称', '未知')}")
            print(f"  剩余天数: {remaining}")
            print(f"  总天数: {total}")
            print(f"  开始时间: {start_date}")
            
            # 如果剩余天数为0，需要重置
            if pd.notna(remaining) and int(remaining) == 0:
                print(f"  🔄 发现到期项目，开始重置...")
                
                # 更新开始时间为今天
                df.at[idx, columns['start_date']] = new_start_date
                
                # 重置剩余天数为总天数
                df.at[idx, columns['remaining']] = total
                
                expired_count += 1
                print(f"  ✅ 重置完成: 开始时间 {start_date} → {new_start_date}, 剩余天数 0 → {total}")
            else:
                print(f"  ✅ 未到期，无需重置")
        
        print(f"\n📊 重置统计:")
        print(f"  总共重置了 {expired_count} 个项目")
        
        # 显示更新后的数据
        print(f"\n📄 更新后的数据:")
        print(df)
        
        # 保存文件
        df.to_excel("yxc.xlsx", index=False)
        print(f"\n💾 文件已保存")
        
        return True
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

if __name__ == "__main__":
    test_reset_function() 