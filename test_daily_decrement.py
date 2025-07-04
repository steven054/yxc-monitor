#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

def main():
    """主函数"""
    excel_file = os.getenv('EXCEL_FILE', 'yxc.xlsx')
    
    print(f"🧪 测试每天减1逻辑: {excel_file}")
    
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file)
        print(f"✅ 成功读取Excel文件，共 {len(df)} 行数据")
        
        # 查找关键列
        columns = {}
        for col in df.columns:
            if '剩余' in str(col):
                columns['remaining'] = col
            elif '总天' in str(col):
                columns['total'] = col
            elif '开始时间' in str(col):
                columns['start_date'] = col
        
        print(f"🎯 找到的列: {columns}")
        
        # 显示当前状态
        print(f"\n📊 当前状态:")
        for idx, row in df.iterrows():
            print(f"行 {idx + 1}: {row.get(' 店铺名称', 'N/A')}")
            print(f"  总天数: {row[columns['total']]}")
            print(f"  剩余天数: {row[columns['remaining']]}")
            print(f"  开始时间: {row[columns['start_date']]}")
        
        # 模拟每天减1的逻辑
        print(f"\n🔄 模拟每天减1...")
        updated_count = 0
        
        for idx, row in df.iterrows():
            current_remaining = row[columns['remaining']]
            
            # 如果当前剩余天数为0，跳过
            if pd.notna(current_remaining) and int(current_remaining) == 0:
                print(f"⏭️  跳过减1: {row.get(' 店铺名称', f'行{idx+1}')} 已经是0天")
                continue
            
            # 如果剩余天数不为0，每天减1
            if pd.notna(current_remaining) and int(current_remaining) > 0:
                old_remaining = int(current_remaining)
                new_remaining = max(0, old_remaining - 1)  # 不能小于0
                
                df.at[idx, columns['remaining']] = new_remaining
                updated_count += 1
                print(f"📅 剩余天数减1: {row.get(' 店铺名称', f'行{idx+1}')} {old_remaining} → {new_remaining}")
        
        print(f"\n📈 总共更新了 {updated_count} 个项目的剩余天数")
        
        # 检查是否有项目变成0天
        print(f"\n🔍 检查是否有项目变成0天...")
        zero_count = 0
        for idx, row in df.iterrows():
            remaining = row[columns['remaining']]
            if pd.notna(remaining) and int(remaining) == 0:
                zero_count += 1
                print(f"🚨 发现0天项目: {row.get(' 店铺名称', f'行{idx+1}')}")
        
        if zero_count > 0:
            print(f"\n🔄 模拟重置0天项目...")
            current_date = datetime.now()
            new_start_date = current_date.strftime('%Y%m%d')
            
            for idx, row in df.iterrows():
                remaining = row[columns['remaining']]
                if pd.notna(remaining) and int(remaining) == 0:
                    total_days = row[columns['total']]
                    old_start_date = row[columns['start_date']]
                    
                    # 先更新开始时间为今天
                    df.at[idx, columns['start_date']] = int(new_start_date)
                    
                    # 然后重置剩余天数为总天数
                    df.at[idx, columns['remaining']] = total_days
                    
                    print(f"✅ 重置项目: {row.get(' 店铺名称', f'行{idx+1}')}")
                    print(f"  开始时间: {old_start_date} → {new_start_date}")
                    print(f"  剩余天数: 0 → {total_days}")
        
        # 保存测试文件
        test_file = "test_daily_decrement_result.xlsx"
        df.to_excel(test_file, index=False)
        print(f"\n💾 测试结果已保存到: {test_file}")
        
        # 显示最终状态
        print(f"\n📊 最终状态:")
        for idx, row in df.iterrows():
            print(f"行 {idx + 1}: {row.get(' 店铺名称', 'N/A')}")
            print(f"  总天数: {row[columns['total']]}")
            print(f"  剩余天数: {row[columns['remaining']]}")
            print(f"  开始时间: {row[columns['start_date']]}")
        
        print(f"\n🎯 测试完成！")
        print(f"💡 现在剩余天数不为0的项目每天都会减1，为0的项目会被重置。")
        
    except Exception as e:
        print(f"❌ 测试每天减1逻辑时出错: {e}")

if __name__ == "__main__":
    main() 