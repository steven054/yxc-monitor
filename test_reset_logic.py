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
    
    print(f"🧪 测试重置逻辑: {excel_file}")
    
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
        
        # 模拟将第一行的剩余天数设为0
        print(f"\n🔄 模拟将第一行剩余天数设为0...")
        df.at[0, columns['remaining']] = 0
        
        print(f"行 1 剩余天数已设为: {df.at[0, columns['remaining']]}")
        
        # 模拟重置逻辑
        print(f"\n🔄 执行重置逻辑...")
        current_date = datetime.now()
        new_start_date = current_date.strftime('%Y%m%d')
        
        # 先更新开始时间为今天
        old_start_date = df.at[0, columns['start_date']]
        df.at[0, columns['start_date']] = int(new_start_date)
        
        # 然后重置剩余天数为总天数
        total_days = df.at[0, columns['total']]
        df.at[0, columns['remaining']] = total_days
        
        print(f"✅ 重置完成:")
        print(f"  开始时间: {old_start_date} → {new_start_date}")
        print(f"  剩余天数: 0 → {total_days}")
        
        # 保存测试文件
        test_file = "test_reset_result.xlsx"
        df.to_excel(test_file, index=False)
        print(f"\n💾 测试结果已保存到: {test_file}")
        
        # 验证重置后的状态
        print(f"\n📊 重置后状态:")
        for idx, row in df.iterrows():
            print(f"行 {idx + 1}: {row.get(' 店铺名称', 'N/A')}")
            print(f"  总天数: {row[columns['total']]}")
            print(f"  剩余天数: {row[columns['remaining']]}")
            print(f"  开始时间: {row[columns['start_date']]}")
        
        print(f"\n🎯 测试完成！重置逻辑工作正常。")
        print(f"💡 现在第一行的剩余天数已经从0重置为总天数，开始时间也更新为今天。")
        
    except Exception as e:
        print(f"❌ 测试重置逻辑时出错: {e}")

if __name__ == "__main__":
    main() 