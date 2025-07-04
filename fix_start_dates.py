#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

def main():
    """主函数"""
    excel_file = os.getenv('EXCEL_FILE', 'yxc.xlsx')
    
    print(f"🔧 修正Excel文件开始日期: {excel_file}")
    
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file)
        print(f"✅ 成功读取Excel文件，共 {len(df)} 行数据")
        
        # 查找开始时间列
        start_date_col = None
        for col in df.columns:
            if '开始时间' in str(col):
                start_date_col = col
                break
        
        if start_date_col is None:
            print("❌ 未找到开始时间列")
            return
        
        print(f"📅 找到开始时间列: {start_date_col}")
        
        # 显示当前开始日期
        print(f"\n📊 当前开始日期:")
        for idx, row in df.iterrows():
            current_date = row[start_date_col]
            print(f"行 {idx + 1}: {row.get(' 店铺名称', 'N/A')} - {current_date}")
        
        # 计算新的开始日期（从今天往前推，让剩余天数合理）
        today = datetime.now()
        
        print(f"\n🔄 修正开始日期:")
        for idx, row in df.iterrows():
            total_days = row['总天']
            current_remaining = row['剩余']
            
            # 计算新的开始日期：今天减去已经过去的天数
            # 假设剩余天数应该是当前剩余天数，那么已经过去的天数 = 总天数 - 当前剩余天数
            days_passed = total_days - current_remaining
            
            # 新的开始日期 = 今天 - 已经过去的天数
            new_start_date = today - timedelta(days=days_passed)
            
            # 格式化为YYYYMMDD
            new_start_date_str = new_start_date.strftime('%Y%m%d')
            
            print(f"行 {idx + 1}: {row.get(' 店铺名称', 'N/A')}")
            print(f"  总天数: {total_days}, 当前剩余: {current_remaining}")
            print(f"  已过去天数: {days_passed}")
            print(f"  原开始日期: {row[start_date_col]} → 新开始日期: {new_start_date_str}")
            
            # 更新开始日期
            df.at[idx, start_date_col] = int(new_start_date_str)
        
        # 保存文件
        backup_file = f"{excel_file.replace('.xlsx', '')}_backup_before_fix.xlsx"
        df.to_excel(backup_file, index=False)
        print(f"\n💾 已备份原文件为: {backup_file}")
        
        df.to_excel(excel_file, index=False)
        print(f"✅ 已更新Excel文件: {excel_file}")
        
        print(f"\n🎯 修正完成！现在剩余天数应该会每天自动减少。")
        print(f"💡 建议运行测试脚本验证: python3 test_data_check.py")
        
    except Exception as e:
        print(f"❌ 修正Excel文件时出错: {e}")

if __name__ == "__main__":
    main() 