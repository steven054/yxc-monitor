#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

def test_reset_logic():
    """测试重置逻辑"""
    print("=== 测试重置逻辑 ===")
    
        # 读取Excel文件
    df = pd.read_excel(os.getenv('EXCEL_FILE', 'yxc.xlsx'))
    print(f"原始数据:")
    print(df)
    print()
        
    # 模拟列名查找
    columns = {
        'remaining': '剩余',
        'total': '总天', 
        'start_date': '开始时间'
    }
    print(f"使用的列名: {columns}")
    print()
    
    # 模拟update_expired_items函数
    updated_items = []
    current_date = datetime.now()
    
    print("检查每个项目:")
        for idx, row in df.iterrows():
        remaining = row[columns['remaining']]
        total = row[columns['total']]
        start_date = row[columns['start_date']]
        
        print(f"行{idx+1}: 剩余={remaining}, 总天={total}, 开始时间={start_date}")
        
        # 如果剩余天数为0，需要重置
        if pd.notna(remaining) and int(remaining) == 0:
            print(f"  🔄 发现剩余天数为0，准备重置")
            
            # 更新开始时间为今天
        new_start_date = current_date.strftime('%Y%m%d')
            print(f"    新开始时间: {new_start_date}")
            
            # 重置剩余天数为总天数
            print(f"    重置剩余天数: {total}")
            
            # 实际更新数据
            df.at[idx, columns['start_date']] = str(new_start_date)
            df.at[idx, columns['remaining']] = total
            
            updated_items.append({
                'row': idx + 1,
                'name': row.get(' 店铺名称', f'行{idx+1}'),
                'total_days': total,
                'old_start': start_date,
                'new_start': new_start_date
            })
        else:
            print(f"  ✅ 剩余天数不为0，跳过")
        print()
    
    print(f"重置了 {len(updated_items)} 个项目")
    print()
    
    print("更新后的数据:")
    print(df)
    print()
    
    # 保存更新后的文件
        test_file = "test_reset_result.xlsx"
        df.to_excel(test_file, index=False)
    print(f"已保存到 {test_file}")

if __name__ == "__main__":
    test_reset_logic() 