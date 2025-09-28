#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
from datetime import datetime
import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

def add_row_numbers_and_notes():
    """为Excel文件添加行号并填充备注内容"""
    excel_file = os.getenv('EXCEL_FILE', 'yxc.xlsx')
    
    print(f"🔧 为Excel文件添加行号和备注内容: {excel_file}")
    
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_file)
        print(f"✅ 成功读取Excel文件，共 {len(df)} 行数据")
        
        # 显示当前列名
        print(f"📋 当前列名: {list(df.columns)}")
        
        # 检查是否已有行号列，如果没有则添加
        if '行号' not in df.columns:
            df.insert(0, '行号', range(1, len(df) + 1))
            print(f"✅ 已添加行号列")
        else:
            print(f"✅ 行号列已存在")
        
        # 填充备注1列的内容（如果为空）
        if '备注1' in df.columns:
            for idx, row in df.iterrows():
                if pd.isna(row['备注1']) or str(row['备注1']).strip() == '':
                    df.at[idx, '备注1'] = '无'
            print(f"✅ 已填充备注1列内容")
        
        # 填充备注2列的内容（如果为空）
        if '备注2' in df.columns:
            for idx, row in df.iterrows():
                if pd.isna(row['备注2']) or str(row['备注2']).strip() == '':
                    df.at[idx, '备注2'] = '无'
            print(f"✅ 已填充备注2列内容")
        
        # 显示更新后的数据
        print(f"\n📊 更新后的前5行数据:")
        print(df.head().to_string())
        
        # 备份原文件
        backup_file = f"yxc_backup_before_add_notes_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df_backup = pd.read_excel(excel_file)
        df_backup.to_excel(backup_file, index=False)
        print(f"💾 已创建备份文件: {backup_file}")
        
        # 保存更新后的文件
        df.to_excel(excel_file, index=False)
        print(f"✅ Excel文件已更新并保存: {excel_file}")
        
        # 显示新的列名
        print(f"📋 更新后的列名: {list(df.columns)}")
        
        return True
        
    except Exception as e:
        print(f"❌ 处理失败: {e}")
        return False

def main():
    """主函数"""
    print("=== 添加行号和备注内容工具 ===")
    
    success = add_row_numbers_and_notes()
    
    if success:
        print("\n🎉 任务完成！")
        print("✅ 已添加行号列")
        print("✅ 已填充备注1和备注2列内容")
        print("💾 已创建备份文件")
    else:
        print("\n❌ 任务失败，请检查错误信息")

if __name__ == "__main__":
    main()
