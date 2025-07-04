#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试恭喜通知功能
"""

import pandas as pd
from datetime import datetime

def test_congratulations():
    """测试恭喜通知功能"""
    print("🎉 测试恭喜通知功能")
    print("=" * 50)
    
    # 读取Excel文件
    try:
        df = pd.read_excel("yxc.xlsx")
        print(f"✅ 成功读取Excel文件")
        print(f"📊 表格形状: {df.shape}")
        
        # 显示原始数据
        print("\n📄 原始数据:")
        print(df)
        
        # 模拟所有项目都有剩余天数（没有到期项目）
        print(f"\n🔄 模拟所有项目都有剩余天数...")
        
        for idx, row in df.iterrows():
            # 给每个项目设置一些剩余天数
            remaining_days = (idx + 1) * 5  # 5, 10, 15, 20天
            df.at[idx, '剩余'] = remaining_days
            print(f"  行 {idx + 1}: {row.get(' 店铺名称', '未知')} 剩余天数设置为 {remaining_days}")
        
        # 保存修改后的数据
        df.to_excel("yxc.xlsx", index=False)
        print(f"\n💾 已保存修改后的数据")
        
        print(f"\n📄 修改后的数据:")
        print(df)
        
        print(f"\n🎯 现在可以运行智能监控脚本测试恭喜通知功能:")
        print(f"   python3 smart_monitor.py")
        
        return True
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

if __name__ == "__main__":
    test_congratulations() 