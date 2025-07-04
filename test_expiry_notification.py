#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试到期通知功能
"""

import pandas as pd
from datetime import datetime

def test_expiry_notification():
    """测试到期通知功能"""
    print("🚨 测试到期通知功能")
    print("=" * 50)
    
    # 读取Excel文件
    try:
        df = pd.read_excel("yxc.xlsx")
        print(f"✅ 成功读取Excel文件")
        print(f"📊 表格形状: {df.shape}")
        
        # 显示原始数据
        print("\n📄 原始数据:")
        print(df)
        
        # 模拟有项目到期（剩余天数为0）
        print(f"\n🔄 模拟有项目到期...")
        
        # 让前两个项目到期
        df.at[0, '剩余'] = 0  # 清江饭店到期
        df.at[1, '剩余'] = 0  # 积分饭店到期
        
        print(f"  行 1: 清江饭店 剩余天数设置为 0 (到期)")
        print(f"  行 2: 积分饭店 剩余天数设置为 0 (到期)")
        print(f"  行 3: 份额文件发我饭店 保持 15 天")
        print(f"  行 4: 三只羊饭店 保持 20 天")
        
        # 保存修改后的数据
        df.to_excel("yxc.xlsx", index=False)
        print(f"\n💾 已保存修改后的数据")
        
        print(f"\n📄 修改后的数据:")
        print(df)
        
        print(f"\n🎯 现在可以运行智能监控脚本测试到期通知功能:")
        print(f"   python3 smart_monitor.py")
        
        return True
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

if __name__ == "__main__":
    test_expiry_notification() 