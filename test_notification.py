#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试通知功能脚本
"""

import pandas as pd
import os
from datetime import datetime
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

def test_excel_reading():
    """测试Excel文件读取"""
    print("=== 测试Excel文件读取 ===")
    
    try:
        df = pd.read_excel("yxc.xlsx")
        print(f"✅ 成功读取Excel文件")
        print(f"📊 表格形状: {df.shape}")
        print(f"📋 列名: {list(df.columns)}")
        
        # 显示前几行数据
        print("\n📄 表格内容预览:")
        print(df.head())
        
        return df
    except Exception as e:
        print(f"❌ 读取Excel文件失败: {e}")
        return None

def test_expiry_detection(df):
    """测试到期检测"""
    print("\n=== 测试到期检测 ===")
    
    if df is None:
        return
    
    # 查找剩余天数列
    possible_names = ['剩余天数', '剩余时间', '到期天数', '过期天数', '天数', 'days', 'remaining_days', '剩余', '到期', '过期']
    
    expiry_column = None
    for col in df.columns:
        col_str = str(col).lower()
        for name in possible_names:
            if name in col_str:
                expiry_column = col
                break
        if expiry_column:
            break
    
    if not expiry_column:
        print("⚠️  未找到明确的剩余天数列，使用第一列")
        expiry_column = df.columns[0]
    
    print(f"🎯 找到剩余天数列: {expiry_column}")
    
    # 检查剩余天数
    try:
        df[expiry_column] = pd.to_numeric(df[expiry_column], errors='coerce')
        
        # 查找剩余天数为0的项目
        expired_mask = df[expiry_column] == 0
        expired_df = df[expired_mask]
        
        if not expired_df.empty:
            print(f"🚨 发现 {len(expired_df)} 个剩余天数为0的项目:")
            for idx, row in expired_df.iterrows():
                print(f"   行 {idx + 1}: {dict(row)}")
        else:
            print("✅ 没有发现剩余天数为0的项目")
            
        # 显示所有剩余天数
        print(f"\n📊 所有项目的剩余天数:")
        for idx, row in df.iterrows():
            days = row[expiry_column]
            print(f"   行 {idx + 1}: {days} 天")
            
    except Exception as e:
        print(f"❌ 检查剩余天数时出错: {e}")

def test_configuration():
    """测试配置"""
    print("\n=== 测试配置 ===")
    
    configs = {
        'email': {
            'enabled': os.getenv('EMAIL_ENABLED', 'false').lower() == 'true',
            'username': os.getenv('EMAIL_USERNAME'),
            'to_email': os.getenv('TO_EMAIL')
        },
        'wechat': {
            'enabled': os.getenv('WECHAT_ENABLED', 'false').lower() == 'true',
            'webhook_url': os.getenv('WECHAT_WEBHOOK_URL')
        },
        'sms': {
            'enabled': os.getenv('SMS_ENABLED', 'false').lower() == 'true',
            'phone_number': os.getenv('SMS_PHONE_NUMBER')
        }
    }
    
    for method, config in configs.items():
        status = "✅ 已启用" if config['enabled'] else "❌ 未启用"
        print(f"{method.upper()} 通知: {status}")
        
        if config['enabled']:
            for key, value in config.items():
                if key != 'enabled':
                    if value:
                        print(f"   {key}: {value}")
                    else:
                        print(f"   {key}: ❌ 未配置")

def main():
    """主函数"""
    print("🧪 Excel监控脚本测试工具")
    print("=" * 50)
    
    # 测试Excel读取
    df = test_excel_reading()
    
    # 测试到期检测
    test_expiry_detection(df)
    
    # 测试配置
    test_configuration()
    
    print("\n" + "=" * 50)
    print("📝 使用说明:")
    print("1. 编辑 .env 文件配置通知参数")
    print("2. 运行 python3 check_expiry.py 启动监控")
    print("3. 或运行 ./start_monitor.sh 使用启动脚本")

if __name__ == "__main__":
    main() 