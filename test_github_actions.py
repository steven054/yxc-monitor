#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import requests
import json
from dotenv import load_dotenv

def test_webhook_connection():
    """测试webhook连接"""
    load_dotenv()
    
    webhook_url = os.getenv('WECHAT_WEBHOOK_URL', '')
    
    if not webhook_url:
        print("❌ 未找到WECHAT_WEBHOOK_URL环境变量")
        return False
    
    print(f"🔗 Webhook地址: {webhook_url}")
    
    # 发送测试消息
    test_message = {
        "msgtype": "text",
        "text": {
            "content": "🧪 GitHub Actions配置测试消息\n时间: " + str(os.popen('date').read().strip())
        }
    }
    
    try:
        response = requests.post(webhook_url, json=test_message, timeout=10)
        
        if response.status_code == 200:
            result = response.json()
            if result.get('errcode') == 0:
                print("✅ Webhook连接测试成功！")
                print("📱 测试消息已发送到企业微信群")
                return True
            else:
                print(f"❌ Webhook返回错误: {result}")
                return False
        else:
            print(f"❌ HTTP请求失败: {response.status_code}")
            print(f"响应内容: {response.text}")
            return False
            
    except Exception as e:
        print(f"❌ 连接测试失败: {str(e)}")
        return False

def test_environment():
    """测试环境配置"""
    print("🔍 检查环境配置...")
    
    # 检查必要文件
    required_files = ['yxc.xlsx', 'wechat_with_image_fix.py', 'requirements.txt']
    for file in required_files:
        if os.path.exists(file):
            print(f"✅ {file} 存在")
        else:
            print(f"❌ {file} 不存在")
    
    # 检查环境变量
    webhook_url = os.getenv('WECHAT_WEBHOOK_URL', '')
    if webhook_url:
        print("✅ WECHAT_WEBHOOK_URL 已配置")
    else:
        print("❌ WECHAT_WEBHOOK_URL 未配置")

if __name__ == "__main__":
    print("🚀 开始GitHub Actions配置测试...\n")
    
    test_environment()
    print()
    
    if test_webhook_connection():
        print("\n🎉 所有测试通过！GitHub Actions配置应该可以正常工作。")
    else:
        print("\n⚠️  测试失败，请检查配置。")
