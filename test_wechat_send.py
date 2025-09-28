#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
import json
import os
from dotenv import load_dotenv
from datetime import datetime

def test_wechat_webhook():
    """测试微信群webhook配置"""
    load_dotenv()
    
    webhook_url = os.getenv('WECHAT_WEBHOOK_URL', '')
    
    if not webhook_url:
        print("❌ 未配置WECHAT_WEBHOOK_URL")
        print("📝 请在.env文件中添加：")
        print("WECHAT_WEBHOOK_URL=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=YOUR_KEY")
        return False
    
    print(f"✅ 找到webhook配置: {webhook_url[:50]}...")
    
    # 发送测试消息
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    test_message = f"""🧪 微信群机器人测试消息

⏰ 发送时间: {current_time}
📊 测试内容: 店铺监控系统
✅ 如果看到此消息，说明配置成功！

接下来可以发送完整的监控数据表格。"""
    
    try:
        data = {
            "msgtype": "text",
            "text": {
                "content": test_message
            }
        }
        
        print("📤 发送测试消息...")
        response = requests.post(webhook_url, json=data)
        result = response.json()
        
        if result.get('errcode') == 0:
            print("✅ 测试消息发送成功！")
            print("📱 请检查微信群，应该能看到测试消息")
            return True
        else:
            print(f"❌ 测试消息发送失败: {result}")
            return False
            
    except Exception as e:
        print(f"❌ 发送测试消息时出错: {e}")
        return False

if __name__ == "__main__":
    print("=== 微信群Webhook测试 ===")
    success = test_wechat_webhook()
    
    if success:
        print("\n🎉 测试成功！现在可以发送完整的监控数据了")
        print("💡 运行以下命令发送完整报告：")
        print("python3 wechat_group_sender.py")
    else:
        print("\n❌ 测试失败，请检查配置")
        print("📖 详细配置说明请查看: WECHAT_SETUP.md")
