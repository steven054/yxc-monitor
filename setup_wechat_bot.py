#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import requests
import json
from dotenv import load_dotenv

def setup_wechat_bot():
    """设置企业微信机器人"""
    print("=== 企业微信机器人设置指南 ===")
    print()
    print("📱 步骤1：在微信群中添加企业微信机器人")
    print("1. 打开您要发送消息的微信群")
    print("2. 点击群聊右上角的'...'菜单")
    print("3. 选择'群机器人'或'群助手'")
    print("4. 选择'企业微信机器人'")
    print("5. 点击'添加'")
    print()
    print("🔑 步骤2：获取Webhook地址")
    print("1. 机器人添加成功后，会显示一个webhook地址")
    print("2. 地址格式类似：https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxxxxxxx")
    print("3. 复制这个地址")
    print()
    print("⚙️  步骤3：配置环境变量")
    print("在.env文件中添加以下内容：")
    print("WECHAT_WEBHOOK_URL=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=YOUR_KEY_HERE")
    print()
    
    # 检查当前配置
    load_dotenv()
    webhook_url = os.getenv('WECHAT_WEBHOOK_URL', '')
    
    if webhook_url:
        print(f"✅ 当前已配置webhook地址: {webhook_url[:50]}...")
        print("🧪 测试配置...")
        
        # 测试webhook
        test_message = "🧪 企业微信机器人测试消息\n\n如果您看到此消息，说明配置成功！"
        
        try:
            data = {
                "msgtype": "text",
                "text": {
                    "content": test_message
                }
            }
            
            response = requests.post(webhook_url, json=data)
            result = response.json()
            
            if result.get('errcode') == 0:
                print("✅ 测试消息发送成功！")
                print("📱 请检查微信群，应该能看到测试消息")
                print("🎉 配置完成，现在可以自动发送监控数据了")
                return True
            else:
                print(f"❌ 测试失败: {result}")
                print("💡 可能的原因：")
                print("- webhook地址不正确")
                print("- 机器人已被删除或禁用")
                print("- 网络连接问题")
                return False
                
        except Exception as e:
            print(f"❌ 测试时出错: {e}")
            return False
    else:
        print("❌ 未配置webhook地址")
        print("请按照上述步骤配置后重新运行此脚本")
        return False

def main():
    """主函数"""
    success = setup_wechat_bot()
    
    if success:
        print("\n🚀 现在可以运行自动发送脚本：")
        print("python3 auto_wechat_sender.py")
    else:
        print("\n📖 详细配置说明请查看: WECHAT_SETUP.md")

if __name__ == "__main__":
    main()
