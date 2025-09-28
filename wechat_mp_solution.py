#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
import json
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import io
import base64
from datetime import datetime
import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

class WeChatMPSender:
    def __init__(self):
        # 微信公众号配置
        self.app_id = os.getenv('WECHAT_MP_APP_ID', '')
        self.app_secret = os.getenv('WECHAT_MP_APP_SECRET', '')
        self.access_token = None
        
    def get_access_token(self):
        """获取微信公众号access_token"""
        if not self.app_id or not self.app_secret:
            print("❌ 未配置微信公众号APP_ID和APP_SECRET")
            return False
            
        try:
            url = f"https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid={self.app_id}&secret={self.app_secret}"
            response = requests.get(url)
            result = response.json()
            
            if 'access_token' in result:
                self.access_token = result['access_token']
                print("✅ 获取access_token成功")
                return True
            else:
                print(f"❌ 获取access_token失败: {result}")
                return False
                
        except Exception as e:
            print(f"❌ 获取access_token时出错: {e}")
            return False
    
    def upload_image(self, image_buffer):
        """上传图片到微信服务器"""
        if not self.access_token:
            print("❌ 没有有效的access_token")
            return None
            
        try:
            url = f"https://api.weixin.qq.com/cgi-bin/media/upload?access_token={self.access_token}&type=image"
            
            files = {
                'media': ('image.png', image_buffer, 'image/png')
            }
            
            response = requests.post(url, files=files)
            result = response.json()
            
            if 'media_id' in result:
                print("✅ 图片上传成功")
                return result['media_id']
            else:
                print(f"❌ 图片上传失败: {result}")
                return None
                
        except Exception as e:
            print(f"❌ 上传图片时出错: {e}")
            return None
    
    def send_template_message(self, openid, template_id, data):
        """发送模板消息"""
        if not self.access_token:
            print("❌ 没有有效的access_token")
            return False
            
        try:
            url = f"https://api.weixin.qq.com/cgi-bin/message/template/send?access_token={self.access_token}"
            
            message_data = {
                "touser": openid,
                "template_id": template_id,
                "data": data
            }
            
            response = requests.post(url, json=message_data)
            result = response.json()
            
            if result.get('errcode') == 0:
                print("✅ 模板消息发送成功")
                return True
            else:
                print(f"❌ 模板消息发送失败: {result}")
                return False
                
        except Exception as e:
            print(f"❌ 发送模板消息时出错: {e}")
            return False

def main():
    """主函数"""
    print("=== 微信公众号自动发送工具 ===")
    print()
    print("📱 设置步骤：")
    print("1. 注册微信公众号（服务号）")
    print("2. 获取APP_ID和APP_SECRET")
    print("3. 创建模板消息")
    print("4. 获取用户的openid")
    print("5. 配置到.env文件中")
    print()
    print("⚠️  注意：此方案需要用户关注公众号，且需要申请模板消息权限")
    
    # 检查配置
    app_id = os.getenv('WECHAT_MP_APP_ID', '')
    app_secret = os.getenv('WECHAT_MP_APP_SECRET', '')
    
    if not app_id or not app_secret:
        print("\n❌ 请在.env文件中配置微信公众号参数")
        print("\n🔧 配置示例：")
        print("WECHAT_MP_APP_ID=your_app_id")
        print("WECHAT_MP_APP_SECRET=your_app_secret")
        return
    
    sender = WeChatMPSender()
    
    # 获取access_token
    if not sender.get_access_token():
        return
    
    print("📤 微信公众号自动发送功能已准备就绪")
    print("💡 注意：此方案需要用户主动关注公众号")

if __name__ == "__main__":
    main()
