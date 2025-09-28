#!/usr/bin/env python3
# -*- coding: utf-8 -*-

def wechat_webhook_guide():
    """企业微信机器人Webhook获取详细指南"""
    print("=== 企业微信机器人Webhook获取指南 ===")
    print()
    
    print("📱 方法1：通过企业微信群添加机器人")
    print("1. 打开企业微信群聊")
    print("2. 点击群聊右上角的'...'菜单")
    print("3. 选择'群机器人'或'群助手'")
    print("4. 点击'添加机器人'")
    print("5. 选择'自定义机器人'")
    print("6. 填写机器人名称和描述")
    print("7. 点击'添加'")
    print("8. 添加成功后，会显示一个webhook地址")
    print("9. 地址格式：https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxxxxxxx")
    print()
    
    print("🔍 如果添加机器人后没有看到webhook地址：")
    print("1. 点击机器人头像")
    print("2. 选择'设置'或'详情'")
    print("3. 查找'API地址'、'Webhook'或'接收消息URL'")
    print("4. 复制完整的地址")
    print()
    
    print("📋 方法2：通过企业微信管理后台")
    print("1. 访问：https://work.weixin.qq.com/")
    print("2. 使用企业微信管理员账号登录")
    print("3. 进入'应用管理'")
    print("4. 点击'创建应用'")
    print("5. 填写应用信息")
    print("6. 在应用详情中找到'接收消息'部分")
    print("7. 复制webhook地址")
    print()
    
    print("⚠️  重要提示：")
    print("- webhook地址必须以'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key='开头")
    print("- key参数是唯一的，不能泄露给他人")
    print("- 如果机器人被删除，webhook地址就会失效")
    print()
    
    print("🧪 获取到webhook地址后，运行测试：")
    print("python3 setup_wechat_bot.py")

if __name__ == "__main__":
    wechat_webhook_guide()
