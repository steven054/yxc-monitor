#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

def test_email():
    """测试邮件发送功能"""
    print("=== 测试邮件发送功能 ===")
    
    # 邮件配置
    config = {
        'smtp_server': os.getenv('SMTP_SERVER', 'smtp.163.com'),
        'smtp_port': int(os.getenv('SMTP_PORT', '465')),
        'username': os.getenv('EMAIL_USERNAME', 'shidewei054@163.com'),
        'password': os.getenv('EMAIL_PASSWORD', 'CEpJp32m4rX6weNH'),
        'to_email': os.getenv('TO_EMAIL', 'yangxingchao87@163.com,408838485@qq.com')
    }
    
    print(f"SMTP服务器: {config['smtp_server']}:{config['smtp_port']}")
    print(f"发送者: {config['username']}")
    print(f"接收者: {config['to_email']}")
    print()
    
    try:
        # 创建邮件内容
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        subject = f"测试邮件 - {current_time}"
        body = f"""
这是一封测试邮件！

时间: {current_time}
发送者: {config['username']}
接收者: {config['to_email']}

如果您收到这封邮件，说明邮件配置正常工作！

测试内容：
- 剩余天数为0的项目会被自动重置
- 开始时间会更新为当天日期
- 剩余天数会重置为总天数

祝您使用愉快！
"""
        
        # 创建邮件对象
        msg = MIMEMultipart()
        msg['From'] = config['username']
        
        # 处理多个接收者
        to_emails = [email.strip() for email in config['to_email'].split(',')]
        msg['To'] = ', '.join(to_emails)
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        print("正在连接SMTP服务器...")
        
        # 发送邮件
        if config['smtp_port'] == 465:
            # 使用SSL连接（163邮箱推荐）
            server = smtplib.SMTP_SSL(config['smtp_server'], config['smtp_port'])
        else:
            # 使用TLS连接
            server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
            server.starttls()
        
        print("正在登录...")
        server.login(config['username'], config['password'])
        
        print("正在发送邮件...")
        server.send_message(msg, to_addrs=to_emails)
        server.quit()
        
        print("✅ 测试邮件发送成功！")
        print(f"📧 邮件已发送到: {', '.join(to_emails)}")
        print("请检查您的邮箱（包括垃圾邮件文件夹）")
        
    except Exception as e:
        print(f"❌ 发送测试邮件失败: {e}")
        print("可能的原因：")
        print("1. 邮箱密码错误")
        print("2. 网络连接问题")
        print("3. SMTP服务器配置错误")
        print("4. 邮箱需要开启SMTP服务")

if __name__ == "__main__":
    test_email()
