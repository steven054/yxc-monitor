#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
每日检查脚本 - 检查Excel表格中的剩余天数并发送通知
"""

import pandas as pd
import schedule
import time
import smtplib
import requests
import os
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

class ExpiryChecker:
    def __init__(self):
        self.excel_file = "yxc.xlsx"
        self.notification_config = {
            'email': {
                'enabled': os.getenv('EMAIL_ENABLED', 'false').lower() == 'true',
                'smtp_server': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
                'smtp_port': int(os.getenv('SMTP_PORT', '587')),
                'username': os.getenv('EMAIL_USERNAME'),
                'password': os.getenv('EMAIL_PASSWORD'),
                'to_email': os.getenv('TO_EMAIL')
            },
            'wechat': {
                'enabled': os.getenv('WECHAT_ENABLED', 'false').lower() == 'true',
                'webhook_url': os.getenv('WECHAT_WEBHOOK_URL')
            },
            'sms': {
                'enabled': os.getenv('SMS_ENABLED', 'false').lower() == 'true',
                'api_key': os.getenv('SMS_API_KEY'),
                'api_url': os.getenv('SMS_API_URL'),
                'phone_number': os.getenv('SMS_PHONE_NUMBER')
            }
        }
    
    def read_excel_file(self):
        """读取Excel文件"""
        try:
            # 尝试读取Excel文件
            df = pd.read_excel(self.excel_file)
            print(f"成功读取Excel文件: {self.excel_file}")
            print(f"表格形状: {df.shape}")
            print(f"列名: {list(df.columns)}")
            return df
        except Exception as e:
            print(f"读取Excel文件失败: {e}")
            return None
    
    def find_expiry_column(self, df):
        """查找剩余天数列"""
        # 常见的剩余天数列名
        possible_names = [
            '剩余天数', '剩余时间', '到期天数', '过期天数', '天数', 'days', 'remaining_days',
            '剩余', '到期', '过期', '剩余日', '到期日', '过期日'
        ]
        
        for col in df.columns:
            col_str = str(col).lower()
            for name in possible_names:
                if name in col_str:
                    print(f"找到剩余天数列: {col}")
                    return col
        
        # 如果没有找到，返回第一列作为示例
        print(f"未找到明确的剩余天数列，使用第一列: {df.columns[0]}")
        return df.columns[0]
    
    def check_expiry_items(self, df, expiry_column):
        """检查剩余天数为0的项目"""
        expired_items = []
        
        try:
            # 尝试将列转换为数值类型
            df[expiry_column] = pd.to_numeric(df[expiry_column], errors='coerce')
            
            # 查找剩余天数为0的项目
            expired_mask = df[expiry_column] == 0
            expired_df = df[expired_mask]
            
            if not expired_df.empty:
                print(f"发现 {len(expired_df)} 个剩余天数为0的项目:")
                for idx, row in expired_df.iterrows():
                    item_info = {
                        'row': idx + 1,
                        'data': row.to_dict()
                    }
                    expired_items.append(item_info)
                    print(f"  行 {idx + 1}: {dict(row)}")
            else:
                print("没有发现剩余天数为0的项目")
                
        except Exception as e:
            print(f"检查剩余天数时出错: {e}")
        
        return expired_items
    
    def send_email_notification(self, expired_items):
        """发送邮件通知"""
        if not self.notification_config['email']['enabled']:
            print("邮件通知未启用")
            return False
        
        try:
            config = self.notification_config['email']
            
            # 创建邮件内容
            subject = f"紧急通知 - {len(expired_items)} 个项目已到期"
            body = self.create_notification_content(expired_items, "邮件")
            
            # 创建邮件对象
            msg = MIMEMultipart()
            msg['From'] = config['username']
            msg['To'] = config['to_email']
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # 发送邮件
            server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
            server.starttls()
            server.login(config['username'], config['password'])
            server.send_message(msg)
            server.quit()
            
            print("邮件通知发送成功")
            return True
            
        except Exception as e:
            print(f"发送邮件通知失败: {e}")
            return False
    
    def send_wechat_notification(self, expired_items):
        """发送微信通知"""
        if not self.notification_config['wechat']['enabled']:
            print("微信通知未启用")
            return False
        
        try:
            config = self.notification_config['wechat']
            content = self.create_notification_content(expired_items, "微信")
            
            # 发送到企业微信或钉钉webhook
            payload = {
                "msgtype": "text",
                "text": {
                    "content": content
                }
            }
            
            response = requests.post(config['webhook_url'], json=payload)
            if response.status_code == 200:
                print("微信通知发送成功")
                return True
            else:
                print(f"微信通知发送失败: {response.status_code}")
                return False
                
        except Exception as e:
            print(f"发送微信通知失败: {e}")
            return False
    
    def send_sms_notification(self, expired_items):
        """发送短信通知"""
        if not self.notification_config['sms']['enabled']:
            print("短信通知未启用")
            return False
        
        try:
            config = self.notification_config['sms']
            content = self.create_notification_content(expired_items, "短信")
            
            # 这里需要根据具体的短信服务商API进行调整
            payload = {
                "api_key": config['api_key'],
                "phone": config['phone_number'],
                "message": content
            }
            
            response = requests.post(config['api_url'], json=payload)
            if response.status_code == 200:
                print("短信通知发送成功")
                return True
            else:
                print(f"短信通知发送失败: {response.status_code}")
                return False
                
        except Exception as e:
            print(f"发送短信通知失败: {e}")
            return False
    
    def create_notification_content(self, expired_items, notification_type):
        """创建通知内容"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if notification_type == "短信":
            # 短信内容需要简短
            content = f"紧急通知: {len(expired_items)}个项目已到期，请及时处理。时间:{current_time}"
        else:
            # 邮件和微信可以包含详细信息
            content = f"""
紧急通知 - 项目到期提醒

时间: {current_time}
到期项目数量: {len(expired_items)}

到期项目详情:
"""
            for item in expired_items:
                content += f"\n行 {item['row']}: {item['data']}"
            
            content += "\n\n请及时处理这些到期项目！"
        
        return content
    
    def send_notifications(self, expired_items):
        """发送所有类型的通知"""
        if not expired_items:
            print("没有到期项目，无需发送通知")
            return
        
        print(f"开始发送通知，共有 {len(expired_items)} 个到期项目")
        
        # 发送邮件通知
        self.send_email_notification(expired_items)
        
        # 发送微信通知
        self.send_wechat_notification(expired_items)
        
        # 发送短信通知
        self.send_sms_notification(expired_items)
    
    def run_check(self):
        """执行检查流程"""
        print(f"\n=== 开始执行检查任务 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
        
        # 读取Excel文件
        df = self.read_excel_file()
        if df is None:
            print("无法读取Excel文件，检查任务终止")
            return
        
        # 查找剩余天数列
        expiry_column = self.find_expiry_column(df)
        
        # 检查到期项目
        expired_items = self.check_expiry_items(df, expiry_column)
        
        # 发送通知
        self.send_notifications(expired_items)
        
        print("=== 检查任务完成 ===\n")

def main():
    """主函数"""
    checker = ExpiryChecker()
    
    # 设置定时任务 - 每天早晨7点执行
    schedule.every().day.at("07:00").do(checker.run_check)
    
    print("定时检查脚本已启动")
    print("检查时间: 每天早晨 7:00")
    print("按 Ctrl+C 停止脚本")
    
    # 立即执行一次检查（用于测试）
    print("\n执行初始检查...")
    checker.run_check()
    
    # 运行定时任务
    while True:
        schedule.run_pending()
        time.sleep(60)  # 每分钟检查一次

if __name__ == "__main__":
    main() 