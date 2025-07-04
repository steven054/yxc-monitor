#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import requests
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import os
import json
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

class CloudExpiryChecker:
    def __init__(self):
        """初始化云平台监控器"""
        self.excel_file = os.getenv('EXCEL_FILE', 'yxc.xlsx')
        self.backup_dir = 'backups'
        
        # 确保备份目录存在
        if not os.path.exists(self.backup_dir):
            os.makedirs(self.backup_dir)
        
        # 通知配置
        self.notification_config = {
            'email': {
                'enabled': os.getenv('EMAIL_ENABLED', 'false').lower() == 'true',
                'smtp_server': os.getenv('SMTP_SERVER', 'smtp.163.com'),
                'smtp_port': int(os.getenv('SMTP_PORT', '465')),
                'username': os.getenv('EMAIL_USERNAME', 'shidewei054@163.com'),
                'password': os.getenv('EMAIL_PASSWORD', 'CEpJp32m4rX6weNH'),
                'to_email': os.getenv('TO_EMAIL', '408838485@qq.com')
            },
            'wechat': {
                'enabled': os.getenv('WECHAT_ENABLED', 'false').lower() == 'true',
                'webhook_url': os.getenv('WECHAT_WEBHOOK_URL', '')
            },
            'sms': {
                'enabled': os.getenv('SMS_ENABLED', 'false').lower() == 'true',
                'api_url': os.getenv('SMS_API_URL', ''),
                'api_key': os.getenv('SMS_API_KEY', ''),
                'phone_number': os.getenv('SMS_PHONE_NUMBER', '')
            }
        }
    
    def backup_excel(self):
        """备份Excel文件"""
        try:
            if os.path.exists(self.excel_file):
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                backup_file = os.path.join(self.backup_dir, f'{self.excel_file.replace(".xlsx", "")}_backup_{timestamp}.xlsx')
                
                # 复制文件
                import shutil
                shutil.copy2(self.excel_file, backup_file)
                print(f"💾 已备份Excel文件: {backup_file}")
                return True
        except Exception as e:
            print(f"❌ 备份Excel文件失败: {e}")
        return False
    
    def read_excel_file(self):
        """读取Excel文件"""
        try:
            if os.path.exists(self.excel_file):
                df = pd.read_excel(self.excel_file)
                print(f"✅ 成功读取Excel文件，共 {len(df)} 行数据")
                return df
            else:
                print(f"❌ Excel文件不存在: {self.excel_file}")
                return None
        except Exception as e:
            print(f"❌ 读取Excel文件失败: {e}")
            return None
    
    def save_excel_file(self, df):
        """保存Excel文件"""
        try:
            df.to_excel(self.excel_file, index=False)
            print(f"✅ Excel文件已保存: {self.excel_file}")
            return True
        except Exception as e:
            print(f"❌ 保存Excel文件失败: {e}")
            return False
    
    def find_columns(self, df):
        """查找关键列"""
        columns = {}
        
        # 查找剩余天数列
        possible_remaining = ['剩余天数', '剩余时间', '到期天数', '过期天数', '天数', 'days', 'remaining_days', '剩余', '到期', '过期']
        for col in df.columns:
            col_str = str(col).lower()
            for name in possible_remaining:
                if name in col_str:
                    columns['remaining'] = col
                    break
            if 'remaining' in columns:
                break
        
        # 查找总天数列
        possible_total = ['总天数', '总时间', '总天', 'total_days', 'total', '天']
        for col in df.columns:
            col_str = str(col).lower()
            for name in possible_total:
                if name in col_str:
                    columns['total'] = col
                    break
            if 'total' in columns:
                break
        
        # 查找开始时间列
        possible_start = ['开始时间', '开始日期', 'start_date', 'start_time', '开始']
        for col in df.columns:
            col_str = str(col).lower()
            for name in possible_start:
                if name in col_str:
                    columns['start_date'] = col
                    break
            if 'start_date' in columns:
                break
        
        print(f"🎯 找到的列: {columns}")
        return columns
    
    def update_remaining_days(self, df, columns):
        """更新所有项目的剩余天数 - 每天减1"""
        updated_count = 0
        
        for idx, row in df.iterrows():
            current_remaining = row[columns['remaining']]
            
            # 如果当前剩余天数为0，跳过（这些会在update_expired_items中处理）
            if pd.notna(current_remaining) and int(current_remaining) == 0:
                print(f"⏭️  跳过减1: {row.get(' 店铺名称', f'行{idx+1}')} 已经是0天")
                continue
            
            # 如果剩余天数不为0，每天减1
            if pd.notna(current_remaining) and int(current_remaining) > 0:
                old_remaining = int(current_remaining)
                new_remaining = max(0, old_remaining - 1)  # 不能小于0
                
                df.at[idx, columns['remaining']] = new_remaining
                updated_count += 1
                print(f"📅 剩余天数减1: {row.get(' 店铺名称', f'行{idx+1}')} {old_remaining} → {new_remaining}")
        
        return updated_count
    
    def check_expiry_items(self, df, columns):
        """检查剩余天数为0的项目"""
        expired_items = []
        
        try:
            # 确保剩余天数列是数值类型
            df[columns['remaining']] = pd.to_numeric(df[columns['remaining']], errors='coerce')
            
            # 查找剩余天数为0的项目
            expired_mask = df[columns['remaining']] == 0
            expired_df = df[expired_mask]
            
            if not expired_df.empty:
                print(f"🚨 发现 {len(expired_df)} 个剩余天数为0的项目:")
                for idx, row in expired_df.iterrows():
                    item_info = {
                        'row': idx + 1,
                        'data': row.to_dict()
                    }
                    expired_items.append(item_info)
                    print(f"  行 {idx + 1}: {dict(row)}")
            else:
                print("✅ 没有发现剩余天数为0的项目")
                
        except Exception as e:
            print(f"❌ 检查剩余天数时出错: {e}")
        
        return expired_items
    
    def update_expired_items(self, df, columns):
        """更新到期项目"""
        updated_items = []
        current_date = datetime.now()
        
        for idx, row in df.iterrows():
            remaining = row[columns['remaining']]
            total = row[columns['total']]
            start_date = row[columns['start_date']]
            
            # 如果剩余天数为0，需要重置
            if pd.notna(remaining) and int(remaining) == 0:
                print(f"🔄 重置项目: {row.get(' 店铺名称', f'行{idx+1}')}")
                
                # 更新开始时间为今天
                new_start_date = current_date.strftime('%Y%m%d')
                df.at[idx, columns['start_date']] = int(new_start_date)
                
                # 重置剩余天数为总天数
                df.at[idx, columns['remaining']] = total
                
                updated_items.append({
                    'row': idx + 1,
                    'name': row.get(' 店铺名称', f'行{idx+1}'),
                    'address': row.get('地址', '未知地址'),
                    'total_days': total,
                    'old_start': start_date,
                    'new_start': new_start_date
                })
        
        return updated_items
    
    def send_notifications(self, expired_items, updated_items):
        """发送通知"""
        print(f"📧 开始发送通知，共有 {len(expired_items)} 个到期项目，{len(updated_items)} 个重置项目")
        
        # 发送邮件通知
        self.send_email_notification(expired_items, updated_items)
        
        # 发送微信通知
        self.send_wechat_notification(expired_items, updated_items)
        
        # 发送短信通知
        self.send_sms_notification(expired_items, updated_items)
    
    def send_email_notification(self, expired_items, updated_items):
        """发送邮件通知"""
        if not self.notification_config['email']['enabled']:
            print("邮件通知未启用")
            return False
        
        try:
            config = self.notification_config['email']
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # 创建邮件内容
            subject = f"智能监控报告 - {len(expired_items)}个到期项目，{len(updated_items)}个项目已重置"
            body = self.create_notification_content(expired_items, updated_items, "邮件")
            
            # 创建邮件对象
            msg = MIMEMultipart()
            msg['From'] = config['username']
            msg['To'] = config['to_email']
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # 发送邮件
            if config['smtp_port'] == 465:
                server = smtplib.SMTP_SSL(config['smtp_server'], config['smtp_port'])
            else:
                server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
                server.starttls()
            
            server.login(config['username'], config['password'])
            server.send_message(msg)
            server.quit()
            
            print("✅ 邮件通知发送成功")
            return True
            
        except Exception as e:
            print(f"❌ 发送邮件通知失败: {e}")
            return False
    
    def send_wechat_notification(self, expired_items, updated_items):
        """发送微信通知"""
        if not self.notification_config['wechat']['enabled']:
            print("微信通知未启用")
            return False
        
        try:
            config = self.notification_config['wechat']
            current_time = datetime.now().strftime("%Y-%m-%d")
            
            content = f"【智能监控】{len(expired_items)}个到期，{len(updated_items)}个已重置。时间:{current_time}"
            
            payload = {
                "msgtype": "text",
                "text": {
                    "content": content
                }
            }
            
            response = requests.post(config['webhook_url'], json=payload)
            if response.status_code == 200:
                print("✅ 微信通知发送成功")
                return True
            else:
                print(f"❌ 微信通知发送失败: {response.status_code}")
                return False
                
        except Exception as e:
            print(f"❌ 发送微信通知失败: {e}")
            return False
    
    def send_sms_notification(self, expired_items, updated_items):
        """发送短信通知"""
        if not self.notification_config['sms']['enabled']:
            print("短信通知未启用")
            return False
        
        try:
            config = self.notification_config['sms']
            current_time = datetime.now().strftime("%Y-%m-%d")
            
            content = f"【智能监控】{len(expired_items)}个到期，{len(updated_items)}个已重置。{current_time}"
            
            payload = {
                "api_key": config['api_key'],
                "phone": config['phone_number'],
                "message": content
            }
            
            response = requests.post(config['api_url'], json=payload)
            if response.status_code == 200:
                print("✅ 短信通知发送成功")
                return True
            else:
                print(f"❌ 短信通知发送失败: {response.status_code}")
                return False
                
        except Exception as e:
            print(f"❌ 发送短信通知失败: {e}")
            return False
    
    def create_notification_content(self, expired_items, updated_items, notification_type):
        """创建通知内容"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if notification_type == "短信":
            content = f"【智能监控】{len(expired_items)}个到期，{len(updated_items)}个已重置。时间:{current_time[:10]}"
        else:
            content = f"""
智能监控报告 - 项目到期提醒

时间: {current_time}
到期项目数量: {len(expired_items)}
重置项目数量: {len(updated_items)}

"""
            
            if expired_items:
                content += "\n🚨 到期项目详情:\n"
                for item in expired_items:
                    item_data = item['data']
                    store_name = item_data.get(' 店铺名称', '未知店铺')
                    address = item_data.get('地址', '未知地址')
                    total_days = item_data.get('总天', '未知')
                    content += f"  行 {item['row']}: {store_name} - {address} - {total_days}天\n"
            
            if updated_items:
                content += "\n🔄 已重置项目:\n"
                for item in updated_items:
                    content += f"  行 {item['row']}: {item['name']} - {item['address']} - 重置为{item['total_days']}天\n"
            
            content += "\n请及时处理这些到期项目！"
        
        return content
    
    def send_congratulations_notification(self):
        """发送恭喜通知"""
        print("🎉 发送恭喜通知...")
        
        # 发送恭喜邮件
        self.send_congratulations_email()
        
        # 发送恭喜微信通知
        self.send_congratulations_wechat()
        
        # 发送恭喜短信
        self.send_congratulations_sms()
    
    def send_congratulations_email(self):
        """发送恭喜邮件"""
        if not self.notification_config['email']['enabled']:
            print("邮件通知未启用")
            return False
        
        try:
            config = self.notification_config['email']
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            subject = "🎉 恭喜！所有项目运行正常"
            body = f"""
🎉 恭喜通知 - 项目状态良好

时间: {current_time}

🎊 好消息！今天检查发现所有项目都运行正常，没有到期的项目需要处理。

📊 项目状态:
✅ 所有项目都在有效期内
✅ 无需任何操作
✅ 继续保持良好状态

💡 温馨提示:
- 继续保持当前的优秀管理
- 定期检查项目进度
- 如有需要可以调整项目周期

祝您工作顺利！🎊
"""
            
            msg = MIMEMultipart()
            msg['From'] = config['username']
            msg['To'] = config['to_email']
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            if config['smtp_port'] == 465:
                server = smtplib.SMTP_SSL(config['smtp_server'], config['smtp_port'])
            else:
                server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
                server.starttls()
            
            server.login(config['username'], config['password'])
            server.send_message(msg)
            server.quit()
            
            print("✅ 恭喜邮件发送成功")
            return True
            
        except Exception as e:
            print(f"❌ 发送恭喜邮件失败: {e}")
            return False
    
    def send_congratulations_wechat(self):
        """发送恭喜微信通知"""
        if not self.notification_config['wechat']['enabled']:
            print("微信通知未启用")
            return False
        
        try:
            config = self.notification_config['wechat']
            current_time = datetime.now().strftime("%Y-%m-%d")
            
            content = f"🎉恭喜！所有项目运行正常，无需处理。时间:{current_time}"
            
            payload = {
                "msgtype": "text",
                "text": {
                    "content": content
                }
            }
            
            response = requests.post(config['webhook_url'], json=payload)
            if response.status_code == 200:
                print("✅ 恭喜微信通知发送成功")
                return True
            else:
                print(f"❌ 恭喜微信通知发送失败: {response.status_code}")
                return False
                
        except Exception as e:
            print(f"❌ 发送恭喜微信通知失败: {e}")
            return False
    
    def send_congratulations_sms(self):
        """发送恭喜短信"""
        if not self.notification_config['sms']['enabled']:
            print("短信通知未启用")
            return False
        
        try:
            config = self.notification_config['sms']
            current_time = datetime.now().strftime("%Y-%m-%d")
            
            content = f"🎉恭喜！所有项目正常，无需处理。{current_time}"
            
            payload = {
                "api_key": config['api_key'],
                "phone": config['phone_number'],
                "message": content
            }
            
            response = requests.post(config['api_url'], json=payload)
            if response.status_code == 200:
                print("✅ 恭喜短信发送成功")
                return True
            else:
                print(f"❌ 恭喜短信发送失败: {response.status_code}")
                return False
                
        except Exception as e:
            print(f"❌ 发送恭喜短信失败: {e}")
            return False
    
    def run_check(self):
        """执行智能检查流程"""
        print(f"\n=== 开始执行智能检查任务 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
        
        # 备份Excel文件
        self.backup_excel()
        
        # 读取Excel文件
        df = self.read_excel_file()
        if df is None:
            print("❌ 无法读取Excel文件，检查任务终止")
            return
        
        # 查找关键列
        columns = self.find_columns(df)
        
        # 第一步：先更新剩余天数（每天减1）
        updated_count = self.update_remaining_days(df, columns)
        print(f"📅 更新了 {updated_count} 个项目的剩余天数（每天减1）")
        
        # 第二步：检查到期项目（减1后可能变成0的项目）
        expired_items = self.check_expiry_items(df, columns)
        
        # 第三步：如果有到期项目，发送通知
        if expired_items:
            print(f"📧 发现 {len(expired_items)} 个到期项目，发送通知")
            self.send_notifications(expired_items, [])
        
        # 第四步：更新到期项目（重置开始时间和剩余天数）
        updated_items = self.update_expired_items(df, columns)
        print(f"🔄 重置了 {len(updated_items)} 个到期项目")
        
        # 保存更新后的Excel文件
        if self.save_excel_file(df):
            print("💾 Excel文件已更新并保存")
        
        # 如果没有到期项目，发送恭喜通知
        if not expired_items:
            print("🎉 没有到期项目，发送恭喜通知")
            self.send_congratulations_notification()
        
        print("=== 智能检查任务完成 ===\n")

def main():
    """主函数 - 云平台版本"""
    checker = CloudExpiryChecker()
    
    # 直接执行一次检查（适合云平台定时任务）
    print("🚀 云平台监控脚本启动")
    print("📋 功能: 自动更新剩余天数 + 重置到期项目 + 发送通知")
    
    checker.run_check()

if __name__ == "__main__":
    main() 