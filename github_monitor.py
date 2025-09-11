#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GitHub Actions监控脚本 - 执行一次检查
"""

import pandas as pd
import smtplib
import requests
import os
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from dotenv import load_dotenv
import io

# 加载环境变量
load_dotenv()

class GitHubExpiryChecker:
    def __init__(self):
        self.excel_file = "yxc.xlsx"
        self.backup_file = "yxc_backup.xlsx"
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
    
    def backup_excel(self):
        """备份Excel文件"""
        try:
            import shutil
            shutil.copy2(self.excel_file, self.backup_file)
            print(f"✅ Excel文件已备份: {self.backup_file}")
        except Exception as e:
            print(f"❌ 备份失败: {e}")
    
    def read_excel_file(self):
        """读取Excel文件"""
        try:
            df = pd.read_excel(self.excel_file)
            print(f"✅ 成功读取Excel文件: {self.excel_file}")
            print(f"📊 表格形状: {df.shape}")
            return df
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
                # 确保日期格式正确，保持与原始数据类型一致
                if pd.api.types.is_integer_dtype(df[columns['start_date']]):
                    df.at[idx, columns['start_date']] = int(new_start_date)
                else:
                    df.at[idx, columns['start_date']] = str(new_start_date)
                
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
    
    def create_table_image(self, df, title="项目监控表"):
        """创建表格图片"""
        try:
            # 设置中文字体
            plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False
            
            # 创建图形
            fig, ax = plt.subplots(figsize=(16, 10))
            ax.axis('tight')
            ax.axis('off')
            
            # 准备表格数据
            table_data = []
            for idx, row in df.iterrows():
                table_data.append([
                    row.get(' 店铺名称', ''),
                    row.get('地址', ''),
                    str(row.get('总天', '')),
                    str(row.get('剩余', '')),
                    str(row.get('开始时间', ''))
                ])
            
            # 创建表格
            table = ax.table(
                cellText=table_data,
                colLabels=['店铺名称', '地址', '总天', '剩余', '开始时间'],
                cellLoc='center',
                loc='center',
                bbox=[0, 0, 1, 1]
            )
            
            # 设置表格样式
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 2)
            
            # 设置标题
            plt.title(title, fontsize=16, fontweight='bold', pad=20)
            
            # 高亮剩余天数为0的行
            for i in range(1, len(table_data) + 1):
                remaining_days = table_data[i-1][3]
                if remaining_days == '0':
                    for j in range(5):
                        table[(i, j)].set_facecolor('#ffcccc')  # 红色背景
                        table[(i, j)].set_text_props(weight='bold', color='red')  # 红色粗体字体
            
            # 保存图片到内存
            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
            img_buffer.seek(0)
            
            plt.close()
            return img_buffer
            
        except Exception as e:
            print(f"❌ 创建表格图片失败: {e}")
            return None
    
    def send_email_notification(self, expired_items, updated_items, df=None):
        """发送邮件通知"""
        if not self.notification_config['email']['enabled']:
            print("邮件通知未启用")
            return False
        
        try:
            config = self.notification_config['email']
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # 创建邮件内容
            subject = f"智能监控报告 - {len(expired_items)}个到期项目，{len(updated_items)}个项目已重置"
            
            # HTML邮件内容
            html_body = f"""
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    .header {{ background-color: #f0f0f0; padding: 15px; border-radius: 5px; }}
                    .section {{ margin: 20px 0; }}
                    .expired {{ background-color: #ffebee; padding: 10px; border-left: 4px solid #f44336; }}
                    .updated {{ background-color: #e8f5e8; padding: 10px; border-left: 4px solid #4caf50; }}
                    .table-section {{ margin: 20px 0; }}
                    .footer {{ margin-top: 30px; font-size: 12px; color: #666; }}
                </style>
            </head>
            <body>
                <div class="header">
                    <h2>🤖 智能监控报告</h2>
                    <p><strong>时间:</strong> {current_time}</p>
                    <p><strong>到期项目数量:</strong> {len(expired_items)}</p>
                    <p><strong>重置项目数量:</strong> {len(updated_items)}</p>
                </div>
                
                <div class="section">
                    <h3>📊 当前项目状态表</h3>
                    <p>以下是当前所有项目的状态表格：</p>
                    <img src="cid:table_image" alt="项目状态表" style="max-width: 100%; height: auto;">
                </div>
            """
            
            if expired_items:
                html_body += """
                <div class="section expired">
                    <h3>🚨 到期项目详情</h3>
                    <ul>
                """
                for item in expired_items:
                    item_data = item['data']
                    store_name = item_data.get(' 店铺名称', '未知店铺')
                    address = item_data.get('地址', '未知地址')
                    total_days = item_data.get('总天', '未知')
                    html_body += f"<li><strong>行 {item['row']}:</strong> {store_name} - {address} - {total_days}天</li>"
                html_body += "</ul></div>"
            
            if updated_items:
                html_body += """
                <div class="section updated">
                    <h3>🔄 已重置项目</h3>
                    <ul>
                """
                for item in updated_items:
                    html_body += f"<li><strong>行 {item['row']}:</strong> {item['name']} - {item['address']} - 重置为{item['total_days']}天</li>"
                html_body += "</ul></div>"
            
            html_body += """
                <div class="footer">
                    <p>此邮件由智能监控系统自动发送，请及时处理到期项目！</p>
                    <p>如有问题，请联系系统管理员。</p>
                </div>
            </body>
            </html>
            """
            
            # 创建邮件对象
            msg = MIMEMultipart('related')
            msg['From'] = config['username']
            
            # 处理多个接收者
            to_emails = [email.strip() for email in config['to_email'].split(',')]
            msg['To'] = ', '.join(to_emails)
            msg['Subject'] = subject
            
            # 添加HTML内容
            html_part = MIMEText(html_body, 'html', 'utf-8')
            msg.attach(html_part)
            
            # 创建并添加表格图片
            if df is not None:
                table_image = self.create_table_image(df, "项目监控状态表")
                if table_image:
                    image_part = MIMEImage(table_image.getvalue())
                    image_part.add_header('Content-ID', '<table_image>')
                    image_part.add_header('Content-Disposition', 'inline', filename='table.png')
                    msg.attach(image_part)
                    print("✅ 表格图片已添加到邮件")
            
            # 发送邮件
            if config['smtp_port'] == 465:
                server = smtplib.SMTP_SSL(config['smtp_server'], config['smtp_port'])
            else:
                server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
                server.starttls()
            
            server.login(config['username'], config['password'])
            server.send_message(msg, to_addrs=to_emails)
            server.quit()
            
            print("✅ 增强版邮件通知发送成功")
            print(f"📧 邮件已发送到: {', '.join(to_emails)}")
            print("📊 邮件包含表格图片和详细状态信息")
            return True
            
        except Exception as e:
            print(f"❌ 发送邮件通知失败: {e}")
            return False
    
    def create_notification_content(self, expired_items, updated_items, notification_type):
        """创建通知内容"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        if notification_type == "短信":
            content = f"【GitHub监控】{len(expired_items)}个到期，{len(updated_items)}个已重置。时间:{current_time[:10]}"
        else:
            content = f"""
GitHub监控报告 - 项目到期提醒

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
    
    def run_check(self):
        """执行一次检查流程"""
        print(f"\n=== 开始执行GitHub监控任务 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
        
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
        
        # 第三步：更新到期项目（重置开始时间和剩余天数）
        updated_items = self.update_expired_items(df, columns)
        print(f"🔄 重置了 {len(updated_items)} 个到期项目")
        
        # 第四步：发送通知
        if expired_items or updated_items:
            print(f"📧 发送通知")
            self.send_email_notification(expired_items, updated_items, df)
        
        # 保存更新后的Excel文件
        if self.save_excel_file(df):
            print("💾 Excel文件已更新并保存")
        
        print("=== GitHub监控任务完成 ===\n")

def main():
    """主函数"""
    checker = GitHubExpiryChecker()
    checker.run_check()

if __name__ == "__main__":
    main() 