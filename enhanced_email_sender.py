#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import smtplib
import os
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from datetime import datetime
from dotenv import load_dotenv
import io
import base64

# 加载环境变量
load_dotenv()

class EnhancedEmailSender:
    def __init__(self):
        self.excel_file = "yxc.xlsx"
        self.notification_config = {
            'email': {
                'enabled': os.getenv('EMAIL_ENABLED', 'false').lower() == 'true',
                'smtp_server': os.getenv('SMTP_SERVER', 'smtp.gmail.com'),
                'smtp_port': int(os.getenv('SMTP_PORT', '587')),
                'username': os.getenv('EMAIL_USERNAME'),
                'password': os.getenv('EMAIL_PASSWORD'),
                'to_email': os.getenv('TO_EMAIL', 'yangxingchao87@163.com')
            }
        }
    
    def create_table_image(self, df, title="项目监控表"):
        """创建表格图片"""
        try:
            # 设置中文字体
            plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False
            
            # 创建图形
            fig, ax = plt.subplots(figsize=(20, 12))
            ax.axis('tight')
            ax.axis('off')
            
            # 准备表格数据
            table_data = []
            for idx, row in df.iterrows():
                table_data.append([
                    str(row.get('行号', '')),
                    row.get(' 店铺名称', ''),
                    row.get('地址', ''),
                    str(row.get('总天', '')),
                    str(row.get('剩余', '')),
                    str(row.get('开始时间', '')),
                    str(row.get('备注1', '')),
                    str(row.get('备注2', ''))
                ])
            
            # 创建表格
            table = ax.table(
                cellText=table_data,
                colLabels=['行号', '店铺名称', '地址', '总天', '剩余', '开始时间', '备注1', '备注2'],
                cellLoc='center',
                loc='center',
                bbox=[0, 0, 1, 1]
            )
            
            # 设置表格样式
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 2)
            
            # 设置标题（包含发送日期）
            from datetime import datetime
            current_date = datetime.now().strftime('%Y年%m月%d日')
            plt.title(f'{title} - {current_date}', fontsize=16, fontweight='bold', pad=20)
            
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
    
    def send_enhanced_email(self, expired_items, updated_items, df):
        """发送增强版邮件（包含表格图片）"""
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
            msg['To'] = config['to_email']
            msg['Subject'] = subject
            
            # 添加HTML内容
            html_part = MIMEText(html_body, 'html', 'utf-8')
            msg.attach(html_part)
            
            # 创建并添加表格图片
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
            server.send_message(msg)
            server.quit()
            
            print("✅ 增强版邮件发送成功！")
            print(f"📧 邮件已发送到: {config['to_email']}")
            print("📊 邮件包含表格图片和详细状态信息")
            return True
            
        except Exception as e:
            print(f"❌ 发送增强版邮件失败: {e}")
            return False
    
    def send_real_data_email(self):
        """发送真实数据的增强版邮件"""
        print("=== 发送真实数据邮件 ===")
        
        # 读取Excel文件
        try:
            df = pd.read_excel(self.excel_file)
            print(f"✅ 成功读取Excel文件，共 {len(df)} 行数据")
        except Exception as e:
            print(f"❌ 读取Excel文件失败: {e}")
            return
        
        # 查找关键列
        columns = self.find_columns(df)
        
        # 检查真实的过期项目
        expired_items = []
        for idx, row in df.iterrows():
            remaining = row[columns['remaining']]
            if pd.notna(remaining) and int(remaining) == 0:
                expired_items.append({
                    'row': row.get('行号', idx + 1),
                    'data': row.to_dict()
                })
        
        # 创建空的updated_items（因为这是查看当前状态，不是重置后的状态）
        updated_items = []
        
        print(f"📊 发现 {len(expired_items)} 个过期项目")
        
        # 发送增强版邮件
        self.send_enhanced_email(expired_items, updated_items, df)
    
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

def main():
    sender = EnhancedEmailSender()
    sender.send_real_data_email()

if __name__ == "__main__":
    main()
