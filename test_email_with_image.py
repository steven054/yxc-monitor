#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import smtplib
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os

# 加载环境变量
load_dotenv()

def create_table_image():
    """创建表格图片"""
    try:
        # 读取Excel文件
        df = pd.read_excel('yxc.xlsx')
        print(f"✅ 成功读取Excel文件，共 {len(df)} 行数据")
        
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
        table.set_fontsize(8)
        table.scale(1, 1.5)
        
        # 设置标题（包含发送日期）
        from datetime import datetime
        current_date = datetime.now().strftime('%Y年%m月%d日')
        plt.title(f'店铺监控数据表（包含行号和备注）- {current_date}', fontsize=16, fontweight='bold', pad=20)
        
        # 保存图片到内存
        img_buffer = io.BytesIO()
        plt.savefig(img_buffer, format='png', dpi=300, bbox_inches='tight')
        img_buffer.seek(0)
        
        plt.close()
        
        print("✅ 表格图片创建成功")
        return img_buffer
        
    except Exception as e:
        print(f"❌ 创建表格图片失败: {e}")
        return None

def send_test_email():
    """发送测试邮件"""
    try:
        # 邮件配置
        config = {
            'smtp_server': os.getenv('SMTP_SERVER', 'smtp.163.com'),
            'smtp_port': int(os.getenv('SMTP_PORT', 465)),
            'username': os.getenv('EMAIL_USERNAME', 'shidewei054@163.com'),
            'password': os.getenv('EMAIL_PASSWORD', ''),
            'from_email': os.getenv('FROM_EMAIL', 'shidewei054@163.com'),
            'to_emails': os.getenv('TO_EMAILS', 'yangxingchao87@163.com,408838485@qq.com').split(',')
        }
        
        print(f"📧 准备发送测试邮件到: {config['to_emails']}")
        
        # 创建邮件
        msg = MIMEMultipart('related')
        msg['From'] = config['from_email']
        msg['To'] = ', '.join(config['to_emails'])
        msg['Subject'] = '测试邮件 - 包含行号和备注的表格图片'
        
        # 创建HTML内容
        html_body = """
        <html>
        <head>
            <meta charset="utf-8">
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                .header { background-color: #f0f0f0; padding: 15px; border-radius: 5px; }
                .content { margin: 20px 0; }
                .table-container { text-align: center; margin: 20px 0; }
                img { max-width: 100%; height: auto; border: 1px solid #ddd; }
            </style>
        </head>
        <body>
            <div class="header">
                <h2>📊 店铺监控数据表测试</h2>
                <p>此邮件包含完整的表格图片，显示行号和备注信息</p>
            </div>
            
            <div class="content">
                <h3>📋 表格包含以下列：</h3>
                <ul>
                    <li>行号 - 显示每行的序号</li>
                    <li>店铺名称 - 店铺名称</li>
                    <li>地址 - 店铺地址</li>
                    <li>总天 - 总天数</li>
                    <li>剩余 - 剩余天数</li>
                    <li>开始时间 - 开始日期</li>
                    <li>备注1 - 大桶数量信息</li>
                    <li>备注2 - 状态信息</li>
                </ul>
            </div>
            
            <div class="table-container">
                <h3>📊 完整数据表格：</h3>
                <img src="cid:table_image" alt="店铺监控数据表">
            </div>
            
            <div class="content">
                <p><strong>注意：</strong>如果图片没有显示，请检查邮件客户端的设置，确保允许显示图片。</p>
            </div>
        </body>
        </html>
        """
        
        # 添加HTML内容
        html_part = MIMEText(html_body, 'html', 'utf-8')
        msg.attach(html_part)
        
        # 创建并添加表格图片
        print("🖼️  创建表格图片...")
        table_image = create_table_image()
        if table_image:
            image_part = MIMEImage(table_image.getvalue())
            image_part.add_header('Content-ID', '<table_image>')
            image_part.add_header('Content-Disposition', 'inline', filename='table.png')
            msg.attach(image_part)
            print("✅ 表格图片已添加到邮件")
        else:
            print("❌ 无法创建表格图片")
            return False
        
        # 发送邮件
        print("📤 发送邮件...")
        if config['smtp_port'] == 465:
            server = smtplib.SMTP_SSL(config['smtp_server'], config['smtp_port'])
        else:
            server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
            server.starttls()
        
        server.login(config['username'], config['password'])
        server.send_message(msg)
        server.quit()
        
        print("✅ 测试邮件发送成功！")
        print(f"📧 邮件已发送到: {', '.join(config['to_emails'])}")
        print("📊 邮件包含完整的表格图片，显示行号和备注信息")
        
        return True
        
    except Exception as e:
        print(f"❌ 发送邮件失败: {e}")
        return False

if __name__ == "__main__":
    print("=== 测试邮件发送（包含图片） ===")
    success = send_test_email()
    if success:
        print("\n🎉 测试邮件发送成功！")
        print("📧 请检查邮箱，邮件应该包含完整的表格图片")
    else:
        print("\n❌ 测试邮件发送失败")
