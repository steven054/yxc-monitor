#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import smtplib
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

def create_and_save_table_image():
    """创建并保存表格图片"""
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
        
        # 保存图片到文件
        filename = 'email_table_with_notes.png'
        plt.savefig(filename, format='png', dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"✅ 表格图片已保存为: {filename}")
        return filename
        
    except Exception as e:
        print(f"❌ 创建表格图片失败: {e}")
        return None

def send_email_with_attachment():
    """发送带附件的邮件"""
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
        
        print(f"📧 准备发送带附件的邮件到: {config['to_emails']}")
        
        # 创建并保存表格图片
        print("🖼️  创建表格图片...")
        image_filename = create_and_save_table_image()
        if not image_filename:
            print("❌ 无法创建表格图片")
            return False
        
        # 创建邮件
        msg = MIMEMultipart()
        msg['From'] = config['from_email']
        msg['To'] = ', '.join(config['to_emails'])
        msg['Subject'] = '店铺监控数据表 - 包含行号和备注（附件版本）'
        
        # 创建邮件正文
        body = """
        您好！
        
        这是店铺监控数据表，包含行号和备注信息。
        
        表格包含以下列：
        - 行号：显示每行的序号
        - 店铺名称：店铺名称
        - 地址：店铺地址
        - 总天：总天数
        - 剩余：剩余天数
        - 开始时间：开始日期
        - 备注1：大桶数量信息
        - 备注2：状态信息
        
        表格图片已作为附件发送，请查看附件中的PNG文件。
        
        数据统计：
        - 总行数：34行
        - 备注1分布：大桶1个(26个)、大桶2个(8个)
        - 备注2分布：正常(15个)、小桶1个(7个)、一周内到期(7个)、即将到期(5个)
        
        谢谢！
        """
        
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # 添加图片附件
        print("📎 添加图片附件...")
        with open(image_filename, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f'attachment; filename= {image_filename}',
        )
        msg.attach(part)
        print("✅ 图片附件已添加")
        
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
        
        print("✅ 带附件的邮件发送成功！")
        print(f"📧 邮件已发送到: {', '.join(config['to_emails'])}")
        print(f"📎 附件文件: {image_filename}")
        print("📊 邮件包含完整的表格图片附件，显示行号和备注信息")
        
        return True
        
    except Exception as e:
        print(f"❌ 发送邮件失败: {e}")
        return False

if __name__ == "__main__":
    print("=== 发送带附件的邮件 ===")
    success = send_email_with_attachment()
    if success:
        print("\n🎉 带附件的邮件发送成功！")
        print("📧 请检查邮箱，邮件应该包含表格图片附件")
        print("📎 附件文件名: email_table_with_notes.png")
    else:
        print("\n❌ 带附件的邮件发送失败")
