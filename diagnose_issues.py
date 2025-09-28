#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
诊断邮件发送问题
"""

import os
from dotenv import load_dotenv

def diagnose_issues():
    """诊断邮件发送问题"""
    print("=== 邮件发送问题诊断 ===\n")
    
    # 加载环境变量
    load_dotenv()
    
    # 检查邮件配置
    print("1. 检查邮件配置:")
    email_enabled = os.getenv('EMAIL_ENABLED', 'false').lower() == 'true'
    smtp_server = os.getenv('SMTP_SERVER')
    smtp_port = os.getenv('SMTP_PORT')
    username = os.getenv('EMAIL_USERNAME')
    password = os.getenv('EMAIL_PASSWORD')
    to_email = os.getenv('TO_EMAIL')
    
    print(f"   EMAIL_ENABLED: {email_enabled}")
    print(f"   SMTP_SERVER: {smtp_server}")
    print(f"   SMTP_PORT: {smtp_port}")
    print(f"   EMAIL_USERNAME: {username}")
    print(f"   EMAIL_PASSWORD: {'已设置' if password else '未设置'}")
    print(f"   TO_EMAIL: {to_email}")
    
    if not email_enabled:
        print("   ❌ 邮件功能未启用")
    else:
        print("   ✅ 邮件功能已启用")
    
    print()
    
    # 检查文件是否存在
    print("2. 检查关键文件:")
    files_to_check = [
        'yxc.xlsx',
        'github_monitor.py',
        '.github/workflows/monitor.yml',
        '.env'
    ]
    
    for file in files_to_check:
        if os.path.exists(file):
            print(f"   ✅ {file} 存在")
        else:
            print(f"   ❌ {file} 不存在")
    
    print()
    
    # 检查GitHub Actions配置
    print("3. GitHub Actions配置检查:")
    print("   请手动检查以下项目:")
    print("   - 是否已推送到GitHub仓库")
    print("   - GitHub Secrets是否已设置:")
    print("     * EMAIL_ENABLED=true")
    print("     * SMTP_SERVER=smtp.163.com")
    print("     * SMTP_PORT=465")
    print("     * EMAIL_USERNAME=shidewei054@163.com")
    print("     * EMAIL_PASSWORD=CEpJp32m4rX6weNH")
    print("     * TO_EMAIL=yangxingchao87@163.com,408838485@qq.com")
    print("   - Actions是否已启用")
    print("   - 定时任务是否设置为每天7点")
    
    print()
    
    # 建议
    print("4. 建议的解决步骤:")
    print("   1. 确保网络连接正常")
    print("   2. 推送最新代码到GitHub:")
    print("      git push origin main")
    print("   3. 在GitHub仓库中设置Secrets")
    print("   4. 手动触发一次GitHub Actions测试")
    print("   5. 检查Actions执行日志")
    
    print()
    print("5. 当前邮件发送时间:")
    print("   - UTC时间: 23:00 (晚上11点)")
    print("   - 北京时间: 07:00 (早上7点)")

if __name__ == "__main__":
    diagnose_issues()

