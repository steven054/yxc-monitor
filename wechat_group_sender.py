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

class WeChatGroupSender:
    def __init__(self):
        # 企业微信机器人webhook地址（需要在.env文件中配置）
        self.webhook_url = os.getenv('WECHAT_WEBHOOK_URL', '')
        
    def create_table_image(self):
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
            current_date = datetime.now().strftime('%Y年%m月%d日')
            plt.title(f'店铺监控数据表 - {current_date}', fontsize=16, fontweight='bold', pad=20)
            
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
    
    def send_text_message(self, message):
        """发送文本消息到微信群"""
        if not self.webhook_url:
            print("❌ 未配置企业微信机器人webhook地址")
            return False
            
        try:
            data = {
                "msgtype": "text",
                "text": {
                    "content": message
                }
            }
            
            response = requests.post(self.webhook_url, json=data)
            result = response.json()
            
            if result.get('errcode') == 0:
                print("✅ 文本消息发送成功")
                return True
            else:
                print(f"❌ 文本消息发送失败: {result}")
                return False
                
        except Exception as e:
            print(f"❌ 发送文本消息失败: {e}")
            return False
    
    def send_image_message(self, image_buffer):
        """发送图片消息到微信群"""
        if not self.webhook_url:
            print("❌ 未配置企业微信机器人webhook地址")
            return False
            
        try:
            # 将图片转换为base64
            image_base64 = base64.b64encode(image_buffer.getvalue()).decode('utf-8')
            
            data = {
                "msgtype": "image",
                "image": {
                    "base64": image_base64,
                    "md5": ""  # 可以计算MD5，但通常不是必需的
                }
            }
            
            response = requests.post(self.webhook_url, json=data)
            result = response.json()
            
            if result.get('errcode') == 0:
                print("✅ 图片消息发送成功")
                return True
            else:
                print(f"❌ 图片消息发送失败: {result}")
                return False
                
        except Exception as e:
            print(f"❌ 发送图片消息失败: {e}")
            return False
    
    def send_comprehensive_message(self):
        """发送完整的监控报告到微信群"""
        try:
            # 读取Excel文件获取统计信息
            df = pd.read_excel('yxc.xlsx')
            
            # 创建统计信息
            total_stores = len(df)
            remaining_stats = df['剩余'].value_counts().to_dict()
            note1_stats = df['备注1'].value_counts().to_dict()
            note2_stats = df['备注2'].value_counts().to_dict()
            
            current_date = datetime.now().strftime('%Y年%m月%d日')
            
            # 创建文本消息
            text_message = f"""📊 店铺监控数据报告 - {current_date}

📈 数据统计：
• 总店铺数：{total_stores}个
• 剩余天数分布：{remaining_stats}
• 备注1分布：{note1_stats}
• 备注2分布：{note2_stats}

📋 表格包含以下列：
• 行号 - 显示每行的序号
• 店铺名称 - 店铺名称
• 地址 - 店铺地址
• 总天 - 总天数
• 剩余 - 剩余天数
• 开始时间 - 开始日期
• 备注1 - 大桶数量信息
• 备注2 - 状态信息

详细数据请查看下方图片表格。"""
            
            # 发送文本消息
            print("📤 发送文本统计信息...")
            self.send_text_message(text_message)
            
            # 创建并发送图片
            print("🖼️  创建表格图片...")
            image_buffer = self.create_table_image()
            if image_buffer:
                print("📤 发送表格图片...")
                self.send_image_message(image_buffer)
            
            print("✅ 完整报告发送完成")
            return True
            
        except Exception as e:
            print(f"❌ 发送完整报告失败: {e}")
            return False

def main():
    """主函数"""
    print("=== 微信群消息发送工具 ===")
    
    # 检查配置
    webhook_url = os.getenv('WECHAT_WEBHOOK_URL', '')
    if not webhook_url:
        print("❌ 请在.env文件中配置WECHAT_WEBHOOK_URL")
        print("📝 配置示例：")
        print("WECHAT_WEBHOOK_URL=https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=YOUR_KEY")
        print("\n🔧 如何获取webhook地址：")
        print("1. 在微信群中添加企业微信机器人")
        print("2. 获取机器人的webhook地址")
        print("3. 将地址添加到.env文件中")
        return
    
    sender = WeChatGroupSender()
    
    print("📤 发送完整监控报告到微信群...")
    success = sender.send_comprehensive_message()
    
    if success:
        print("\n🎉 微信群消息发送成功！")
        print("📱 请检查微信群，应该能看到统计信息和表格图片")
    else:
        print("\n❌ 微信群消息发送失败")

if __name__ == "__main__":
    main()
