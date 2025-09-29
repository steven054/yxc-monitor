#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
import json
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
import io
import base64
import hashlib
from datetime import datetime
import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

class WeChatImageSender:
    def __init__(self):
        # 企业微信机器人webhook地址
        self.webhook_url = os.getenv('WECHAT_WEBHOOK_URL', '')
        
    def create_table_image(self):
        """创建表格图片"""
        try:
            # 读取Excel文件
            df = pd.read_excel('yxc.xlsx')
            print(f"✅ 成功读取Excel文件，共 {len(df)} 行数据")
            
            # 设置中文字体 - 支持本地和GitHub Actions环境
            import platform
            import os
            
            # 检查是否在GitHub Actions环境中
            is_github_actions = os.getenv('GITHUB_ACTIONS') == 'true'
            system = platform.system()
            
            if is_github_actions or system == 'Linux':
                # GitHub Actions或Linux环境 - 使用安装的中文字体
                font_list = ['WenQuanYi Micro Hei', 'WenQuanYi Zen Hei', 'Noto Sans CJK SC', 'DejaVu Sans']
                print("🔧 检测到GitHub Actions/Linux环境，使用中文字体")
            elif system == 'Darwin':  # macOS
                font_list = ['STHeiti', 'Hiragino Sans GB', 'PingFang SC', 'Arial Unicode MS']
                print("🍎 检测到macOS环境，使用系统字体")
            elif system == 'Windows':  # Windows
                font_list = ['SimHei', 'Microsoft YaHei', 'KaiTi', 'FangSong']
                print("🪟 检测到Windows环境，使用系统字体")
            else:
                font_list = ['DejaVu Sans']
                print("⚠️ 未知环境，使用默认字体")
            
            # 设置字体
            plt.rcParams['font.sans-serif'] = font_list
            plt.rcParams['axes.unicode_minus'] = False
            
            # 在GitHub Actions中重建字体缓存
            if is_github_actions:
                try:
                    import matplotlib.font_manager as fm
                    fm._rebuild()
                    print("✅ 已重建字体缓存")
                except:
                    print("⚠️ 字体缓存重建失败，但继续执行")
            
            print(f"✅ 已设置字体: {font_list[0]}")
            
            # 创建图形 - 增加高度以适应更高的行高
            fig, ax = plt.subplots(figsize=(16, 15))
            ax.axis('tight')
            ax.axis('off')
            
            # 准备表格数据
            table_data = []
            for idx, row in df.iterrows():
                # 处理备注1和备注2的空值显示
                note1 = row.get('备注1', '')
                note2 = row.get('备注2', '')
                if pd.isna(note1) or note1 == '':
                    note1 = ''
                if pd.isna(note2) or note2 == '':
                    note2 = ''
                
                # 处理备注3的智能显示
                remaining_days = int(row.get('剩余', 0))
                if remaining_days < 3:
                    note3 = '即将到期'
                elif remaining_days < 7:
                    note3 = '1周内到期'
                else:
                    note3 = '正常'
                
                table_data.append([
                    str(row.get('行号', '')),
                    row.get(' 店铺名称', ''),
                    row.get('地址', ''),
                    str(row.get('总天', '')),
                    str(row.get('剩余', '')),
                    str(row.get('开始时间', '')),
                    str(note1),
                    str(note2),
                    str(note3)
                ])
            
            # 创建表格
            table = ax.table(
                cellText=table_data,
                colLabels=['行号', '店铺名称', '地址', '总天', '剩余', '开始时间', '备注1', '备注2', '备注3'],
                cellLoc='center',
                loc='center',
                bbox=[0, 0, 1, 1]
            )
            
            # 设置表格样式
            table.auto_set_font_size(False)
            table.set_fontsize(8)  # 稍微减小字体避免重叠
            table.scale(1, 3.0)  # 进一步增加行高到3.0
            
            # 设置表格边框和字体颜色
            for i in range(len(table_data) + 1):  # +1 for header
                for j in range(9):  # 总共9列
                    if i == 0:  # 表头
                        table[(i, j)].set_facecolor('#f0f0f0')  # 浅灰色背景
                        table[(i, j)].set_text_props(weight='bold', color='#333333')  # 深灰色粗体
                    else:  # 数据行
                        table[(i, j)].set_text_props(color='#000000')  # 黑色字体
            
            # 高亮剩余天数小于等于1的行
            for i in range(1, len(table_data) + 1):
                remaining_days = table_data[i-1][4]  # 剩余天数列是第5列（索引4）
                if remaining_days in ['0', '1']:  # 剩余天数为0或1
                    for j in range(9):  # 总共9列
                        table[(i, j)].set_facecolor('#ff9999')  # 更鲜艳的红色背景
                        table[(i, j)].set_text_props(weight='bold', color='#990000')  # 更深的红色粗体字体
            
            # 设置标题（包含发送日期）
            current_date = datetime.now().strftime('%Y年%m月%d日')
            plt.title(f'店铺监控数据表 - {current_date}', fontsize=16, fontweight='bold', pad=20)
            
            # 保存图片到内存
            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=150, bbox_inches='tight', 
                       facecolor='white', edgecolor='none')
            img_buffer.seek(0)
            
            plt.close()
            
            print("✅ 表格图片创建成功")
            return img_buffer
            
        except Exception as e:
            print(f"❌ 创建表格图片失败: {e}")
            return None
    
    def send_text_message(self, message):
        """发送文本消息"""
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
        """发送图片消息（修复版本）"""
        if not self.webhook_url:
            print("❌ 未配置企业微信机器人webhook地址")
            return False
            
        try:
            # 将图片转换为base64
            image_data = image_buffer.getvalue()
            image_base64 = base64.b64encode(image_data).decode('utf-8')
            
            # 计算MD5
            md5_hash = hashlib.md5(image_data).hexdigest()
            
            data = {
                "msgtype": "image",
                "image": {
                    "base64": image_base64,
                    "md5": md5_hash
                }
            }
            
            print(f"📊 图片信息: 大小={len(image_data)}字节, MD5={md5_hash[:8]}...")
            
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
    
    def calculate_remaining_days(self, df, columns):
        """基于实际日期计算剩余天数"""
        updated_count = 0
        current_date = datetime.now()
        
        for idx, row in df.iterrows():
            start_date_str = str(row[columns['start_date']])
            total_days = int(row[columns['total']])
            old_remaining = int(row[columns['remaining']])
            
            try:
                # 解析开始时间
                start_date = datetime.strptime(start_date_str, '%Y%m%d')
                
                # 计算已过去的天数
                days_passed = (current_date - start_date).days
                
                # 计算正确的剩余天数
                correct_remaining = max(0, total_days - days_passed)
                
                # 更新剩余天数
                df.at[idx, columns['remaining']] = correct_remaining
                
                row_num = row.get('行号', idx+1)
                store_name = row.get(' 店铺名称', f'行{row_num}')
                print(f"📊 {store_name}: 开始{start_date_str}, 总{total_days}天, 已过{days_passed}天, {old_remaining}→{correct_remaining}")
                
                if old_remaining != correct_remaining:
                    updated_count += 1
                    
            except Exception as e:
                print(f"❌ 处理行{idx+1}失败: {e}")
        
        return updated_count
    
    def reset_expired_items(self, df, columns):
        """重置剩余天数为0的项目"""
        reset_count = 0
        current_date = datetime.now()
        
        for idx, row in df.iterrows():
            remaining = row[columns['remaining']]
            total = row[columns['total']]
            start_date = row[columns['start_date']]
            
            # 如果剩余天数为0，需要重置
            if pd.notna(remaining) and int(remaining) == 0:
                row_num = row.get('行号', idx+1)
                store_name = row.get(' 店铺名称', f'行{row_num}')
                print(f"🔄 重置项目: {store_name}")
                
                # 更新开始时间为今天
                new_start_date = current_date.strftime('%Y%m%d')
                print(f"   开始时间: {start_date} → {new_start_date}")
                print(f"   剩余天数: 0 → {total}")
                
                # 确保日期格式正确，保持与原始数据类型一致
                if pd.api.types.is_integer_dtype(df[columns['start_date']]):
                    df.at[idx, columns['start_date']] = int(new_start_date)
                else:
                    df.at[idx, columns['start_date']] = str(new_start_date)
                
                # 重置剩余天数为总天数
                df.at[idx, columns['remaining']] = total
                
                reset_count += 1
        
        return reset_count

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

    def send_comprehensive_report(self):
        """发送完整的监控报告"""
        try:
            # 读取Excel文件
            df = pd.read_excel('yxc.xlsx')
            
            # 查找关键列
            columns = self.find_columns(df)
            
            # 基于实际日期计算剩余天数
            print("📅 计算剩余天数...")
            updated_count = self.calculate_remaining_days(df, columns)
            print(f"✅ 更新了 {updated_count} 个项目的剩余天数")
            
            # 重置剩余天数为0的项目
            print("🔄 重置过期项目...")
            reset_count = self.reset_expired_items(df, columns)
            print(f"✅ 重置了 {reset_count} 个过期项目")
            
            # 保存更新后的Excel文件
            df.to_excel('yxc.xlsx', index=False)
            print("💾 已保存更新后的Excel文件")
            
            # 创建统计信息
            total_stores = len(df)
            remaining_stats = df['剩余'].value_counts().to_dict()
            
            # 处理备注1和备注2的统计（过滤空值）
            note1_stats = df['备注1'].dropna().value_counts().to_dict()
            note2_stats = df['备注2'].dropna().value_counts().to_dict()
            
            # 计算备注3的统计（基于剩余天数）
            note3_stats = {'即将到期': 0, '1周内到期': 0, '正常': 0}
            for idx, row in df.iterrows():
                remaining_days = int(row.get('剩余', 0))
                if remaining_days < 3:
                    note3_stats['即将到期'] += 1
                elif remaining_days < 7:
                    note3_stats['1周内到期'] += 1
                else:
                    note3_stats['正常'] += 1
            
            current_date = datetime.now().strftime('%Y年%m月%d日 %H:%M')
            
            # 创建文本消息
            text_message = f"""📊 店铺监控数据报告 - {current_date}

📈 数据统计：
• 总店铺数：{total_stores}个
• 剩余天数分布：{remaining_stats}
• 备注1分布：{note1_stats}
• 备注2分布：{note2_stats}
• 备注3分布：{note3_stats}

📋 表格包含：行号、店铺名称、地址、总天、剩余、开始时间、备注1、备注2、备注3

详细数据请查看下方图片表格。"""
            
            # 发送文本消息
            print("📤 发送统计信息...")
            text_success = self.send_text_message(text_message)
            
            # 创建并发送图片
            print("🖼️  创建表格图片...")
            image_buffer = self.create_table_image()
            if image_buffer:
                print("📤 发送表格图片...")
                image_success = self.send_image_message(image_buffer)
            else:
                image_success = False
            
            return text_success and image_success
            
        except Exception as e:
            print(f"❌ 发送完整报告失败: {e}")
            return False

def main():
    """主函数"""
    print("=== 企业微信群图片发送工具（修复版） ===")
    
    # 检查配置
    webhook_url = os.getenv('WECHAT_WEBHOOK_URL', '')
    if not webhook_url:
        print("❌ 请在.env文件中配置WECHAT_WEBHOOK_URL")
        return
    
    sender = WeChatImageSender()
    
    print("📤 自动发送监控报告到微信群...")
    success = sender.send_comprehensive_report()
    
    if success:
        print("\n🎉 微信群自动发送成功！")
        print("📱 请检查微信群，应该能看到统计信息和表格图片")
    else:
        print("\n❌ 微信群自动发送失败")

if __name__ == "__main__":
    main()
