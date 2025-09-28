#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from datetime import datetime
import os

def generate_wechat_image():
    """生成适合微信群发送的图片"""
    try:
        # 读取Excel文件
        df = pd.read_excel('yxc.xlsx')
        print(f"✅ 成功读取Excel文件，共 {len(df)} 行数据")
        
        # 设置中文字体
        plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS', 'DejaVu Sans']
        plt.rcParams['axes.unicode_minus'] = False
        
        # 创建图形（适合手机屏幕的尺寸）
        fig, ax = plt.subplots(figsize=(12, 16))
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
        table.set_fontsize(7)
        table.scale(1, 1.2)
        
        # 设置标题（包含发送日期）
        current_date = datetime.now().strftime('%Y年%m月%d日')
        plt.title(f'店铺监控数据表 - {current_date}', fontsize=14, fontweight='bold', pad=20)
        
        # 添加统计信息
        total_stores = len(df)
        remaining_stats = df['剩余'].value_counts().to_dict()
        note1_stats = df['备注1'].value_counts().to_dict()
        note2_stats = df['备注2'].value_counts().to_dict()
        
        stats_text = f"""
数据统计：
• 总店铺数：{total_stores}个
• 剩余天数分布：{remaining_stats}
• 备注1分布：{note1_stats}
• 备注2分布：{note2_stats}
        """
        
        # 在图片底部添加统计信息
        plt.figtext(0.5, 0.02, stats_text, ha='center', fontsize=10, 
                   bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray", alpha=0.8))
        
        # 保存图片
        filename = f'微信群_店铺监控表_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png'
        plt.savefig(filename, format='png', dpi=200, bbox_inches='tight', 
                   facecolor='white', edgecolor='none')
        plt.close()
        
        print(f"✅ 微信群图片已生成: {filename}")
        print(f"📱 图片尺寸: 12x16英寸，适合手机查看")
        print(f"📊 包含完整的行号和备注信息")
        print(f"📈 包含数据统计信息")
        
        # 显示文件大小
        file_size = os.path.getsize(filename) / 1024 / 1024  # MB
        print(f"📦 文件大小: {file_size:.2f} MB")
        
        if file_size > 2:
            print("⚠️  文件较大，建议压缩后发送")
        else:
            print("✅ 文件大小适合微信发送")
        
        return filename
        
    except Exception as e:
        print(f"❌ 生成图片失败: {e}")
        return None

def main():
    """主函数"""
    print("=== 生成微信群图片 ===")
    
    filename = generate_wechat_image()
    
    if filename:
        print(f"\n🎉 图片生成成功！")
        print(f"📁 文件位置: {os.path.abspath(filename)}")
        print(f"\n📱 发送到微信群的步骤：")
        print(f"1. 打开微信群聊")
        print(f"2. 点击输入框左侧的'+'号")
        print(f"3. 选择'相册'")
        print(f"4. 选择刚生成的图片文件: {filename}")
        print(f"5. 点击发送")
        print(f"\n💡 提示：图片包含完整的行号和备注信息，以及数据统计")
    else:
        print(f"\n❌ 图片生成失败")

if __name__ == "__main__":
    main()
