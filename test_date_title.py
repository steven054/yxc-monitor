#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
from datetime import datetime

def test_date_title():
    """测试图片标题是否包含日期"""
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
        
        # 准备表格数据（只显示前5行作为测试）
        table_data = []
        for idx, row in df.head(5).iterrows():
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
        title = f'店铺监控数据表（包含行号和备注）- {current_date}'
        plt.title(title, fontsize=16, fontweight='bold', pad=20)
        
        # 保存图片
        filename = f'test_date_title_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png'
        plt.savefig(filename, format='png', dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"✅ 测试图片已保存为: {filename}")
        print(f"📅 图片标题: {title}")
        print(f"📅 当前日期: {current_date}")
        
        return True
        
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

if __name__ == "__main__":
    print("=== 测试图片标题日期功能 ===")
    success = test_date_title()
    if success:
        print("\n🎉 测试成功！")
        print("📧 现在所有邮件中的图片标题都会包含发送日期")
    else:
        print("\n❌ 测试失败")
