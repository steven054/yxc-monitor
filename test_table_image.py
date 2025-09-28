#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # 使用非交互式后端

def create_table_image():
    """创建包含行号和备注的表格图片"""
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
        
        # 设置标题
        plt.title('店铺监控数据表（包含行号和备注）', fontsize=16, fontweight='bold', pad=20)
        
        # 保存图片
        plt.savefig('table_with_notes.png', format='png', dpi=300, bbox_inches='tight')
        plt.close()
        
        print("✅ 表格图片已保存为: table_with_notes.png")
        print("📊 图片包含以下列:")
        print("   - 行号")
        print("   - 店铺名称") 
        print("   - 地址")
        print("   - 总天")
        print("   - 剩余")
        print("   - 开始时间")
        print("   - 备注1")
        print("   - 备注2")
        
        return True
        
    except Exception as e:
        print(f"❌ 创建表格图片失败: {e}")
        return False

if __name__ == "__main__":
    print("=== 创建表格图片测试 ===")
    success = create_table_image()
    if success:
        print("\n🎉 表格图片创建成功！")
        print("📧 邮件中的表格图片现在包含了行号和备注列")
    else:
        print("\n❌ 表格图片创建失败")
