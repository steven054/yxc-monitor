#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import platform
import os

def list_available_fonts():
    """列出所有可用的中文字体"""
    print("🔍 检查系统中可用的中文字体...")
    
    # 获取所有字体
    fonts = [f.name for f in fm.fontManager.ttflist]
    
    # 中文字体关键词
    chinese_keywords = ['SimHei', 'Microsoft YaHei', 'PingFang', 'Hiragino', 'STHeiti', 
                       'WenQuanYi', 'Noto Sans CJK', 'Source Han', 'Droid Sans Fallback']
    
    chinese_fonts = []
    for font in fonts:
        for keyword in chinese_keywords:
            if keyword.lower() in font.lower():
                chinese_fonts.append(font)
                break
    
    print(f"找到 {len(chinese_fonts)} 个中文字体:")
    for i, font in enumerate(chinese_fonts, 1):
        print(f"  {i}. {font}")
    
    return chinese_fonts

def test_font(font_name):
    """测试字体是否可用"""
    try:
        plt.rcParams['font.sans-serif'] = [font_name]
        plt.rcParams['axes.unicode_minus'] = False
        
        # 创建测试图形
        fig, ax = plt.subplots(figsize=(8, 4))
        ax.text(0.5, 0.5, f'测试中文显示: {font_name}', 
                fontsize=16, ha='center', va='center')
        ax.set_title(f'字体测试: {font_name}')
        ax.set_xlim(0, 1)
        ax.set_ylim(0, 1)
        
        # 保存测试图片
        test_filename = f'font_test_{font_name.replace(" ", "_")}.png'
        plt.savefig(test_filename, dpi=150, bbox_inches='tight')
        plt.close()
        
        print(f"✅ {font_name} - 测试成功，已保存: {test_filename}")
        return True
        
    except Exception as e:
        print(f"❌ {font_name} - 测试失败: {str(e)}")
        return False

def install_fonts_macos():
    """macOS字体安装建议"""
    print("\n🍎 macOS字体安装建议:")
    print("1. 打开字体册应用")
    print("2. 下载并安装以下字体:")
    print("   - PingFang SC (系统自带)")
    print("   - Hiragino Sans GB (系统自带)")
    print("   - STHeiti (系统自带)")
    print("3. 或者从网上下载 SimHei.ttf 字体文件")

def install_fonts_linux():
    """Linux字体安装建议"""
    print("\n🐧 Linux字体安装建议:")
    print("Ubuntu/Debian:")
    print("  sudo apt-get install fonts-wqy-microhei fonts-wqy-zenhei")
    print("  sudo apt-get install fonts-noto-cjk")
    print("\nCentOS/RHEL:")
    print("  sudo yum install wqy-microhei-fonts wqy-zenhei-fonts")

def create_font_config():
    """创建字体配置文件"""
    system = platform.system()
    
    if system == 'Darwin':  # macOS
        font_list = ['PingFang SC', 'Hiragino Sans GB', 'STHeiti', 'SimHei', 'Arial Unicode MS']
    elif system == 'Windows':  # Windows
        font_list = ['SimHei', 'Microsoft YaHei', 'KaiTi', 'FangSong']
    else:  # Linux
        font_list = ['WenQuanYi Micro Hei', 'WenQuanYi Zen Hei', 'Noto Sans CJK SC', 'SimHei']
    
    config_code = f'''# 字体配置代码
import matplotlib.pyplot as plt
import platform

def setup_chinese_font():
    """设置中文字体"""
    system = platform.system()
    
    if system == 'Darwin':  # macOS
        font_list = {font_list}
    elif system == 'Windows':  # Windows
        font_list = {font_list}
    else:  # Linux
        font_list = {font_list}
    
    # 尝试设置字体
    for font in font_list:
        try:
            plt.rcParams['font.sans-serif'] = [font]
            plt.rcParams['axes.unicode_minus'] = False
            print(f"✅ 使用字体: {{font}}")
            return font
        except:
            continue
    
    print("⚠️ 未找到合适的中文字体")
    return None

# 使用示例
setup_chinese_font()
'''
    
    with open('font_config.py', 'w', encoding='utf-8') as f:
        f.write(config_code)
    
    print(f"\n📝 已创建字体配置文件: font_config.py")

def main():
    print("🔧 中文字体问题诊断和修复工具")
    print("=" * 50)
    
    # 检查系统
    system = platform.system()
    print(f"🖥️  当前系统: {system}")
    
    # 列出可用字体
    chinese_fonts = list_available_fonts()
    
    if not chinese_fonts:
        print("\n❌ 未找到中文字体!")
        if system == 'Darwin':
            install_fonts_macos()
        elif system == 'Linux':
            install_fonts_linux()
        else:
            print("请手动安装中文字体")
        return
    
    # 测试前几个字体
    print(f"\n🧪 测试前3个字体...")
    working_fonts = []
    for font in chinese_fonts[:3]:
        if test_font(font):
            working_fonts.append(font)
    
    if working_fonts:
        print(f"\n✅ 找到 {len(working_fonts)} 个可用字体:")
        for font in working_fonts:
            print(f"  - {font}")
        
        # 创建配置文件
        create_font_config()
        
        print(f"\n🎉 建议在代码中使用: {working_fonts[0]}")
    else:
        print("\n❌ 所有测试的字体都不可用")
        print("请检查字体安装或尝试其他字体")

if __name__ == "__main__":
    main()
