# 字体配置代码
import matplotlib.pyplot as plt
import platform

def setup_chinese_font():
    """设置中文字体"""
    system = platform.system()
    
    if system == 'Darwin':  # macOS
        font_list = ['PingFang SC', 'Hiragino Sans GB', 'STHeiti', 'SimHei', 'Arial Unicode MS']
    elif system == 'Windows':  # Windows
        font_list = ['PingFang SC', 'Hiragino Sans GB', 'STHeiti', 'SimHei', 'Arial Unicode MS']
    else:  # Linux
        font_list = ['PingFang SC', 'Hiragino Sans GB', 'STHeiti', 'SimHei', 'Arial Unicode MS']
    
    # 尝试设置字体
    for font in font_list:
        try:
            plt.rcParams['font.sans-serif'] = [font]
            plt.rcParams['axes.unicode_minus'] = False
            print(f"✅ 使用字体: {font}")
            return font
        except:
            continue
    
    print("⚠️ 未找到合适的中文字体")
    return None

# 使用示例
setup_chinese_font()
