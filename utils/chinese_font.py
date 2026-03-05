import platform
from matplotlib import rcParams


def setup_chinese_font() -> None:
    """配置 matplotlib 中文字体（自动适配 Windows/Mac/Linux）"""
    system = platform.system()
    if system == 'Windows':
        rcParams['font.family'] = 'SimHei'
    elif system == 'Darwin':
        rcParams['font.family'] = 'Arial Unicode MS'
    else:
        rcParams['font.family'] = 'WenQuanYi Micro Hei'
    rcParams['axes.unicode_minus'] = False
    rcParams['font.size'] = 12
