"""
颜色解析 - 支持中英文名称与十六进制
"""

# 颜色名称映射（中文 / 英文 → 十六进制）
COLOR_NAMES = {
    # 基础颜色
    '红色': '#FF0000',
    'red': '#FF0000',
    '绿色': '#00FF00',
    'green': '#00FF00',
    '蓝色': '#0000FF',
    'blue': '#0000FF',
    '黄色': '#FFFF00',
    'yellow': '#FFFF00',
    '橙色': '#FFA500',
    'orange': '#FFA500',
    '紫色': '#800080',
    'purple': '#800080',
    '粉色': '#FFC0CB',
    'pink': '#FFC0CB',
    '青色': '#00FFFF',
    'cyan': '#00FFFF',
    '白色': '#FFFFFF',
    'white': '#FFFFFF',
    '黑色': '#000000',
    'black': '#000000',
    '灰色': '#808080',
    'gray': '#808080',

    # 内控主题常用色（便于 override 时用中文表达）
    '浅蓝色': '#b4bff5',
    '淡蓝色': '#b4bff5',
    '深蓝色': '#2c3e50',
    '浅红': '#f8cecc',
    '浅黄': '#fff2cc',
    '浅绿': '#d5e8d4',
}


def parse_color(color_input: str) -> str:
    """
    解析颜色输入，返回标准十六进制格式。

    支持：
    - 中文：'红色'、'浅蓝色'
    - 英文：'red'、'blue'
    - 十六进制：'#FF0000'（原样返回，转大写）

    无法识别时返回黑色并告警。
    """
    if not color_input:
        return '#000000'

    color_str = str(color_input).strip()

    # 已经是十六进制，直接返回
    if color_str.startswith('#'):
        return color_str.upper()

    # 查找中英文名称映射
    color_key = color_str.lower()
    for name, hex_value in COLOR_NAMES.items():
        if name.lower() == color_key:
            return hex_value.upper()

    print(f"⚠️ 未识别的颜色: {color_str}，使用默认黑色")
    return '#000000'


if __name__ == '__main__':
    # 简单自测
    for c in ['红色', 'red', '#FF0000', '浅蓝色', '橙色', '不存在色']:
        print(f"parse_color({c!r}) = {parse_color(c)}")
