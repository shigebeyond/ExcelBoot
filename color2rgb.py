#!/usr/bin/python3
# -*- coding: utf-8 -*-
from openpyxl.styles import colors

# 颜色argb选取: http://t.zoukankan.com/jytblog-p-8134744.html

# 颜色代码对rgb值的映射
code2rgb = {
    'white': 'FFFFFF', # 白色
    'ivory': 'FFFFF0', # 象牙色
    'lightyellow': 'FFFFE0', # 亮黄色
    'yellow': 'FFFF00', # 黄色
    'snow': 'FFFAFA', # 雪白色
    'floralwhite': 'FFFAF0', # 花白色
    'lemonchiffon': 'FFFACD', # 柠檬绸色
    'cornsilk': 'FFF8DC', # 米绸色
    'seashell': 'FFF5EE', # 海贝色
    'lavenderblush': 'FFF0F5', # 淡紫红
    'papayawhip': 'FFEFD5', # 番木色
    'blanchedalmond': 'FFEBCD', # 白杏色
    'mistyrose': 'FFE4E1', # 浅玫瑰色
    'bisque': 'FFE4C4', # 桔黄色
    'moccasin': 'FFE4B5', # 鹿皮色
    'navajowhite': 'FFDEAD', # 纳瓦白
    'peachpuff': 'FFDAB9', # 桃色
    'gold': 'FFD700', # 金色
    'pink': 'FFC0CB', # 粉红色
    'lightpink': 'FFB6C1', # 亮粉红色
    'orange': 'FFA500', # 橙色
    'lightsalmon': 'FFA07A', # 亮肉色
    'darkorange': 'FF8C00', # 暗桔黄色
    'coral': 'FF7F50', # 珊瑚色
    'hotpink': 'FF69B4', # 热粉红色
    'tomato': 'FF6347', # 西红柿色
    'orangered': 'FF4500', # 红橙色
    'deeppink': 'FF1493', # 深粉红色
    'fuchsia': 'FF00FF', # 紫红色
    'magenta': 'FF00FF', # 红紫色
    'red': 'FF0000', # 红色
    'oldlace': 'FDF5E6', # 老花色
    'lightgoldenrodyellow': 'FAFAD2', # 亮金黄色
    'linen': 'FAF0E6', # 亚麻色
    'antiquewhite': 'FAEBD7', # 古董白
    'salmon': 'FA8072', # 鲜肉色
    'ghostwhite': 'F8F8FF', # 幽灵白
    'mintcream': 'F5FFFA', # 薄荷色
    'whitesmoke': 'F5F5F5', # 烟白色
    'beige': 'F5F5DC', # 米色
    'wheat': 'F5DEB3', # 浅黄色
    'sandybrown': 'F4A460', # 沙褐色
    'azure': 'F0FFFF', # 天蓝色
    'honeydew': 'F0FFF0', # 蜜色
    'aliceblue': 'F0F8FF', # 艾利斯兰
    'khaki': 'F0E68C', # 黄褐色
    'lightcoral': 'F08080', # 亮珊瑚色
    'palegoldenrod': 'EEE8AA', # 苍麒麟色
    'violet': 'EE82EE', # 紫罗兰色
    'darksalmon': 'E9967A', # 暗肉色
    'lavender': 'E6E6FA', # 淡紫色
    'lightcyan': 'E0FFFF', # 亮青色
    'burlywood': 'DEB887', # 实木色
    'plum': 'DDA0DD', # 洋李色
    'gainsboro': 'DCDCDC', # 淡灰色
    'crimson': 'DC143C', # 暗深红色
    'palevioletred': 'DB7093', # 苍紫罗兰色
    'goldenrod': 'DAA520', # 金麒麟色
    'orchid': 'DA70D6', # 淡紫色
    'thistle': 'D8BFD8', # 蓟色
    'lightgray': 'D3D3D3', # 亮灰色
    'lightgrey': 'D3D3D3', # 亮灰色
    'tan': 'D2B48C', # 茶色
    'chocolate': 'D2691E', # 巧可力色
    'peru': 'CD853F', # 秘鲁色
    'indianred': 'CD5C5C', # 印第安红
    'mediumvioletred': 'C71585', # 中紫罗兰色
    'silver': 'C0C0C0', # 银色
    'darkkhaki': 'BDB76B', # 暗黄褐色
    'rosybrown': 'BC8F8F', # 褐玫瑰红
    'mediumorchid': 'BA55D3', # 中粉紫色
    'darkgoldenrod': 'B8860B', # 暗金黄色
    'firebrick': 'B22222', # 火砖色
    'powderblue': 'B0E0E6', # 粉蓝色
    'lightsteelblue': 'B0C4DE', # 亮钢兰色
    'paleturquoise': 'AFEEEE', # 苍宝石绿
    'greenyellow': 'ADFF2F', # 黄绿色
    'lightblue': 'ADD8E6', # 亮蓝色
    'darkgray': 'A9A9A9', # 暗灰色
    'darkgrey': 'A9A9A9', # 暗灰色
    'brown': 'A52A2A', # 褐色
    'sienna': 'A0522D', # 赭色
    'darkorchid': '9932CC', # 暗紫色
    'palegreen': '98FB98', # 苍绿色
    'darkviolet': '9400D3', # 暗紫罗兰色
    'mediumpurple': '9370DB', # 中紫色
    'lightgreen': '90EE90', # 亮绿色
    'darkseagreen': '8FBC8F', # 暗海兰色
    'saddlebrown': '8B4513', # 重褐色
    'darkmagenta': '8B008B', # 暗洋红
    'darkred': '8B0000', # 暗红色
    'blueviolet': '8A2BE2', # 紫罗兰蓝色
    'lightskyblue': '87CEFA', # 亮天蓝色
    'skyblue': '87CEEB', # 天蓝色
    'gray': '808080', # 灰色
    'grey': '808080', # 灰色
    'olive': '808000', # 橄榄色
    'purple': '800080', # 紫色
    'maroon': '800000', # 粟色
    'aquamarine': '7FFFD4', # 碧绿色
    'chartreuse': '7FFF00', # 黄绿色
    'lawngreen': '7CFC00', # 草绿色
    'mediumslateblue': '7B68EE', # 中暗蓝色
    'lightslategray': '778899', # 亮蓝灰
    'lightslategrey': '778899', # 亮蓝灰
    'slategray': '708090', # 灰石色
    'slategrey': '708090', # 灰石色
    'olivedrab': '6B8E23', # 深绿褐色
    'slateblue': '6A5ACD', # 石蓝色
    'dimgray': '696969', # 暗灰色
    'dimgrey': '696969', # 暗灰色
    'mediumaquamarine': '66CDAA', # 中绿色
    'cornflowerblue': '6495ED', # 菊兰色
    'cadetblue': '5F9EA0', # 军兰色
    'darkolivegreen': '556B2F', # 暗橄榄绿
    'indigo': '4B0082', # 靛青色
    'mediumturquoise': '48D1CC', # 中绿宝石
    'darkslateblue': '483D8B', # 暗灰蓝色
    'steelblue': '4682B4', # 钢兰色
    'royalblue': '4169E1', # 皇家蓝
    'turquoise': '40E0D0', # 青绿色
    'mediumseagreen': '3CB371', # 中海蓝
    'limegreen': '32CD32', # 橙绿色
    'darkslategray': '2F4F4F', # 暗瓦灰色
    'darkslategrey': '2F4F4F', # 暗瓦灰色
    'seagreen': '2E8B57', # 海绿色
    'forestgreen': '228B22', # 森林绿
    'lightseagreen': '20B2AA', # 亮海蓝色
    'dodgerblue': '1E90FF', # 闪兰色
    'midnightblue': '191970', # 中灰兰色
    'aqua': '00FFFF', # 浅绿色
    'cyan': '00FFFF', # 青色
    'springgreen': '00FF7F', # 春绿色
    'lime': '00FF00', # 酸橙色
    'mediumspringgreen': '00FA9A', # 中春绿色
    'darkturquoise': '00CED1', # 暗宝石绿
    'deepskyblue': '00BFFF', # 深天蓝色
    'darkcyan': '008B8B', # 暗青色
    'teal': '008080', # 水鸭色
    'green': '008000', # 绿色
    'darkgreen': '006400', # 暗绿色
    'blue': '0000FF', # 蓝色
    'mediumblue': '0000CD', # 中兰色
    'darkblue': '00008B', # 暗蓝色
    'navy': '000080', # 海军色
    'black': '000000', # 黑色
}

# 获得颜色代码对应的rgb值
def get_rgb(code):
    # rgb
    m = colors.aRGB_REGEX.match(code)
    if m != None:
        return code

    # 纯数字
    if code.isdigit():
        i = int(code)
        if i < len(colors):
            return colors[i]

    # 纯字母
    if code.isalpha():
        if code in code2rgb:
            return code2rgb[code]

    raise Exception("Invalid color code")

