#  conding:utf-8
"""
需要保证title里只能有一个商品名
不断完善
"""


def discount(title):
    brands = {'cpb': 0.75, 'GINZA': 0.85, '百优': 0.9, '悦薇珀翡': 0.9, '盼丽风姿': 0.9, '红腰子': 0.9, '时光琉璃': 0.9,
              '怡丽丝尔': 0.8, 'EFFECTIM': 1, '心机': 0.75, '黛珂': 0.7, '白檀': 0.7, 'sk2': 0.68,
              'ELEGANCE': 0.96, '澳尔滨': 0.9, 'ALBION': 0.9, 'TWANY': 0.67, 'ipsa': 0.9,
              'pola': 0.67, 'SUQQU': 0.9, 'THREE': 0.9, 'hacci': 0.9, 'HABA': 0.9,
              'fancl': 0.85, '芳珂': 0.85, '艾天然': 0.8, 'covermark': 0.9, 'paul': 0.9, '植村秀': 1, '赫莲娜': 0.9}
    for brand, dis in brands.items():
        if brand.lower() in title.lower():
            return dis
    # a = (dis for brand, dis in brands.items() if brand.lower() in title.lower())
    # return a


if __name__ == '__main__':
    pass
