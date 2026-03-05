"""Step 3: 已下单地址触达(d)"""

import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pandas as pd

from utils.chinese_font import setup_chinese_font


def _extract_true_address_number(addr):
    addr = str(addr)
    addr = re.sub(r'\b\d{4,5}\b$', '', addr)
    numbers = re.findall(r'\d+', addr)
    return str(max(map(int, numbers))) if numbers else ''


def _extract_address_lib_number(addr):
    numbers = re.findall(r'\d+', str(addr))
    return numbers[0] if numbers else ''


def _extract_unit(addr):
    m = re.search(r'(APT|#|UNIT|SUITE|ROOM|FL|楼)\s*[\w\-]+', str(addr), re.IGNORECASE)
    return m.group(0).upper().strip() if m else ''


def run(today: datetime, data_dir: Path, output_dir: Path, log: Callable) -> dict:
    setup_chinese_font()
    plt.rcParams['axes.unicode_minus'] = False

    yesterday = today - timedelta(days=1)
    yesterday_str = yesterday.strftime('%m-%d')

    module_dir = data_dir / 'input' / 'tubt_pending_orders'
    orders_path = module_dir / f'{yesterday_str}已下单.xlsx'
    address_path = data_dir / '地址库.xlsx'

    log(f'读取: {orders_path.name}')
    orders = pd.read_excel(orders_path, engine='openpyxl')
    address_lib = pd.read_excel(address_path, engine='openpyxl')

    m = re.search(r'(\d{1,2})\-(\d{1,2})', orders_path.name)
    d_date_str = f'2026-{int(m.group(1)):02d}-{int(m.group(2)):02d}' if m else str(yesterday.date())

    orders['门牌号'] = orders['发件人详细地址'].apply(_extract_true_address_number)
    address_lib['门牌号'] = address_lib['实际主地址'].apply(_extract_address_lib_number)
    seen = set(address_lib['门牌号'])
    orders['是否触达'] = orders['门牌号'].apply(lambda x: '已触达' if x in seen else '未触达')

    count = orders['是否触达'].value_counts()
    total_orders = len(orders)
    reach_rate = round(count.get('已触达', 0) / total_orders * 100, 1) if total_orders > 0 else 0
    log(f'已下单总数: {total_orders}, 触达率: {reach_rate}%')

    # 图一：触达饼图
    fig1, ax1 = plt.subplots(figsize=(4, 4))
    ax1.pie(count.values, labels=count.index, autopct='%1.1f%%',
            colors=['#4CAF50', '#F44336'][:len(count)], startangle=140,
            explode=[0.05] + [0] * (len(count) - 1))
    ax1.set_title(f'订单地址触达情况分析（{d_date_str}）', fontsize=14)
    ax1.axis('equal')
    plt.tight_layout()
    d_image_pie = output_dir / f'订单地址触达情况分析（{d_date_str}）.png'
    fig1.savefig(d_image_pie, dpi=150)
    plt.close(fig1)

    # 图二：前十地址柱状图
    orders['单元号'] = orders['发件人详细地址'].apply(_extract_unit)
    orders['地址单位'] = orders.apply(
        lambda row: f"{row['门牌号']} {row['单元号']}" if row['门牌号'] == '1111' and row['单元号'] else row['门牌号'],
        axis=1
    )
    unit_count = orders['地址单位'].value_counts().reset_index()
    unit_count.columns = ['地址单位', '数量']
    total = unit_count['数量'].sum()
    unit_count = unit_count.nlargest(10, '数量').sort_values(by='数量', ascending=True)

    fig2, ax2 = plt.subplots(figsize=(10, 6))
    bars = ax2.barh(unit_count['地址单位'], unit_count['数量'], color='cornflowerblue')
    for bar, cnt in zip(bars, unit_count['数量']):
        ax2.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2,
                 f'{cnt}单 ({cnt/total:.1%})', va='center', fontsize=9)
    ax2.set_title(f'已下单前十地址单位分布（仅1111细分UNIT） {d_date_str}', fontsize=14)
    ax2.set_xlabel('订单数量')
    plt.tight_layout()
    d_image_bar = output_dir / f'已下单前十地址单位分布（仅1111细分UNIT） {d_date_str}.png'
    fig2.savefig(d_image_bar, dpi=150)
    plt.close(fig2)

    log('图片已保存')
    return {
        'orders': total_orders,
        'reach_rate': reach_rate,
        'd_date_str': d_date_str,
        'd_image_pie': str(d_image_pie),
        'd_image_bar': str(d_image_bar),
    }
