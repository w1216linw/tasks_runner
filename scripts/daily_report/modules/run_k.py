"""Step 6: 揽收未达成分析(k)"""

import re
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pandas as pd

from utils.chinese_font import setup_chinese_font


def _extract_address_lib_number(addr):
    numbers = re.findall(r'\d+', str(addr))
    return numbers[0] if numbers else ''


def _extract_unit(addr):
    m = re.search(r'(APT|#|UNIT|SUITE|ROOM|FL|楼)\s*[\w\-]+', str(addr), re.IGNORECASE)
    return m.group(0).upper().strip() if m else ''


def _extract_main_address_number(address):
    address = str(address).lower()
    address = re.sub(r'il\s+\d{4,5}(\s+\d{4,5})?', '', address)
    numbers = re.findall(r'\d+', address)
    return str(max(map(int, numbers))) if numbers else ''


def run(today: datetime, data_dir: Path, output_dir: Path, log: Callable) -> dict:
    setup_chinese_font()
    plt.rcParams['axes.unicode_minus'] = False

    last2day = today - timedelta(days=2)
    last2day_str = last2day.strftime('%m-%d')
    k_date_str = last2day.strftime('%Y-%m-%d')

    module_dir = data_dir / 'input' / 'yy_unachieved'
    orders_path = module_dir / f'{last2day_str}未达成.xlsx'
    address_path = data_dir / '地址库.xlsx'

    log(f'读取: {orders_path.name}')
    k_orders = pd.read_excel(orders_path, engine='openpyxl')
    address_lib = pd.read_excel(address_path, engine='openpyxl')

    k_orders['门牌号'] = k_orders['发件人详细地址'].apply(_extract_main_address_number)
    address_lib['门牌号'] = address_lib['实际主地址'].apply(_extract_address_lib_number)
    k_seen = set(address_lib['门牌号'].dropna().astype(str))
    k_orders['是否触达'] = k_orders['门牌号'].astype(str).apply(lambda x: '已触达' if x in k_seen else '未触达')

    k_total = len(k_orders)
    k_reach = int(k_orders['是否触达'].value_counts().get('已触达', 0))
    log(f'未达成总数: {k_total}, 已触达: {k_reach}')

    k_orders['单元号'] = k_orders['发件人详细地址'].apply(_extract_unit)
    k_orders['地址单位'] = k_orders.apply(
        lambda row: f"{row['门牌号']} {row['单元号']}" if row['门牌号'] == '1111' and row['单元号'] else row['门牌号'],
        axis=1
    )
    k_unit = k_orders['地址单位'].value_counts().reset_index()
    k_unit.columns = ['地址单位', '数量']
    total = k_unit['数量'].sum()
    k_unit = k_unit.nlargest(10, '数量').sort_values(by='数量', ascending=True)

    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.barh(k_unit['地址单位'], k_unit['数量'])
    for bar, cnt in zip(bars, k_unit['数量']):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2,
                f'{cnt}单 ({cnt/total:.1%})', va='center', fontsize=9)
    ax.set_title(f'预约未达成前十地址分布 (仅1111细分UNIT) {k_date_str}', fontsize=14)
    ax.set_xlabel('订单数量')
    plt.tight_layout()
    k_image_bar = output_dir / f'预约未达成前十地址分布 (仅1111细分UNIT) {k_date_str}.png'
    fig.savefig(k_image_bar, dpi=150)
    plt.close(fig)

    log(f'图片已保存: {k_image_bar.name}')
    return {
        'k_date_str': k_date_str,
        'k_image_bar': str(k_image_bar),
    }
