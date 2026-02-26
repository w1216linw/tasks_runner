"""
每日订单量扫描脚本 — 整合自 func/weekly_order_scan/count_orders.py
"""

from __future__ import annotations

from datetime import timedelta
from pathlib import Path
from typing import Callable

import pandas as pd

LogFn = Callable[[str], None]

CUTOFF_HOUR = 8  # 每天时间窗口: 前一天08:00 ~ 当天08:00


def _get_day_label(dt) -> str:
    """08:00 截止规则：当天08:00前归当天，之后归次日。"""
    if dt.hour >= CUTOFF_HOUR:
        return (dt + timedelta(days=1)).strftime('%Y-%m-%d')
    return dt.strftime('%Y-%m-%d')


def run(orders_dir: Path, output_dir: Path, log: LogFn) -> Path:
    """
    读取 orders_dir 中所有 XLSX，按日期统计订单量。
    返回输出 Excel 路径。
    """
    log('=' * 60)
    log('每日订单量扫描')
    log('=' * 60)
    log(f'数据目录: {orders_dir}')

    xlsx_files = sorted(orders_dir.glob('*.xlsx'))
    if not xlsx_files:
        raise ValueError('未找到 XLSX 文件，请检查 orders/ 目录。')

    dfs = []
    for fp in xlsx_files:
        log(f'读取: {fp.name}')
        df = pd.read_excel(fp, usecols=['下单时间'])
        log(f'  订单数: {len(df):,}')
        dfs.append(df)

    all_orders = pd.concat(dfs, ignore_index=True)
    all_orders['下单时间'] = pd.to_datetime(all_orders['下单时间'], format='%m/%d/%Y %H:%M:%S')
    all_orders['日期'] = all_orders['下单时间'].apply(_get_day_label)

    daily = all_orders.groupby('日期').size().reset_index(name='单量').sort_values('日期')

    log(f'\n{"日期":<14} 单量')
    log('-' * 26)
    for _, row in daily.iterrows():
        log(f'  {row["日期"]}  {row["单量"]:,}')
    log('-' * 26)
    log(f'  合计: {daily["单量"].sum():,}')

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / '每日单量统计.xlsx'
    daily.to_excel(output_path, index=False)
    log(f'\n✓ 已保存: {output_path.name}')

    return output_path
