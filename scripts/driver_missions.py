"""
司机任务点数统计脚本 — 整合自 func/driver_missions/main.py
"""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Callable

import openpyxl
import pandas as pd

LogFn = Callable[[str], None]

TARGET_DRIVERS = ['韩境烨', '安博', '李健', '朱海山', '王洋', '吴波', '陈磊', '邢旭', '刘增榕','吴军']
TARGET_SET = set(TARGET_DRIVERS)


def run(input_file: Path, output_dir: Path, log: LogFn) -> Path:
    """
    统计指定司机每天的任务点数。
    input_file: Excel（第1列=实际揽收时间，第17列=司机）
    返回输出 CSV 路径。
    """
    log('=' * 60)
    log('司机任务点数统计')
    log('=' * 60)
    log(f'输入文件: {input_file.name}')

    wb = openpyxl.load_workbook(input_file)
    ws = wb["Sheet1"]

    counts = defaultdict(lambda: defaultdict(int))
    total = 0

    for row in ws.iter_rows(min_row=2, values_only=True):
        driver = row[16]  # 第17列（0-indexed）
        if driver not in TARGET_SET:
            continue
        pickup_time = row[0]  # 第1列
        if pickup_time is None:
            continue
        date = str(pickup_time).split(' ')[0]
        counts[date][driver] += 1
        total += 1

    wb.close()
    log(f'共统计 {total:,} 条有效记录，{len(counts)} 个日期')

    rows = []
    for date in sorted(counts.keys(), reverse=True):
        day_summary = '  '.join(f'{d}={counts[date].get(d, 0)}' for d in TARGET_DRIVERS)
        log(f'  {date}: {day_summary}')
        for driver in TARGET_DRIVERS:
            rows.append({'司机': driver, '任务数': counts[date].get(driver, 0), '日期': date})

    df = pd.DataFrame(rows)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / '司机点数统计.csv'
    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    log(f'\n✓ 已保存: {output_path.name}')

    return output_path
