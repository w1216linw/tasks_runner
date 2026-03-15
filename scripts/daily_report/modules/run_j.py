"""Step 5: 预约达成率与揽收占比(j)"""

from copy import copy
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import load_workbook

from utils.chinese_font import setup_chinese_font


def _copy_style_from_above(ws, row: int, col: int):
    """将上一行同列单元格的样式复制到目标单元格。"""
    if row <= 2:
        return
    src = ws.cell(row=row - 1, column=col)
    dst = ws.cell(row=row, column=col)
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)


def run(today: datetime, data_dir: Path, output_dir: Path, log: Callable) -> dict:
    setup_chinese_font()
    plt.rcParams['axes.unicode_minus'] = False

    last2day = today - timedelta(days=2)
    last2day_str = last2day.strftime('%m-%d')
    last2day_full = last2day.strftime('%Y-%m-%d')

    module_dir = data_dir / 'input' / 'yy_pickup_rate'
    reach_file = signin_file = None
    for f in sorted(module_dir.glob('*.csv')):
        cols = pd.read_csv(f, encoding='utf-16', sep='\t', nrows=0).columns
        if '揽收达成率' in cols:
            reach_file = f
        elif '签入率' in cols:
            signin_file = f
    if reach_file is None:
        raise FileNotFoundError(f'未在 {module_dir} 找到含「揽收达成率」列的 CSV')
    if signin_file is None:
        raise FileNotFoundError(f'未在 {module_dir} 找到含「签入率」列的 CSV')

    log(f'读取预约达成率数据 ({last2day_str})')
    df_reach = pd.read_csv(reach_file, encoding='utf-16', sep='\t')
    df_signin = pd.read_csv(signin_file, encoding='utf-16', sep='\t')

    def to_float_percent(x):
        if isinstance(x, str) and '%' in x:
            return float(x.replace('%', '')) / 100
        return float(x) if pd.notnull(x) else 0.0

    df_reach['揽收达成率'] = df_reach['揽收达成率'].apply(to_float_percent)
    df_signin['签入率'] = df_signin['签入率'].apply(to_float_percent)

    total_count = len(df_reach)
    reach_count = int((df_reach['揽收达成率'] == 1.0).sum())
    reach_fail_count = int((df_reach['揽收达成率'] == 0.0).sum())
    reach_zero_ids = df_reach[df_reach['揽收达成率'] == 0.0]['单据号'].astype(str).tolist()
    reach_percent = round(reach_count / total_count * 100, 1) if total_count > 0 else 0

    signin_count = int((df_signin['签入率'] == 1.0).sum())
    signin_fail_count = int((df_signin['签入率'] == 0.0).sum())
    signin_percent = round(signin_count / reach_count * 100, 1) if reach_count > 0 else 0
    fail_percent = round(100 - signin_percent, 1)

    log(f'揽收达成: {reach_count}/{total_count} ({reach_percent}%), 签入: {signin_count} ({signin_percent}%)')

    # 饼图
    fig, ax = plt.subplots(figsize=(6, 6))
    ax.pie([signin_percent, fail_percent], labels=['揽收达成', '揽收未达成'],
           autopct='%1.1f%%', startangle=90,
           colors=['#4CAF50', '#FF5722'], textprops={'fontsize': 14})
    ax.set_title('揽收达成与未达成占比', fontsize=18)
    ax.axis('equal')
    plt.tight_layout()
    j_image_pie = output_dir / '揽收达成与未达成占比.png'
    fig.savefig(j_image_pie, dpi=150)
    plt.close(fig)

    # 保存到 gofo_pickup_data.xlsx 预约 sheet
    excel_path = data_dir / 'gofo_pickup_data.xlsx'
    wb = load_workbook(excel_path)
    ws = wb['预约']
    last_row = ws.max_row
    last_date = ws.cell(row=last_row, column=1).value
    last_date_str = last_date.strftime('%Y-%m-%d') if hasattr(last_date, 'strftime') else (str(last_date) if last_date else '')
    if last_date_str == last2day_full:
        log(f'gofo_pickup_data.xlsx 预约 sheet 已有 {last2day_full} 数据，跳过写入')
    else:
        new_row = last_row + 1
        ws.cell(row=new_row, column=1, value=last2day_full)
        ws.cell(row=new_row, column=2, value=total_count)
        ws.cell(row=new_row, column=3, value=reach_count)
        ws.cell(row=new_row, column=4, value=reach_percent / 100)
        ws.cell(row=new_row, column=5, value=signin_count)
        ws.cell(row=new_row, column=6, value=signin_percent / 100)
        ws.cell(row=new_row, column=7, value=reach_fail_count)
        ws.cell(row=new_row, column=8, value=0)
        for col in range(1, 9):
            _copy_style_from_above(ws, new_row, col)
        wb.save(excel_path)
        log('已保存到 gofo_pickup_data.xlsx 预约 sheet')

    return {
        'reach_percent': reach_percent,
        'signin_count': signin_count,
        'signin_percent': signin_percent,
        'signin_fail_count': signin_fail_count,
        'fail_percent': fail_percent,
        'j_image_pie': str(j_image_pie),
        'reach_zero_ids': reach_zero_ids,
    }
