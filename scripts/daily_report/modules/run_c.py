"""Step 2: 每日下单趋势 + 预约揽收实际单量(c)"""

import re
from datetime import datetime, timedelta, date
from pathlib import Path
from typing import Callable

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator
import pandas as pd

from utils.chinese_font import setup_chinese_font

COLOR_LATEST = '#D32F2F'
COLOR_HIST = '#90A4AE'


def run(today: datetime, data_dir: Path, output_dir: Path, log: Callable) -> dict:
    setup_chinese_font()
    plt.rcParams['axes.unicode_minus'] = False

    module_dir = data_dir / 'input' / 'tubt_daily_trend'
    today_file = module_dir / '下单.xlsx'
    reserve_file = module_dir / '预约.csv'
    history_csv = data_dir / 'historical_orders.csv'

    log(f'读取: {today_file.name}')
    df = pd.read_excel(today_file, engine='openpyxl')
    df['下单时间'] = pd.to_datetime(df.get('下单时间'), errors='coerce')
    df = df.dropna(subset=['下单时间'])
    total_cancel = int((df['状态'] == '已取消').sum())

    # 从文件名解析日期，否则用 today
    m = re.search(r'(\d{1,2})\.(\d{1,2})', today_file.name)
    if m:
        month, day = map(int, m.groups())
        y = today.year
        if (date(y, month, day) - today.date()).days > 180:
            y -= 1
        import_date = datetime(y, month, day)
    else:
        import_date = today

    end = import_date.replace(hour=8, minute=0, second=0, microsecond=0)
    start = end - timedelta(days=1)
    df = df[(df['下单时间'] >= start) & (df['下单时间'] < end)]
    df['hour_shift'] = (df['下单时间'].dt.hour - 8) % 24

    agg = (
        df.groupby('hour_shift')['运单号']
          .nunique()
          .reindex(range(24), fill_value=0)
          .reset_index()
          .rename(columns={'运单号': 'count'})
    )
    agg['import_date'] = import_date.date()
    agg['real_time'] = agg['hour_shift'].apply(lambda h: start + timedelta(hours=int(h)))

    # 合并历史
    if history_csv.exists() and history_csv.stat().st_size > 0:
        hist = pd.read_csv(history_csv)
        hist['import_date'] = pd.to_datetime(hist.get('import_date'), errors='coerce').dt.date
        hist['real_time'] = pd.to_datetime(hist.get('real_time'), errors='coerce')
        hist['hour_shift'] = pd.to_numeric(hist.get('hour_shift', 0), errors='coerce').fillna(0).astype(int)
        hist['count'] = pd.to_numeric(hist.get('count', 0), errors='coerce').fillna(0).astype(int)
        df_all = pd.concat([hist[['hour_shift', 'count', 'import_date', 'real_time']], agg], ignore_index=True)
        df_all = df_all.drop_duplicates(subset=['import_date', 'hour_shift'], keep='last')
    else:
        df_all = agg.copy()

    df_save = df_all.copy()
    df_save['import_date'] = pd.to_datetime(df_save['import_date']).dt.strftime('%Y-%m-%d')
    df_save.to_csv(history_csv, index=False, encoding='utf-8-sig')

    total_orders_today = int(agg['count'].sum())
    log(f'当日下单量: {total_orders_today}, 已取消: {total_cancel}')

    # 折线图
    fig, ax = plt.subplots(figsize=(12, 6))
    latest_date = max(df_all['import_date'])
    for imp_dt, grp in df_all.groupby('import_date'):
        if imp_dt == latest_date:
            continue
        grp = grp.sort_values('hour_shift')
        ax.plot(grp['hour_shift'], grp['count'], linewidth=1.2, color=COLOR_HIST, alpha=0.5)

    latest_grp = df_all[df_all['import_date'] == latest_date].sort_values('hour_shift')
    if not latest_grp.empty:
        x, y = latest_grp['hour_shift'], latest_grp['count']
        ax.plot(x, y, marker='o', linewidth=2.8, color=COLOR_LATEST,
                label=f"{pd.to_datetime(latest_date).strftime('%m-%d')}（最新）", zorder=5)
        ax.text(x.iloc[-1] + 0.3, y.iloc[-1], pd.to_datetime(latest_date).strftime('%m-%d'),
                va='center', fontsize=11, color=COLOR_LATEST, weight='bold')

    ax.set_xticks(range(24))
    ax.set_xticklabels([f'{(h + 8) % 24:02d}:00' for h in range(24)], rotation=45)
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax.set_title('每日下单趋势（窗口：昨日 08:00 → 今日 08:00）', pad=15)
    ax.set_xlabel('时段')
    ax.set_ylabel('运单数')
    ax.grid(True, linestyle='--', alpha=0.5)
    ax.legend(loc='upper left')
    plt.tight_layout()

    c_image_trend = output_dir / '每日下单数据图.png'
    fig.savefig(c_image_trend, dpi=150)
    plt.close(fig)
    log(f'图片已保存: {c_image_trend.name}')

    # 1.b 预约揽收实际单量
    df_temu = pd.read_excel(today_file, dtype={'运单号': str})
    df_reserve = pd.read_csv(reserve_file, dtype={'单据号': str}, encoding='utf-16', sep='\t')
    df_reserve_clean = df_reserve[
        ~(df_reserve['单据号'].isin(df_temu['运单号']) | (df_reserve['预约揽收达成率'] == '100.00%'))
    ]
    log(f'预约揽收: {len(df_reserve)}, 去重后: {len(df_reserve_clean)}')

    return {
        'total_orders_today': total_orders_today,
        'total_cancel': total_cancel,
        'df_reserve': len(df_reserve),
        'df_reserve_clean': len(df_reserve_clean),
        'c_image_trend': str(c_image_trend),
    }
