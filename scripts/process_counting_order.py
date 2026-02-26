"""
订单处理脚本 — 整合自 func/process_counting_order/
支持单日处理 (run_daily) 和周处理 (run_weekly)。
"""

from __future__ import annotations

from datetime import timedelta, date
from pathlib import Path
from typing import Callable

import pandas as pd
import pgeocode

LogFn = Callable[[str], None]

STATE_NAME_TO_CODE: dict[str, str] = {
    'Alabama': 'AL', 'Alaska': 'AK', 'Arizona': 'AZ', 'Arkansas': 'AR',
    'California': 'CA', 'Colorado': 'CO', 'Connecticut': 'CT', 'Delaware': 'DE',
    'Florida': 'FL', 'Georgia': 'GA', 'Hawaii': 'HI', 'Idaho': 'ID',
    'Illinois': 'IL', 'Indiana': 'IN', 'Iowa': 'IA', 'Kansas': 'KS',
    'Kentucky': 'KY', 'Louisiana': 'LA', 'Maine': 'ME', 'Maryland': 'MD',
    'Massachusetts': 'MA', 'Michigan': 'MI', 'Minnesota': 'MN', 'Mississippi': 'MS',
    'Missouri': 'MO', 'Montana': 'MT', 'Nebraska': 'NE', 'Nevada': 'NV',
    'New Hampshire': 'NH', 'New Jersey': 'NJ', 'New Mexico': 'NM', 'New York': 'NY',
    'North Carolina': 'NC', 'North Dakota': 'ND', 'Ohio': 'OH', 'Oklahoma': 'OK',
    'Oregon': 'OR', 'Pennsylvania': 'PA', 'Rhode Island': 'RI', 'South Carolina': 'SC',
    'South Dakota': 'SD', 'Tennessee': 'TN', 'Texas': 'TX', 'Utah': 'UT',
    'Vermont': 'VT', 'Virginia': 'VA', 'Washington': 'WA', 'West Virginia': 'WV',
    'Wisconsin': 'WI', 'Wyoming': 'WY', 'District of Columbia': 'DC',
    'PUERTO RICO': 'PR', 'Puerto Rico': 'PR',
}


def _get_state_from_zip(zip_code, nomi: pgeocode.Nominatim) -> tuple[str, str]:
    if pd.isna(zip_code):
        return ('Unknown', 'Unknown')
    zip_str = str(int(zip_code)).zfill(5)
    if zip_str.startswith(('006', '007', '008', '009')):
        return ('PR', 'Puerto Rico')
    loc = nomi.query_postal_code(zip_str)
    if pd.notna(loc.state_code):
        return (loc.state_code, loc.state_name if pd.notna(loc.state_name) else loc.state_code)
    return ('Unknown', 'Unknown')


def _business_date(ts) -> date:
    """8am 截止规则：hour < 8 归当天，否则归次日。"""
    if ts.hour < 8:
        return ts.date()
    return ts.date() + timedelta(days=1)


def _process_temu(temu_file: Path, nomi: pgeocode.Nominatim, log: LogFn) -> pd.DataFrame:
    log(f"--- 读取 TEMU: {temu_file.name} ---")
    df = pd.read_excel(temu_file)
    log(f"  TEMU 总订单: {len(df):,}")

    df['Date'] = pd.to_datetime(df['下单时间'], format='%m/%d/%Y %H:%M:%S')
    df['Date_Only'] = df['Date'].apply(_business_date)

    unique_zips = df['邮编'].unique()
    log(f"  处理 {len(unique_zips):,} 个唯一邮编...")
    zip_to_code: dict = {}
    zip_to_name: dict = {}
    for i, z in enumerate(unique_zips):
        if (i + 1) % 500 == 0:
            log(f"    已处理 {i + 1:,} / {len(unique_zips):,} 个邮编...")
        code, name = _get_state_from_zip(z, nomi)
        zip_to_code[z] = code
        zip_to_name[z] = name

    df['State_Code'] = df['邮编'].map(zip_to_code)
    df['State_Name'] = df['邮编'].map(zip_to_name)

    temu_min = df['Date_Only'].min()
    temu_max = df['Date_Only'].max()
    log(f"  日期范围: {temu_min} ~ {temu_max}")
    return df


def _process_yy(yy_file: Path, temu_tracking: set, log: LogFn) -> pd.DataFrame:
    log(f"--- 读取 YY: {yy_file.name} ---")
    df = pd.read_csv(yy_file, encoding='utf-16', sep='\t')
    log(f"  YY 总订单: {len(df):,}")

    df['Date'] = pd.to_datetime(df['预约揽收时间'])
    df['Date_Only'] = df['Date'].dt.date

    df_before = df.copy()

    df = df[~df['单据号'].isin(temu_tracking)]
    log(f"  去重后: {len(df):,}")

    df['Weekday'] = df['Date'].dt.dayofweek
    df = df[df['Weekday'] < 5]
    log(f"  过滤工作日后: {len(df):,}")

    state_lower = {k.lower(): v for k, v in STATE_NAME_TO_CODE.items()}
    df['State_Code'] = df['目的地行政洲'].str.strip().str.lower().map(state_lower)
    df['State_Name'] = df['目的地行政洲'].str.strip()

    unmapped = df[df['State_Code'].isna()]
    if len(unmapped) > 0:
        log(f"  WARNING: {len(unmapped)} 条订单无法映射州:")
        for s in unmapped['目的地行政洲'].unique():
            log(f"    - {s}")
        df.loc[df['State_Code'].isna(), 'State_Code'] = 'Unknown'
        df.loc[df['State_Name'] == 'Unknown', 'State_Name'] = 'Unknown'

    return df, df_before


def run_daily(
    temu_file: Path,
    yy_file: Path,
    output_dir: Path,
    log: LogFn,
) -> dict[str, Path]:
    """
    处理单日订单数据。
    返回 {'orders': Path, 'daily_stats': Path}。
    """
    date_str = temu_file.stem.replace('TEMU_', '')
    log("=" * 60)
    log(f"单日处理: {date_str}")
    log("=" * 60)

    log("初始化 pgeocode...")
    nomi = pgeocode.Nominatim('us')

    df_temu = _process_temu(temu_file, nomi, log)
    tracking = set(df_temu['运单号'].dropna())
    log(f"  TEMU tracking numbers: {len(tracking):,}")

    df_yy, df_yy_before = _process_yy(yy_file, tracking, log)

    # 按州汇总
    temu_counts = df_temu.groupby(['State_Code', 'State_Name']).size().reset_index(name='TEMU_Orders')
    yy_counts = df_yy.groupby(['State_Code', 'State_Name']).size().reset_index(name='YY_Orders')

    combined = pd.merge(temu_counts, yy_counts, on=['State_Code', 'State_Name'], how='outer').fillna(0)
    combined['TEMU_Orders'] = combined['TEMU_Orders'].astype(int)
    combined['YY_Orders'] = combined['YY_Orders'].astype(int)
    combined['Total'] = combined['TEMU_Orders'] + combined['YY_Orders']
    combined = combined.sort_values('Total', ascending=False)

    output_dir.mkdir(parents=True, exist_ok=True)
    orders_path = output_dir / f'orders_daily_{date_str}.csv'
    combined.to_csv(orders_path, index=False, encoding='utf-8-sig')

    log(f"\n汇总:")
    log(f"  州/地区: {len(combined)}")
    log(f"  TEMU 总单: {combined['TEMU_Orders'].sum():,}")
    log(f"  YY 总单(去重后): {combined['YY_Orders'].sum():,}")
    log(f"  合并总单: {combined['Total'].sum():,}")
    log(f"\n前 10 州:")
    for _, row in combined.head(10).iterrows():
        log(f"  {row['State_Code']}  TEMU {row['TEMU_Orders']:,}  YY {row['YY_Orders']:,}  合计 {row['Total']:,}")

    # 每日统计
    temu_by_date = df_temu.groupby('Date_Only').size().reset_index(name='TEMU_Orders')
    df_yy_before['Weekday'] = df_yy_before['Date'].dt.dayofweek
    yy_before_wd = df_yy_before[df_yy_before['Weekday'] < 5]
    yy_before_by_date = yy_before_wd.groupby('Date_Only').size().reset_index(name='YY_Before_Dedup')
    yy_after_by_date = df_yy.groupby('Date_Only').size().reset_index(name='YY_After_Dedup')

    stats = temu_by_date.merge(yy_before_by_date, on='Date_Only', how='outer')
    stats = stats.merge(yy_after_by_date, on='Date_Only', how='outer').fillna(0)
    for col in ['TEMU_Orders', 'YY_Before_Dedup', 'YY_After_Dedup']:
        stats[col] = stats[col].astype(int)
    stats['Combined_Total'] = stats['TEMU_Orders'] + stats['YY_After_Dedup']
    stats = stats.sort_values('Date_Only').rename(columns={'Date_Only': 'Date'})

    daily_path = output_dir / f'daily_statistics_{date_str}.csv'
    stats.to_csv(daily_path, index=False, encoding='utf-8-sig')

    log(f"\n✓ 已保存: {orders_path.name}")
    log(f"✓ 已保存: {daily_path.name}")
    return {'orders': orders_path, 'daily_stats': daily_path}


def run_weekly(
    temu_files: list[Path],
    yy_file: Path,
    output_dir: Path,
    log: LogFn,
) -> dict[str, Path]:
    """
    处理周订单数据（支持多个 TEMU 文件）。
    返回 {'orders': Path, 'daily_stats': Path}。
    """
    stems = [f.stem.replace('TEMU_', '') for f in temu_files]
    date_range = f"{stems[0].split('-')[0]}-{stems[-1].split('-')[-1]}" if len(stems) > 1 else stems[0]

    log("=" * 60)
    log(f"周处理: {date_range}")
    log("=" * 60)

    log("初始化 pgeocode...")
    nomi = pgeocode.Nominatim('us')

    df_temu_list: list[pd.DataFrame] = []
    tracking: set = set()

    for tf in temu_files:
        df_t = _process_temu(tf, nomi, log)
        file_tracking = set(df_t['运单号'].dropna())
        tracking.update(file_tracking)
        log(f"  此文件 tracking numbers: {len(file_tracking):,}")
        df_temu_list.append(df_t)

    df_temu_all = pd.concat(df_temu_list, ignore_index=True)
    log(f"\n  TEMU 合并总单: {len(df_temu_all):,}")
    log(f"  唯一 tracking numbers: {len(tracking):,}")

    df_yy, df_yy_before = _process_yy(yy_file, tracking, log)

    # 按日期+州汇总
    temu_counts = df_temu_all.groupby(['Date_Only', 'State_Code', 'State_Name']).size().reset_index(name='Order_Count')
    yy_counts = df_yy.groupby(['Date_Only', 'State_Code', 'State_Name']).size().reset_index(name='Order_Count')

    combined = pd.concat([temu_counts, yy_counts], ignore_index=True)
    combined = combined.groupby(['Date_Only', 'State_Code', 'State_Name'])['Order_Count'].sum().reset_index()

    pivot = combined.pivot_table(
        index=['State_Code', 'State_Name'],
        columns='Date_Only',
        values='Order_Count',
        fill_value=0,
    ).astype(int)
    pivot['Total'] = pivot.sum(axis=1)
    pivot = pivot.sort_values('Total', ascending=False).reset_index()

    date_cols = sorted([c for c in pivot.columns if isinstance(c, date)])
    pivot = pivot[['State_Code', 'State_Name'] + date_cols + ['Total']]

    output_dir.mkdir(parents=True, exist_ok=True)
    orders_path = output_dir / f'orders_weekly_{date_range}.csv'
    pivot.to_csv(orders_path, index=False, encoding='utf-8-sig')

    log(f"\n汇总:")
    log(f"  州/地区: {len(pivot)}")
    log(f"  日期列: {len(date_cols)}")
    log(f"  TEMU 总单: {temu_counts['Order_Count'].sum():,}")
    log(f"  YY 总单(去重后): {yy_counts['Order_Count'].sum():,}")
    log(f"  合并总单: {pivot['Total'].sum():,}")

    # 每日统计
    yy_orig = df_yy_before.groupby('Date_Only').size().to_dict()
    daily_rows = []
    for d in date_cols:
        temu_d = len(df_temu_all[df_temu_all['Date_Only'] == d])
        yy_orig_d = yy_orig.get(d, 0)
        yy_dedup_d = int(yy_counts[yy_counts['Date_Only'] == d]['Order_Count'].sum())
        daily_rows.append({
            'Date': d,
            'TEMU_Orders': temu_d,
            'YY_Original': yy_orig_d,
            'YY_Dedup': yy_dedup_d,
            'Combined_Total': temu_d + yy_dedup_d,
        })

    daily_df = pd.DataFrame(daily_rows)
    daily_path = output_dir / f'daily_statistics_{date_range}.csv'
    daily_df.to_csv(daily_path, index=False, encoding='utf-8-sig')

    log(f"\n✓ 已保存: {orders_path.name}")
    log(f"✓ 已保存: {daily_path.name}")
    return {'orders': orders_path, 'daily_stats': daily_path}
