"""
周趋势图生成脚本 — 整合自 func/generate_weekly_trends/
"""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path
from typing import Callable

import pandas as pd

LogFn = Callable[[str], None]


def _extract_daily_orders(orders_dir: Path, log: LogFn) -> pd.DataFrame:
    daily_data: list[dict] = []

    # 旧格式: orders_*.csv (按州列出，日期为列名)
    for csv_file in sorted(orders_dir.glob('orders_*.csv')):
        df = pd.read_csv(csv_file, encoding='utf-8-sig')
        date_cols = [
            c for c in df.columns
            if '-' in c and any(c.startswith(y) for y in ['2025-', '2026-'])
        ]
        for dc in date_cols:
            daily_data.append({
                'date': dc,
                'orders': df[dc].sum(),
                'is_monday': datetime.strptime(dc, '%Y-%m-%d').weekday() == 0,
            })

    # 新格式: daily_statistics_*.csv (每行一天，有 Combined_Total 列)
    for csv_file in sorted(orders_dir.glob('daily_statistics_*.csv')):
        df = pd.read_csv(csv_file, encoding='utf-8-sig')
        for _, row in df.iterrows():
            date_str = str(row['Date'])
            try:
                date_obj = datetime.strptime(date_str, '%m/%d/%Y')
            except ValueError:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            formatted = date_obj.strftime('%Y-%m-%d')
            daily_data.append({
                'date': formatted,
                'orders': row['Combined_Total'],
                'is_monday': date_obj.weekday() == 0,
            })

    if not daily_data:
        log("  没有找到可用的数据文件。")
        return pd.DataFrame(columns=['date', 'orders', 'is_monday'])

    result = pd.DataFrame(daily_data)
    result['date'] = pd.to_datetime(result['date'])
    result = result.sort_values('date').drop_duplicates(subset=['date'], keep='last').reset_index(drop=True)
    result['date'] = result['date'].dt.strftime('%Y-%m-%d')
    return result


def _plot_trend(df: pd.DataFrame, output_path: Path, log: LogFn) -> None:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False

    fig, ax = plt.subplots(figsize=(20, 8))
    dates = pd.to_datetime(df['date'])
    display_dates = dates.dt.strftime('%m/%d')

    ax.plot(range(len(df)), df['orders'],
            marker='o', linewidth=2, markersize=6,
            color='#2E86AB', label='每日订单量')

    for i in range(len(df)):
        d = display_dates.iloc[i]
        orders = df.iloc[i]['orders']
        is_monday = df.iloc[i]['is_monday']
        is_last = (i == len(df) - 1)

        if is_monday:
            if is_last:
                ax.plot(i, orders, marker='o', markersize=12, color='#E63946', zorder=5)
                ax.annotate(
                    f'{d}\n{orders:,.0f}单',
                    xy=(i, orders), xytext=(10, 15), textcoords='offset points',
                    fontsize=10, fontweight='bold', color='#E63946',
                    bbox=dict(boxstyle='round,pad=0.5', facecolor='white',
                              edgecolor='#E63946', linewidth=2),
                )
            else:
                ax.annotate(
                    f'{d}\n{orders:,.0f}单',
                    xy=(i, orders), xytext=(0, 12), textcoords='offset points',
                    fontsize=8, ha='center', color='#2E86AB',
                    bbox=dict(boxstyle='round,pad=0.4', facecolor='white',
                              edgecolor='#2E86AB', linewidth=1.5, alpha=0.9),
                )
        else:
            ax.text(i, orders, f'{orders:,.0f}',
                    ha='center', va='bottom', fontsize=7, color='#666666')

    ax.set_xticks(range(len(df)))
    x_labels = [display_dates.iloc[i] if df.iloc[i]['is_monday'] else '' for i in range(len(df))]
    ax.set_xticklabels(x_labels, rotation=45, fontsize=9, ha='right')
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, _: f'{int(x):,}'))
    ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.5)
    ax.set_axisbelow(True)
    ax.set_title('每日订单量趋势图（周一重点标注）', fontsize=16, fontweight='bold', pad=20)
    ax.set_xlabel('日期 (月/日)', fontsize=12, fontweight='bold')
    ax.set_ylabel('订单量', fontsize=12, fontweight='bold')
    ax.legend(loc='upper left', fontsize=10)
    plt.tight_layout()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    plt.savefig(output_path, dpi=300, bbox_inches='tight', facecolor='white')
    plt.close(fig)
    log(f"✓ 趋势图已保存: {output_path.name}")


def _merge_into_history(new_df: pd.DataFrame, history_file: Path, log: LogFn) -> pd.DataFrame:
    """
    将新数据合并进历史 CSV，按日期去重后保存，返回完整数据。
    备份策略：数据有变化时，先删除旧备份，再把当前历史备份，最后写入新历史。
    结果始终只保留一份 backup（上一版本）。
    """
    backup_file = history_file.with_name(history_file.stem + '_backup.csv')

    if history_file.exists():
        hist_df = pd.read_csv(history_file, encoding='utf-8-sig')
        log(f"读取历史数据: {len(hist_df)} 天")
        combined = pd.concat(
            [hist_df[['date', 'orders']], new_df[['date', 'orders']]],
            ignore_index=True,
        )
    else:
        log("历史文件不存在，将创建新文件")
        hist_df = None
        combined = new_df[['date', 'orders']].copy()

    combined['date'] = pd.to_datetime(combined['date'])
    combined = (combined
                .sort_values('date')
                .drop_duplicates(subset=['date'], keep='last')
                .reset_index(drop=True))
    combined['is_monday'] = combined['date'].dt.weekday == 0
    combined['date'] = combined['date'].dt.strftime('%Y-%m-%d')

    # 检测数据是否发生变化
    if hist_df is None:
        changed = True
    else:
        changed = not (
            len(hist_df) == len(combined)
            and (hist_df['date'].astype(str) == combined['date']).all()
            and (hist_df['orders'].astype(float) == combined['orders'].astype(float)).all()
        )

    if changed and history_file.exists():
        if backup_file.exists():
            backup_file.unlink()
            log(f"已删除旧备份: {backup_file.name}")
        shutil.copy2(history_file, backup_file)
        log(f"已备份当前历史 → {backup_file.name}")

    history_file.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(history_file, index=False, encoding='utf-8-sig')

    if changed:
        log(f"✓ 历史文件已更新: {history_file.name}（共 {len(combined)} 天）")
    else:
        log(f"  数据无变化，历史文件未修改（共 {len(combined)} 天）")

    return combined


def run(orders_dir: Path, output_dir: Path, log: LogFn, history_file: Path | None = None) -> Path:
    """
    读取 orders_dir 中的 CSV 文件，合并进历史后生成趋势图 PNG。
    history_file: 历史数据 CSV 路径，为 None 时仅用本次数据。
    返回输出图片路径。
    """
    log("=" * 60)
    log("生成周趋势图")
    log("=" * 60)
    log(f"数据目录: {orders_dir}")

    log("读取本次订单数据...")
    df_new = _extract_daily_orders(orders_dir, log)

    if df_new.empty:
        raise ValueError("没有找到任何订单数据，请检查 orders/ 目录。")

    log(f"本次读取: {len(df_new)} 天")

    if history_file is not None:
        log("\n合并到历史文件...")
        df = _merge_into_history(df_new, history_file, log)
    else:
        df = df_new

    mondays = df[df['is_monday']]
    log(f"\n共 {len(df)} 天数据，含 {len(mondays)} 个周一:")
    for _, row in mondays.iterrows():
        log(f"  {row['date']} (周一): {row['orders']:,.0f} 单")

    log("\n生成趋势图...")
    output_path = output_dir / 'monday_order_trend.png'
    _plot_trend(df, output_path, log)

    log("✓ 完成!")
    return output_path
