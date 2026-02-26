import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
from datetime import datetime
import glob
import sys
import io

# 设置输出编码为UTF-8（解决Windows中文显示问题）
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 设置中文字体支持
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

def extract_all_daily_orders():
    """从所有CSV文件中提取每一天的订单总量"""
    orders_dir = Path('orders')

    daily_data = []

    # 处理旧格式文件（orders_*.csv）
    old_format_files = sorted(orders_dir.glob('orders_*.csv'))
    for csv_file in old_format_files:
        # 读取CSV
        df = pd.read_csv(csv_file, encoding='utf-8-sig')

        # 获取所有日期列（格式为 2025-MM-DD 或 YYYY-MM-DD）
        date_columns = [col for col in df.columns if '-' in col and any(col.startswith(year) for year in ['2025-', '2026-'])]

        # 遍历每一天
        for date_col in date_columns:
            # 计算当天的总订单量（所有州求和）
            daily_total = df[date_col].sum()
            daily_data.append({
                'date': date_col,
                'orders': daily_total,
                'is_monday': datetime.strptime(date_col, '%Y-%m-%d').weekday() == 0  # 0=周一
            })

    # 处理新格式文件（daily_statistics_*.csv）
    new_format_files = sorted(orders_dir.glob('daily_statistics_*.csv'))
    for csv_file in new_format_files:
        # 读取CSV
        df = pd.read_csv(csv_file, encoding='utf-8-sig')

        # 遍历每一行
        for _, row in df.iterrows():
            # 解析日期（支持 MM/DD/YYYY 和 YYYY-MM-DD 两种格式）
            date_str = row['Date']
            try:
                date_obj = datetime.strptime(date_str, '%m/%d/%Y')
            except ValueError:
                date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            formatted_date = date_obj.strftime('%Y-%m-%d')

            # 使用 Combined_Total 作为订单总量
            daily_total = row['Combined_Total']

            daily_data.append({
                'date': formatted_date,
                'orders': daily_total,
                'is_monday': date_obj.weekday() == 0  # 0=周一
            })

    # 按日期排序并去重（如果有重复日期，保留最新的数据）
    result_df = pd.DataFrame(daily_data)
    result_df['date'] = pd.to_datetime(result_df['date'])
    result_df = result_df.sort_values('date').reset_index(drop=True)

    # 去重：如果同一天有多条记录，保留最后一条
    result_df = result_df.drop_duplicates(subset=['date'], keep='last')
    result_df = result_df.reset_index(drop=True)

    result_df['date'] = result_df['date'].dt.strftime('%Y-%m-%d')

    return result_df

def plot_trend(df):
    """绘制每日订单趋势图"""
    # 创建图表 - 增加宽度以容纳更多数据点
    fig, ax = plt.subplots(figsize=(20, 8))

    # 转换日期格式用于显示
    dates = pd.to_datetime(df['date'])
    display_dates = dates.dt.strftime('%m/%d')

    # 绘制折线图
    line = ax.plot(range(len(df)), df['orders'],
                   marker='o',
                   linewidth=2,
                   markersize=6,
                   color='#2E86AB',
                   label='每日订单量')

    # 在所有数据点上添加标注
    for i in range(len(df)):
        date = display_dates.iloc[i]
        orders = df.iloc[i]['orders']
        is_monday = df.iloc[i]['is_monday']
        is_last = (i == len(df) - 1)

        if is_monday:
            # 周一的点标注
            if is_last:
                # 最新的一天（应该是11.03周一）用红色特别标注
                ax.plot(i, orders,
                       marker='o',
                       markersize=12,
                       color='#E63946',
                       zorder=5)

                ax.annotate(f'{date}\n{orders:,.0f}单',
                           xy=(i, orders),
                           xytext=(10, 15),
                           textcoords='offset points',
                           fontsize=10,
                           fontweight='bold',
                           color='#E63946',
                           bbox=dict(boxstyle='round,pad=0.5',
                                    facecolor='white',
                                    edgecolor='#E63946',
                                    linewidth=2))
            else:
                # 之前的周一用蓝色标注框
                ax.annotate(f'{date}\n{orders:,.0f}单',
                           xy=(i, orders),
                           xytext=(0, 12),
                           textcoords='offset points',
                           fontsize=8,
                           fontweight='normal',
                           color='#2E86AB',
                           ha='center',
                           bbox=dict(boxstyle='round,pad=0.4',
                                    facecolor='white',
                                    edgecolor='#2E86AB',
                                    linewidth=1.5,
                                    alpha=0.9))
        else:
            # 非周一的点只显示数值，不加框
            ax.text(i, orders, f'{orders:,.0f}',
                   ha='center', va='bottom',
                   fontsize=7,
                   color='#666666')

    # 设置X轴标签 - 只显示周一的日期
    ax.set_xticks(range(len(df)))
    # 创建标签列表，只在周一位置显示日期
    x_labels = [display_dates.iloc[i] if df.iloc[i]['is_monday'] else '' for i in range(len(df))]
    ax.set_xticklabels(x_labels, rotation=45, fontsize=9, ha='right')

    # 设置Y轴格式
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))

    # 添加网格
    ax.grid(True, alpha=0.3, linestyle='--', linewidth=0.5)
    ax.set_axisbelow(True)

    # 设置标题和标签
    ax.set_title('每日订单量趋势图（周一重点标注）', fontsize=16, fontweight='bold', pad=20)
    ax.set_xlabel('日期 (月/日)', fontsize=12, fontweight='bold')
    ax.set_ylabel('订单量', fontsize=12, fontweight='bold')

    # 添加图例
    ax.legend(loc='upper left', fontsize=10)

    # 优化布局
    plt.tight_layout()

    # 保存图表
    output_file = 'monday_order_trend.png'
    plt.savefig(output_file, dpi=300, bbox_inches='tight', facecolor='white')
    print(f"✓ 趋势图已保存到: {output_file}")

    return output_file

def main():
    print("正在读取订单数据...")
    df = extract_all_daily_orders()

    # 统计信息
    mondays_df = df[df['is_monday']]
    print(f"✓ 已读取 {len(df)} 天的数据（{len(mondays_df)} 个周一）")

    print("\n每周一订单量：")
    for _, row in mondays_df.iterrows():
        date_obj = datetime.strptime(row['date'], '%Y-%m-%d')
        print(f"  {date_obj.strftime('%Y-%m-%d')} (周一): {row['orders']:,} 单")

    print("\n正在生成趋势图...")
    plot_trend(df)

    print("✓ 完成!")

if __name__ == '__main__':
    main()
