import os
import glob
from datetime import timedelta

import pandas as pd

ORDER_DIR = os.path.join(os.path.dirname(__file__), "orders")
TIME_COL = "下单时间"
CUTOFF_HOUR = 8  # 每天的时间窗口: 前一天08:00 ~ 当天08:00


def get_day_label(dt):
    """根据08:00切分规则，返回该订单所属的日期标签。
    例如 02/05 08:00:00 ~ 02/06 08:00:00 归为 02/06"""
    if dt.hour >= CUTOFF_HOUR:
        return (dt + timedelta(days=1)).strftime("%Y-%m-%d")
    else:
        return dt.strftime("%Y-%m-%d")


def main():
    files = sorted(glob.glob(os.path.join(ORDER_DIR, "*.xlsx")))
    dfs = []
    for fp in files:
        print(f"读取: {os.path.basename(fp)}")
        df = pd.read_excel(fp, usecols=[TIME_COL])
        print(f"  订单数: {len(df)}")
        dfs.append(df)

    all_orders = pd.concat(dfs, ignore_index=True)
    all_orders[TIME_COL] = pd.to_datetime(all_orders[TIME_COL], format="%m/%d/%Y %H:%M:%S")
    all_orders["日期"] = all_orders[TIME_COL].apply(get_day_label)

    daily = all_orders.groupby("日期").size().reset_index(name="单量").sort_values("日期")

    print(f"\n{'日期':<14}| 单量")
    print("-" * 26)
    for _, row in daily.iterrows():
        print(f"{row['日期']:<14}| {row['单量']}")
    print("-" * 26)
    print(f"{'合计':<14}| {daily['单量'].sum()}")

    # 输出到 xlsx
    output_path = os.path.join(os.path.dirname(__file__), "每日单量统计.xlsx")
    daily.to_excel(output_path, index=False)
    print(f"\n已输出到: {output_path}")


if __name__ == "__main__":
    main()
