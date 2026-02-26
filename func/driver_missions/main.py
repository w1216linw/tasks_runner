import openpyxl
import csv
from collections import defaultdict

# 需要统计的司机
TARGET_DRIVERS = {"韩境烨", "安博", "李健", "朱海山", "王洋", "吴波", "陈磊", "邢旭", "刘增榕"}

# 读取 Excel
wb = openpyxl.load_workbook("t.xlsx")
ws = wb["Sheet1"]

# 统计每个司机每天的任务点数（每行 = 1 个点）
counts = defaultdict(lambda: defaultdict(int))

for row in ws.iter_rows(min_row=2, values_only=True):
    driver = row[16]  # 司机
    if driver not in TARGET_DRIVERS:
        continue
    pickup_time = row[0]  # 实际揽收时间
    if pickup_time is None:
        continue
    date = str(pickup_time).split(" ")[0]  # 取日期部分
    counts[date][driver] += 1

# 按日期排序
sorted_dates = sorted(counts.keys(), reverse=True)

# 输出 CSV
with open("司机点数统计.csv", "w", newline="", encoding="utf-8-sig") as f:
    writer = csv.writer(f)
    writer.writerow(["司机", "任务数", "日期"])
    for date in sorted_dates:
        for driver in ["韩境烨", "安博", "李健", "朱海山", "王洋", "吴波", "陈磊", "邢旭", "刘增榕"]:
            task_count = counts[date].get(driver, 0)
            writer.writerow([driver, task_count, date])

print("完成！结果已写入 司机点数统计.csv")
