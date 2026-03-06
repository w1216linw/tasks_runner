"""首次启动初始化：建立目录结构并写入说明文件。"""

from pathlib import Path

from utils.paths import get_app_dir

_DIRS = [
    'process_counting_order/TEMU',
    'process_counting_order/YY',
    'process_counting_order/output',
    'generate_weekly_trends/orders',
    'weekly_order_scan/orders',
    'driver_week_analyze/input',
    'driver_week_analyze/output',
    'driver_missions/input',
    'daily_report/input/temu_orders',
    'daily_report/input/tubt_daily_trend',
    'daily_report/input/tubt_pending_orders',
    'daily_report/input/self_printed',
    'daily_report/input/yy_pickup_rate',
    'daily_report/input/shein_d2d',
    'daily_report/input/yy_unachieved',
    'daily_report/output',
]


def init_app_structure() -> None:
    """创建所有功能的目录结构。"""
    app_dir = get_app_dir()
    for rel_path in _DIRS:
        (app_dir / rel_path).mkdir(parents=True, exist_ok=True)
