import os
import sys
import threading
from pathlib import Path

# 打包后将 matplotlib 字体缓存持久化到 app/ 目录，避免每次冷启动重建
if getattr(sys, 'frozen', False):
    _mpl_cache = Path(sys.executable).parent / 'app' / 'matplotlib'
    _mpl_cache.mkdir(parents=True, exist_ok=True)
    os.environ['MPLCONFIGDIR'] = str(_mpl_cache)

# 后台预热 matplotlib 字体缓存，趁用户浏览主页时完成
def _prewarm_matplotlib() -> None:
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.font_manager  # noqa: F401

threading.Thread(target=_prewarm_matplotlib, daemon=True).start()

from utils.setup import init_app_structure
init_app_structure()

from nicegui import ui

from pages.home import create as home_page
from pages.process_counting_order import create as process_counting_order_page
from pages.generate_weekly_trends import create as generate_weekly_trends_page
from pages.weekly_order_scan import create as weekly_order_scan_page
from pages.driver_week_analyze import create as driver_week_analyze_page
from pages.driver_missions import create as driver_missions_page
from pages.daily_report import create as daily_report_page


@ui.page('/')
def index():
    home_page()


@ui.page('/process-counting-order')
def process_counting_order():
    process_counting_order_page()


@ui.page('/generate-weekly-trends')
def generate_weekly_trends():
    generate_weekly_trends_page()


@ui.page('/weekly-order-scan')
def weekly_order_scan():
    weekly_order_scan_page()


@ui.page('/driver-week-analyze')
def driver_week_analyze():
    driver_week_analyze_page()


@ui.page('/driver-missions')
def driver_missions():
    driver_missions_page()


@ui.page('/daily-report')
def daily_report():
    daily_report_page()


ui.run(title='TaskRunner', native=True, reload=False)
