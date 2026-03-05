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
