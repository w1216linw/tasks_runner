"""每日订单量扫描页面"""

from __future__ import annotations

import asyncio
import subprocess
import sys
from pathlib import Path

from nicegui import run, ui

from components.layout import back_button, sidebar
from scripts.weekly_order_scan import run as run_scan
from utils.paths import get_feature_dir

FEATURE = 'weekly_order_scan'


def create() -> None:
    sidebar()

    with ui.column().classes('w-full q-pa-lg'):
        with ui.row().classes('items-center gap-sm'):
            back_button()
            ui.label('每日订单量扫描').classes('text-h4 text-bold')
        ui.separator()

        orders_dir = get_feature_dir(FEATURE) / 'orders'
        orders_dir.mkdir(parents=True, exist_ok=True)

        with ui.card().classes('w-full q-mt-md'):
            ui.label('orders/ 目录中的文件').classes('text-subtitle1 text-bold')
            xlsx_files = sorted(orders_dir.glob('*.xlsx'))
            if not xlsx_files:
                ui.label('未找到 XLSX 文件，请放入 data/weekly_order_scan/orders/').classes('text-caption text-negative')
            else:
                for f in xlsx_files:
                    ui.label(f'• {f.name}').classes('text-caption q-ml-sm')

        run_btn = ui.button('运行', icon='play_arrow').classes('q-mt-md')
        log_area = ui.log(max_lines=200).classes('w-full h-48 q-mt-sm font-mono text-xs')

        async def on_run():
            if not xlsx_files:
                ui.notify('请先放入 XLSX 文件', type='negative')
                return

            log_area.clear()
            run_btn.disable()

            loop = asyncio.get_event_loop()
            queue: asyncio.Queue[str | None] = asyncio.Queue()

            def log(msg: str):
                loop.call_soon_threadsafe(queue.put_nowait, msg)

            async def drain():
                while True:
                    msg = await queue.get()
                    if msg is None:
                        break
                    log_area.push(msg)

            drain_task = asyncio.create_task(drain())
            output_dir = get_feature_dir(FEATURE) / 'output'

            try:
                await run.io_bound(run_scan, orders_dir, output_dir, log)
                queue.put_nowait(None)
                await drain_task
                ui.notify('统计完成！', type='positive')
            except Exception as e:
                queue.put_nowait(f'ERROR: {e}')
                queue.put_nowait(None)
                await drain_task
                ui.notify(f'出错: {e}', type='negative')
            finally:
                run_btn.enable()

        run_btn.on('click', on_run)

        def open_output():
            out = get_feature_dir(FEATURE) / 'output'
            out.mkdir(exist_ok=True)
            if sys.platform == 'darwin':
                subprocess.Popen(['open', str(out)])
            elif sys.platform == 'win32':
                subprocess.Popen(['explorer', str(out)])

        ui.button('打开输出目录', icon='folder_open', on_click=open_output).classes('q-mt-xs').props('flat')
