"""生成周趋势图页面"""

from __future__ import annotations

import asyncio
from pathlib import Path

from nicegui import run, ui

from components.layout import back_button, sidebar
from scripts.generate_weekly_trends import run as run_trends
from utils.paths import get_feature_dir, open_path

FEATURE = 'generate_weekly_trends'


def create() -> None:
    sidebar()

    with ui.column().classes('w-full q-pa-lg'):
        with ui.row().classes('items-center gap-sm'):
            back_button()
            ui.label('生成周趋势图').classes('text-h4 text-bold')
        ui.separator()

        orders_dir = get_feature_dir(FEATURE) / 'orders'
        orders_dir.mkdir(parents=True, exist_ok=True)

        csv_files = sorted(orders_dir.glob('*.csv'))

        # 文件列表预览（默认折叠）
        with ui.card().classes('w-full q-mt-md'):
            with ui.row().classes('items-center justify-between w-full'):
                ui.label('orders/ 目录中的文件').classes('text-subtitle1 text-bold')
                toggle_btn = ui.button(icon='expand_more').props('flat round dense')

            file_content = ui.column().classes('q-mt-xs')
            file_content.set_visibility(False)
            with file_content:
                if not csv_files:
                    ui.label('未找到 CSV 文件，请放入 data/generate_weekly_trends/orders/').classes('text-caption text-negative')
                else:
                    for f in csv_files:
                        ui.label(f'• {f.name}').classes('text-caption q-ml-sm')

        def toggle_files():
            visible = not file_content.visible
            file_content.set_visibility(visible)
            toggle_btn.props(f'icon={"expand_less" if visible else "expand_more"}')

        toggle_btn.on('click', toggle_files)

        # 历史数据状态
        history_file = get_feature_dir(FEATURE) / '周一历史单量.csv'

        def _history_label() -> str:
            if history_file.exists():
                import pandas as pd
                n = len(pd.read_csv(history_file))
                return f'历史文件: {history_file.name}（{n} 天）'
            return f'历史文件: 尚未创建'

        with ui.card().classes('w-full q-mt-md'):
            hist_label = ui.label(_history_label()).classes('text-caption text-grey-7')

        run_btn = ui.button('运行', icon='play_arrow').classes('q-mt-md')
        log_area = ui.log(max_lines=200).classes('w-full h-48 q-mt-sm font-mono text-xs')

        # 图片预览区域
        img_container = ui.column().classes('w-full q-mt-md')

        async def on_run():
            if not csv_files:
                ui.notify('请先放入 CSV 文件', type='negative')
                return

            log_area.clear()
            img_container.clear()
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
                output_path = await run.io_bound(
                    run_trends,
                    orders_dir,
                    output_dir,
                    log,
                    history_file,
                )
                hist_label.set_text(_history_label())
                queue.put_nowait(None)
                await drain_task

                # 显示生成的图片
                with img_container:
                    ui.label('趋势图预览').classes('text-subtitle1 text-bold')
                    ui.image(str(output_path)).classes('w-full')

                ui.notify('趋势图已生成！', type='positive')
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
            open_path(out)

        ui.button('打开输出目录', icon='folder_open', on_click=open_output).classes('q-mt-xs').props('flat')
