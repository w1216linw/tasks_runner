"""司机任务点数统计页面"""

from __future__ import annotations

import asyncio
from pathlib import Path

from nicegui import run, ui

from components.layout import back_button, sidebar
from scripts.driver_missions import run as run_missions
from utils.paths import get_feature_dir, open_path

FEATURE = 'driver_missions'


def create() -> None:
    sidebar()

    with ui.column().classes('w-full q-pa-lg'):
        with ui.row().classes('items-center gap-sm'):
            back_button()
            ui.label('司机任务点数统计').classes('text-h4 text-bold')
        ui.separator()

        input_dir = get_feature_dir(FEATURE) / 'input'
        input_dir.mkdir(parents=True, exist_ok=True)

        with ui.card().classes('w-full q-mt-md'):
            ui.label('输入文件').classes('text-subtitle1 text-bold')
            ui.label('请将 Excel 文件放入 data/driver_missions/input/ 目录').classes('text-caption text-grey-7')
            ui.label('要求：第1列=实际揽收时间，第17列=司机').classes('text-caption text-grey-7')

            input_files = sorted(input_dir.glob('*.xlsx'))
            if not input_files:
                ui.label('未找到 .xlsx 文件').classes('text-caption text-negative')
                file_select = None
            else:
                file_select = ui.select(
                    {str(f): f.name for f in input_files},
                    value=str(input_files[0]),
                    label='选择输入文件',
                ).classes('w-full q-mt-sm')

        run_btn = ui.button('运行', icon='play_arrow').classes('q-mt-md')
        log_area = ui.log(max_lines=200).classes('w-full h-48 q-mt-sm font-mono text-xs')

        async def on_run():
            if file_select is None:
                ui.notify('请先放入 .xlsx 文件', type='negative')
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
                await run.io_bound(run_missions, Path(file_select.value), output_dir, log)
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
            open_path(out)

        ui.button('打开输出目录', icon='folder_open', on_click=open_output).classes('q-mt-xs').props('flat')
