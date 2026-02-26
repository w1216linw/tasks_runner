"""处理计单页面"""

from __future__ import annotations

import asyncio
from pathlib import Path

from nicegui import run, ui

from components.layout import back_button, sidebar
from scripts.process_counting_order import run_daily, run_weekly
from utils.paths import get_feature_dir

FEATURE = 'process_counting_order'


def _scan_files(subdir: str, glob: str) -> list[Path]:
    d = get_feature_dir(FEATURE) / subdir
    d.mkdir(parents=True, exist_ok=True)
    return sorted(d.glob(glob))


def create() -> None:
    sidebar()

    with ui.column().classes('w-full q-pa-lg'):
        with ui.row().classes('items-center gap-sm'):
            back_button()
            ui.label('处理计单').classes('text-h4 text-bold')
        ui.separator()

        # 模式选择
        mode = ui.toggle({'daily': '单日', 'weekly': '周'}, value='daily').classes('q-mt-md')

        with ui.card().classes('w-full q-mt-md'):
            ui.label('文件选择').classes('text-subtitle1 text-bold')

            # TEMU 文件列表
            with ui.row().classes('items-start gap-md w-full'):
                with ui.column().classes('flex-1'):
                    ui.label('TEMU 文件 (.xlsx)').classes('text-caption text-grey-7')
                    temu_files = _scan_files('TEMU', '*.xlsx')
                    if not temu_files:
                        ui.label('未找到文件，请放入 data/process_counting_order/TEMU/').classes('text-caption text-negative')
                        temu_select = None
                        temu_multi = None
                    else:
                        options = {str(f): f.name for f in temu_files}
                        temu_select = ui.select(options, label='选择 TEMU 文件（单日）', value=str(temu_files[0])).classes('w-full')
                        temu_multi = ui.select(
                            options, label='选择 TEMU 文件（周，可多选）',
                            multiple=True, value=[str(temu_files[0])],
                        ).classes('w-full')

                with ui.column().classes('flex-1'):
                    ui.label('YY 文件 (.csv)').classes('text-caption text-grey-7')
                    yy_files = _scan_files('YY', '*.csv')
                    if not yy_files:
                        ui.label('未找到文件，请放入 data/process_counting_order/YY/').classes('text-caption text-negative')
                        yy_select = None
                    else:
                        yy_options = {str(f): f.name for f in yy_files}
                        yy_select = ui.select(yy_options, label='选择 YY 文件', value=str(yy_files[0])).classes('w-full')

        # 根据模式显示/隐藏选择器
        if temu_select and temu_multi:
            def update_visibility():
                temu_select.set_visibility(mode.value == 'daily')
                temu_multi.set_visibility(mode.value == 'weekly')
            mode.on('update:model-value', lambda _: update_visibility())
            update_visibility()

        # 运行区域
        run_btn = ui.button('运行', icon='play_arrow').classes('q-mt-md')
        log_area = ui.log(max_lines=200).classes('w-full h-64 q-mt-sm font-mono text-xs')
        output_label = ui.label('').classes('text-caption text-positive q-mt-xs')

        async def on_run():
            if (temu_select is None and temu_multi is None) or yy_select is None:
                ui.notify('请先放入所需文件', type='negative')
                return

            log_area.clear()
            output_label.set_text('')
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
                if mode.value == 'daily':
                    result = await run.io_bound(
                        run_daily,
                        Path(temu_select.value),
                        Path(yy_select.value),
                        output_dir,
                        log,
                    )
                else:
                    selected = temu_multi.value if isinstance(temu_multi.value, list) else [temu_multi.value]
                    result = await run.io_bound(
                        run_weekly,
                        [Path(p) for p in selected],
                        Path(yy_select.value),
                        output_dir,
                        log,
                    )
                queue.put_nowait(None)
                await drain_task

                files = '  |  '.join(p.name for p in result.values())
                output_label.set_text(f'输出文件: {files}')
                ui.notify('处理完成！', type='positive')
            except Exception as e:
                queue.put_nowait(f'ERROR: {e}')
                queue.put_nowait(None)
                await drain_task
                ui.notify(f'出错: {e}', type='negative')
            finally:
                run_btn.enable()

        run_btn.on('click', on_run)

        # 打开输出目录按钮
        import subprocess, sys as _sys

        def open_output():
            out = get_feature_dir(FEATURE) / 'output'
            out.mkdir(exist_ok=True)
            if _sys.platform == 'darwin':
                subprocess.Popen(['open', str(out)])
            elif _sys.platform == 'win32':
                subprocess.Popen(['explorer', str(out)])

        ui.button('打开输出目录', icon='folder_open', on_click=open_output).classes('q-mt-xs').props('flat')
