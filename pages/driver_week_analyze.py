"""司机周数据分析页面"""

from __future__ import annotations

import asyncio
import base64
from pathlib import Path

from nicegui import run, ui

from components.layout import back_button, sidebar
from scripts.driver_week_analyze import run_comparison, run_dwa, run_weekly_chart
from utils.paths import get_feature_dir, open_path

FEATURE = 'driver_week_analyze'


def _scan(subdir: str, pattern: str) -> list[Path]:
    d = get_feature_dir(FEATURE) / subdir
    d.mkdir(parents=True, exist_ok=True)
    return sorted(d.glob(pattern))


def _make_async_log(log_area: ui.log):
    """返回 (log_fn, queue, drain_coro_factory)。"""
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

    return log, queue, drain


def create() -> None:
    sidebar()

    with ui.column().classes('w-full q-pa-lg'):
        with ui.row().classes('items-center gap-sm'):
            back_button()
            ui.label('司机周数据分析').classes('text-h4 text-bold')
        ui.separator()

        # ── 第一步: DWA 分析 ──────────────────────────────
        with ui.card().classes('w-full q-mt-md'):
            ui.label('第一步：生成 DWA 分析').classes('text-subtitle1 text-bold')
            ui.label(
                '所需文件放入 data/driver_week_analyze/input/：'
                '司机数据 dwd_*.xlsx  +  加油交易 transaction_*.xlsx'
            ).classes('text-caption text-grey-7')

            dwd_files = _scan('input', 'dwd_*.xlsx')
            txn_files = _scan('input', 'Transaction*_*.xlsx')

            with ui.row().classes('gap-md w-full q-mt-sm'):
                with ui.column().classes('flex-1'):
                    ui.label('司机数据 (dwd_*.xlsx)').classes('text-caption text-grey-7')
                    if not dwd_files:
                        ui.label('未找到').classes('text-caption text-negative')
                        dwd_select = None
                    else:
                        dwd_select = ui.select(
                            {str(f): f.name for f in dwd_files},
                            value=str(dwd_files[0]),
                        ).classes('w-full')

                with ui.column().classes('flex-1'):
                    ui.label('加油交易 (Transaction*_*.xlsx)').classes('text-caption text-grey-7')
                    if not txn_files:
                        ui.label('未找到').classes('text-caption text-negative')
                        txn_select = None
                    else:
                        txn_select = ui.select(
                            {str(f): f.name for f in txn_files},
                            value=str(txn_files[0]),
                        ).classes('w-full')

            dwa_btn = ui.button('生成 DWA 分析', icon='play_arrow').classes('q-mt-sm')
            dwa_log = ui.log(max_lines=100).classes('w-full h-40 q-mt-sm font-mono text-xs')

            async def on_dwa():
                if dwd_select is None or txn_select is None:
                    ui.notify('请先放入所需文件', type='negative')
                    return

                dwa_log.clear()
                dwa_btn.disable()
                log, queue, drain = _make_async_log(dwa_log)
                drain_task = asyncio.create_task(drain())
                output_dir = get_feature_dir(FEATURE) / 'output'

                try:
                    await run.io_bound(
                        run_dwa,
                        Path(dwd_select.value),
                        Path(txn_select.value),
                        output_dir,
                        log,
                    )
                    queue.put_nowait(None)
                    await drain_task
                    ui.notify('DWA 分析完成！', type='positive')
                except Exception as e:
                    queue.put_nowait(f'ERROR: {e}')
                    queue.put_nowait(None)
                    await drain_task
                    ui.notify(f'出错: {e}', type='negative')
                finally:
                    dwa_btn.enable()

            dwa_btn.on('click', on_dwa)

        # ── 第二步: 周对周对比 ────────────────────────────
        with ui.card().classes('w-full q-mt-md'):
            ui.label('第二步：生成周对周对比报表').classes('text-subtitle1 text-bold')
            ui.label(
                '需要两个 dwa_*.xlsx（第一步的输出），放入 data/driver_week_analyze/output/ 后刷新页面。'
            ).classes('text-caption text-grey-7')

            dwa_output_files = _scan('output', 'dwa_*.xlsx')

            with ui.row().classes('gap-md w-full q-mt-sm'):
                with ui.column().classes('flex-1'):
                    ui.label('上周 dwa_*.xlsx').classes('text-caption text-grey-7')
                    if len(dwa_output_files) < 2:
                        ui.label('请先完成两周的 DWA 分析').classes('text-caption text-negative')
                        prev_select = None
                        curr_select = None
                    else:
                        prev_select = ui.select(
                            {str(f): f.name for f in dwa_output_files},
                            value=str(dwa_output_files[0]),
                        ).classes('w-full')

                if len(dwa_output_files) >= 2:
                    with ui.column().classes('flex-1'):
                        ui.label('本周 dwa_*.xlsx').classes('text-caption text-grey-7')
                        curr_select = ui.select(
                            {str(f): f.name for f in dwa_output_files},
                            value=str(dwa_output_files[-1]),
                        ).classes('w-full')

            cmp_btn = ui.button('生成对比报表', icon='compare_arrows').classes('q-mt-sm')
            cmp_log = ui.log(max_lines=100).classes('w-full h-40 q-mt-sm font-mono text-xs')

            async def on_comparison():
                if prev_select is None or curr_select is None:
                    ui.notify('请先完成两周的 DWA 分析', type='negative')
                    return

                cmp_log.clear()
                cmp_btn.disable()
                log, queue, drain = _make_async_log(cmp_log)
                drain_task = asyncio.create_task(drain())
                output_dir = get_feature_dir(FEATURE) / 'output'

                try:
                    await run.io_bound(
                        run_comparison,
                        Path(prev_select.value),
                        Path(curr_select.value),
                        output_dir,
                        log,
                    )
                    queue.put_nowait(None)
                    await drain_task
                    ui.notify('对比报表已生成！', type='positive')
                except Exception as e:
                    queue.put_nowait(f'ERROR: {e}')
                    queue.put_nowait(None)
                    await drain_task
                    ui.notify(f'出错: {e}', type='negative')
                finally:
                    cmp_btn.enable()

            cmp_btn.on('click', on_comparison)

        # ── 第三步: 周报图表 ──────────────────────────────────
        with ui.card().classes('w-full q-mt-md'):
            ui.label('第三步：生成周报图表').classes('text-subtitle1 text-bold')
            ui.label(
                '选择上周和本周的 dwa_*.xlsx（第一步的输出），生成四图对比周报 PNG。'
            ).classes('text-caption text-grey-7')

            with ui.row().classes('gap-md w-full q-mt-sm'):
                with ui.column().classes('flex-1'):
                    ui.label('上周 dwa_*.xlsx').classes('text-caption text-grey-7')
                    if len(dwa_output_files) < 2:
                        ui.label('请先完成两周的 DWA 分析').classes('text-caption text-negative')
                        chart_prev_select = None
                        chart_curr_select = None
                    else:
                        chart_prev_select = ui.select(
                            {str(f): f.name for f in dwa_output_files},
                            value=str(dwa_output_files[0]),
                        ).classes('w-full')

                if len(dwa_output_files) >= 2:
                    with ui.column().classes('flex-1'):
                        ui.label('本周 dwa_*.xlsx').classes('text-caption text-grey-7')
                        chart_curr_select = ui.select(
                            {str(f): f.name for f in dwa_output_files},
                            value=str(dwa_output_files[-1]),
                        ).classes('w-full')

            chart_img_ref: list[Path | None] = [None]

            with ui.row().classes('q-mt-sm gap-sm items-center'):
                chart_btn = ui.button('生成周报图表', icon='bar_chart')
                copy_chart_btn = ui.button('复制图表', icon='content_copy').props('flat')
                copy_chart_btn.disable()

            chart_log = ui.log(max_lines=50).classes('w-full h-32 q-mt-sm font-mono text-xs')
            chart_img = ui.column().classes('w-full q-mt-md')

            async def on_chart():
                if chart_prev_select is None or chart_curr_select is None:
                    ui.notify('请先完成两周的 DWA 分析', type='negative')
                    return

                chart_log.clear()
                chart_img.clear()
                chart_btn.disable()
                copy_chart_btn.disable()
                log, queue, drain = _make_async_log(chart_log)
                drain_task = asyncio.create_task(drain())
                output_dir = get_feature_dir(FEATURE) / 'output'

                try:
                    img_path = await run.io_bound(
                        run_weekly_chart,
                        Path(chart_prev_select.value),
                        Path(chart_curr_select.value),
                        output_dir,
                        log,
                    )
                    queue.put_nowait(None)
                    await drain_task
                    chart_img_ref[0] = img_path
                    with chart_img:
                        ui.label('周报预览').classes('text-subtitle1 text-bold')
                        ui.image(str(img_path)).classes('w-full')
                    copy_chart_btn.enable()
                    ui.notify('周报图表已生成！', type='positive')
                except Exception as e:
                    queue.put_nowait(f'ERROR: {e}')
                    queue.put_nowait(None)
                    await drain_task
                    ui.notify(f'出错: {e}', type='negative')
                finally:
                    chart_btn.enable()

            async def copy_chart():
                path = chart_img_ref[0]
                if path is None:
                    ui.notify('请先生成图表', type='warning')
                    return
                data = base64.b64encode(path.read_bytes()).decode()
                await ui.run_javascript(f'''
                    const bytes = atob("{data}");
                    const arr = new Uint8Array(bytes.length);
                    for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
                    const blob = new Blob([arr], {{type: 'image/png'}});
                    await navigator.clipboard.write([new ClipboardItem({{'image/png': blob}})]);
                ''', timeout=15.0)
                ui.notify('图表已复制到剪贴板！', type='positive')

            chart_btn.on('click', on_chart)
            copy_chart_btn.on('click', copy_chart)

        def open_output():
            out = get_feature_dir(FEATURE) / 'output'
            out.mkdir(exist_ok=True)
            open_path(out)

        ui.button('打开输出目录', icon='folder_open', on_click=open_output).classes('q-mt-md').props('flat')
