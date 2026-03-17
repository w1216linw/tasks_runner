"""司机周数据分析页面"""

from __future__ import annotations

import asyncio
from pathlib import Path

from nicegui import run, ui

from components.layout import back_button, sidebar
from scripts.driver_week_analyze import run_comparison, run_dwa, run_weekly_chart
from utils.clipboard import copy_image_to_clipboard
from utils.paths import get_feature_dir, open_path

FEATURE = 'driver_week_analyze'


def _scan(subdir: str, pattern: str) -> list[Path]:
    d = get_feature_dir(FEATURE) / subdir
    d.mkdir(parents=True, exist_ok=True)
    return sorted(d.glob(pattern))


def _scan_prefix(subdir: str, prefix: str) -> list[Path]:
    """大小写不敏感地匹配指定前缀的 .xlsx 文件。"""
    d = get_feature_dir(FEATURE) / subdir
    d.mkdir(parents=True, exist_ok=True)
    p = prefix.lower()
    return sorted(f for f in d.glob('*.xlsx') if f.name.lower().startswith(p))


def _file_select(files: list, selected_path: str) -> 'ui.select':
    """文件选择器：label 显示「文件名  修改日期」。"""
    from datetime import datetime
    opts = {
        str(f): f'{f.name}  [{datetime.fromtimestamp(f.stat().st_mtime).strftime("%m/%d %H:%M")}]'
        for f in files
    }
    return ui.select(opts, value=selected_path).classes('w-full')


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
                '司机数据 dwd*.xlsx  +  加油交易 transaction*.xlsx'
            ).classes('text-caption text-grey-7')

            dwd_files = _scan_prefix('input', 'dwd')
            txn_files = _scan_prefix('input', 'transaction')

            with ui.row().classes('gap-md w-full q-mt-sm'):
                with ui.column().classes('flex-1'):
                    ui.label('司机数据 (dwd*.xlsx)').classes('text-caption text-grey-7')
                    if not dwd_files:
                        ui.label('未找到').classes('text-caption text-negative')
                        dwd_select = None
                    else:
                        dwd_select = _file_select(dwd_files, str(dwd_files[0]))

                with ui.column().classes('flex-1'):
                    ui.label('加油交易 (transaction*.xlsx)').classes('text-caption text-grey-7')
                    if not txn_files:
                        ui.label('未找到').classes('text-caption text-negative')
                        txn_select = None
                    else:
                        txn_select = _file_select(txn_files, str(txn_files[0]))

            def _confirm_clear():
                with ui.dialog() as dlg, ui.card():
                    ui.label('确认清理').classes('text-subtitle1 text-bold')
                    ui.label('将删除 input/ 下所有文件，此操作不可撤销。').classes('text-body2')
                    with ui.row().classes('q-mt-md gap-sm justify-end w-full'):
                        ui.button('取消', on_click=dlg.close).props('flat')
                        def confirm():
                            dlg.close()
                            input_dir = get_feature_dir(FEATURE) / 'input'
                            count = 0
                            if input_dir.exists():
                                for f in input_dir.iterdir():
                                    if f.is_file():
                                        f.unlink()
                                        count += 1
                            ui.notify(f'已删除 {count} 个文件' if count > 0 else 'input 目录中没有文件',
                                      type='positive' if count > 0 else 'info')
                        ui.button('确认删除', on_click=confirm).props('color=negative')
                dlg.open()

            with ui.row().classes('q-mt-sm gap-sm items-center'):
                dwa_btn = ui.button('生成 DWA 分析', icon='play_arrow')
                ui.button('清理input', icon='delete_sweep', on_click=_confirm_clear).props('flat color=negative')
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
                    _render_output_selectors()
                    ui.notify('DWA 分析完成！', type='positive')
                except Exception as e:
                    queue.put_nowait(f'ERROR: {e}')
                    queue.put_nowait(None)
                    await drain_task
                    ui.notify(f'出错: {e}', type='negative')
                finally:
                    dwa_btn.enable()

            dwa_btn.on('click', on_dwa)

        # ── 步骤 2/3 共用：output 文件选择器（可刷新）─────────
        prev_ref: list = [None]
        curr_ref: list = [None]
        chart_prev_ref: list = [None]
        chart_curr_ref: list = [None]

        # ── 第二步: 周对周对比 ────────────────────────────
        with ui.card().classes('w-full q-mt-md'):
            with ui.row().classes('items-center gap-sm'):
                ui.label('第二步：生成周对周对比报表').classes('text-subtitle1 text-bold flex-1')
                refresh_btn = ui.button(icon='refresh').props('flat dense round')

            step2_row = ui.row().classes('gap-md w-full q-mt-sm')

            cmp_btn = ui.button('生成对比报表', icon='compare_arrows').classes('q-mt-sm')
            cmp_log = ui.log(max_lines=100).classes('w-full h-40 q-mt-sm font-mono text-xs')

            async def on_comparison():
                if prev_ref[0] is None or curr_ref[0] is None:
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
                        Path(prev_ref[0].value),
                        Path(curr_ref[0].value),
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

            step3_row = ui.row().classes('gap-md w-full q-mt-sm')

            chart_img_ref: list[Path | None] = [None]

            with ui.row().classes('q-mt-sm gap-sm items-center'):
                chart_btn = ui.button('生成周报图表', icon='bar_chart')
                copy_chart_btn = ui.button('复制图表', icon='content_copy').props('flat')
                copy_chart_btn.disable()

            chart_log = ui.log(max_lines=50).classes('w-full h-32 q-mt-sm font-mono text-xs')
            chart_img = ui.column().classes('w-full q-mt-md')

            async def on_chart():
                if chart_prev_ref[0] is None or chart_curr_ref[0] is None:
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
                        Path(chart_prev_ref[0].value),
                        Path(chart_curr_ref[0].value),
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
                await copy_image_to_clipboard(path)

            chart_btn.on('click', on_chart)
            copy_chart_btn.on('click', copy_chart)

        def _render_output_selectors():
            files = _scan('output', 'dwa_*.xlsx')

            step2_row.clear()
            with step2_row:
                with ui.column().classes('flex-1'):
                    ui.label('上周 dwa_*.xlsx').classes('text-caption text-grey-7')
                    if len(files) < 2:
                        ui.label('请先完成两周的 DWA 分析').classes('text-caption text-negative')
                        prev_ref[0] = None
                        curr_ref[0] = None
                    else:
                        prev_ref[0] = _file_select(files, str(files[0]))
                if len(files) >= 2:
                    with ui.column().classes('flex-1'):
                        ui.label('本周 dwa_*.xlsx').classes('text-caption text-grey-7')
                        curr_ref[0] = _file_select(files, str(files[-1]))

            step3_row.clear()
            with step3_row:
                with ui.column().classes('flex-1'):
                    ui.label('上周 dwa_*.xlsx').classes('text-caption text-grey-7')
                    if len(files) < 2:
                        ui.label('请先完成两周的 DWA 分析').classes('text-caption text-negative')
                        chart_prev_ref[0] = None
                        chart_curr_ref[0] = None
                    else:
                        chart_prev_ref[0] = _file_select(files, str(files[0]))
                if len(files) >= 2:
                    with ui.column().classes('flex-1'):
                        ui.label('本周 dwa_*.xlsx').classes('text-caption text-grey-7')
                        chart_curr_ref[0] = _file_select(files, str(files[-1]))

        _render_output_selectors()
        refresh_btn.on('click', lambda: _render_output_selectors())

        def open_output():
            out = get_feature_dir(FEATURE) / 'output'
            out.mkdir(exist_ok=True)
            open_path(out)

        ui.button('打开输出目录', icon='folder_open', on_click=open_output).classes('q-mt-md').props('flat')
