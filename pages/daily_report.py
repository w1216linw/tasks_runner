"""日报制作器页面"""

from __future__ import annotations

import asyncio
import importlib.util
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from nicegui import run, ui

from components.layout import back_button, sidebar
from utils.paths import get_base_dir, get_feature_dir, open_path

SCRIPTS_DIR = get_base_dir() / 'scripts' / 'daily_report'

STEPS = [
    {'id': 'ab', 'label': '整合订单 + 更新揽收数据 (a+b)', 'module': 'run_ab'},
    {'id': 'h',  'label': '自打面单统计 (h)',               'module': 'run_h'},
    {'id': 'c',  'label': '每日下单趋势 (c)',               'module': 'run_c'},
    {'id': 'd',  'label': '已下单地址触达 (d)',              'module': 'run_d'},
    {'id': 'g',  'label': 'SHEIN D2D 下单 (g)',            'module': 'run_g'},
    {'id': 'j',  'label': '预约达成率 (j)',                  'module': 'run_j'},
    {'id': 'k',  'label': '揽收未达成分析 (k)',              'module': 'run_k'},
]

STATUS_ICON = {
    'pending': ('radio_button_unchecked', 'text-grey-5'),
    'running': ('hourglass_empty',        'text-orange'),
    'success': ('check_circle',           'text-positive'),
    'error':   ('cancel',                 'text-negative'),
}


def _load_module(module_name: str):
    spec = importlib.util.spec_from_file_location(
        module_name, SCRIPTS_DIR / 'modules' / f'{module_name}.py'
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _run_step(module_name: str, today: datetime, data_dir: Path, output_dir: Path, log: Callable) -> dict:
    return _load_module(module_name).run(today, data_dir, output_dir, log)


def _load_report_generator():
    spec = importlib.util.spec_from_file_location('report_generator', SCRIPTS_DIR / 'report_generator.py')
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def create() -> None:
    sidebar()

    step_stats: dict[str, dict] = {}
    step_icon_els: dict[str, ui.icon] = {}
    step_btn_els: dict[str, ui.button] = {}
    md_path_ref: list[Path | None] = [None]

    with ui.column().classes('w-full q-pa-lg gap-md'):
        with ui.row().classes('items-center gap-sm'):
            back_button()
            ui.label('日报制作器').classes('text-h4 text-bold')
        ui.separator()

        # 日期 + 操作按钮
        with ui.row().classes('items-center gap-md q-mt-md'):
            date_picker = ui.date(value=date.today().isoformat()).classes('w-40')
            run_all_btn = ui.button('运行全部', icon='play_arrow').props('color=primary')
            gen_md_btn = ui.button('生成日报', icon='description').props('color=secondary')

        ui.separator()

        # 步骤列表
        for step in STEPS:
            with ui.row().classes('items-center gap-sm w-full'):
                icon_el = ui.icon(STATUS_ICON['pending'][0], size='sm').classes(STATUS_ICON['pending'][1])
                ui.label(step['label']).classes('text-body1 flex-1')
                btn = ui.button('单独运行', icon='play_arrow').props('flat dense color=primary')
                step_icon_els[step['id']] = icon_el
                step_btn_els[step['id']] = btn

        ui.separator()

        # 日志区
        log_area = ui.log(max_lines=300).classes('w-full h-64 font-mono text-xs')

        ui.separator()

        # 日报状态栏
        with ui.row().classes('items-center gap-sm'):
            md_status_label = ui.label('日报.md：未生成').classes('text-caption text-grey-6')
            open_md_btn = ui.button('打开文件', icon='open_in_new').props('flat dense')
            open_md_btn.visible = False

        # 清理
        ui.button('清理 input 文件', icon='delete_sweep', on_click=lambda: _confirm_clear()).props('flat color=negative').classes('q-mt-sm')

    # ── 辅助 ────────────────────────────────────────────────────────────────

    def _get_today() -> datetime:
        d = date_picker.value
        if isinstance(d, str):
            d = date.fromisoformat(d)
        return datetime(d.year, d.month, d.day)

    def _get_dirs() -> tuple[Path, Path]:
        data_dir = get_feature_dir('daily_report')
        output_dir = data_dir / 'output' / _get_today().strftime('%Y-%m-%d')
        output_dir.mkdir(parents=True, exist_ok=True)
        return data_dir, output_dir

    def _set_status(step_id: str, status: str):
        icon_name, icon_class = STATUS_ICON[status]
        el = step_icon_els[step_id]
        el.name = icon_name
        el.classes(replace=icon_class)

    # ── 执行单个步骤 ─────────────────────────────────────────────────────────

    async def _execute_step(step_id: str, module_name: str):
        today = _get_today()
        data_dir, output_dir = _get_dirs()
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

        _set_status(step_id, 'running')
        step_btn_els[step_id].disable()
        drain_task = asyncio.create_task(drain())

        try:
            log(f'\n[{step_id.upper()}] 开始...')
            result = await run.io_bound(_run_step, module_name, today, data_dir, output_dir, log)
            step_stats[step_id] = result
            _set_status(step_id, 'success')
            log(f'[{step_id.upper()}] 完成 ✓')
        except Exception as e:
            _set_status(step_id, 'error')
            log(f'[{step_id.upper()}] 失败: {e}')
        finally:
            queue.put_nowait(None)
            await drain_task
            step_btn_els[step_id].enable()

    # ── 生成日报 ─────────────────────────────────────────────────────────────

    def _generate_md():
        today = _get_today()
        _, output_dir = _get_dirs()
        all_stats: dict = {}
        for v in step_stats.values():
            all_stats.update(v)

        gen = _load_report_generator()
        md_path = gen.generate(today, all_stats, output_dir)
        md_path_ref[0] = md_path
        md_status_label.set_text(f'日报.md：已生成 → {md_path.name}')
        md_status_label.classes(replace='text-caption text-positive')
        open_md_btn.visible = True
        ui.notify('日报.md 生成成功', type='positive')

    # ── 运行全部 ─────────────────────────────────────────────────────────────

    async def on_run_all():
        run_all_btn.disable()
        gen_md_btn.disable()
        for step in STEPS:
            await _execute_step(step['id'], step['module'])
        _generate_md()
        run_all_btn.enable()
        gen_md_btn.enable()

    # ── 清理 input ───────────────────────────────────────────────────────────

    def _do_clear() -> int:
        data_dir = get_feature_dir('daily_report')
        input_dir = data_dir / 'input'
        count = 0
        for f in input_dir.rglob('*'):
            if f.is_file():
                f.unlink()
                count += 1
        return count

    def _confirm_clear():
        with ui.dialog() as dlg, ui.card():
            ui.label('确认清理').classes('text-subtitle1 text-bold')
            ui.label('将删除 input/ 下所有文件，共享文件不受影响。此操作不可撤销。').classes('text-body2')
            with ui.row().classes('q-mt-md gap-sm justify-end w-full'):
                ui.button('取消', on_click=dlg.close).props('flat')
                def confirm():
                    dlg.close()
                    count = _do_clear()
                    ui.notify(f'已删除 {count} 个文件', type='positive' if count > 0 else 'info')
                ui.button('确认删除', on_click=confirm).props('color=negative')
        dlg.open()

    # ── 绑定事件 ─────────────────────────────────────────────────────────────

    run_all_btn.on('click', on_run_all)
    gen_md_btn.on('click', _generate_md)
    open_md_btn.on('click', lambda: open_path(md_path_ref[0]) if md_path_ref[0] else None)

    for step in STEPS:
        sid, mod = step['id'], step['module']
        step_btn_els[sid].on('click', lambda s=sid, m=mod: asyncio.create_task(_execute_step(s, m)))
