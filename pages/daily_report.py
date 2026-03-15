"""日报制作器页面"""

from __future__ import annotations

import asyncio
import importlib.util
import json
from datetime import date, datetime
from pathlib import Path
from typing import Callable

from nicegui import run, ui

from components.layout import back_button, sidebar
from utils.clipboard import copy_image_to_clipboard, copy_text_to_clipboard
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

    _stats_file = get_feature_dir('daily_report') / 'step_stats.json'
    step_stats: dict[str, dict] = {}
    try:
        if _stats_file.exists():
            step_stats.update(json.loads(_stats_file.read_text(encoding='utf-8')))
    except Exception:
        pass

    step_icon_els: dict[str, ui.icon] = {}
    step_btn_els: dict[str, ui.button] = {}
    step_output_els: dict[str, ui.column] = {}
    md_path_ref: list[Path | None] = [None]

    with ui.column().classes('w-full q-pa-lg gap-md'):
        with ui.row().classes('items-center gap-sm'):
            back_button()
            ui.label('日报').classes('text-h4 text-bold')
        ui.separator()

        # 日期 + 操作按钮
        with ui.dialog() as date_dialog, ui.card().classes('p-0'):
            date_picker = ui.date(value=date.today().isoformat())

        with ui.row().classes('gap-sm q-mt-md flex-wrap'):
            with ui.row().classes('items-center'):
                date_display = ui.label(date.today().isoformat()).classes('text-body1 text-bold')
                ui.button('更改', icon='edit', on_click=date_dialog.open).props('flat dense color=grey-7')
                ui.separator().props('vertical').classes('self-stretch mx-1')
                run_all_btn = ui.button('运行全部', icon='play_arrow').props('color=primary')
                gen_md_btn = ui.button('生成日报', icon='description').props('color=secondary')
            with ui.row().classes('items.center'):
                ui.button('打开目录', icon='folder',
                      on_click=lambda: open_path(get_feature_dir('daily_report'))).props('flat color=grey-8')
                ui.button('打开输出目录', icon='folder_open',
                      on_click=lambda: open_path(_get_dirs()[1])).props('flat color=grey-8')
                ui.button('清理input', icon='delete_sweep',
                      on_click=lambda: _confirm_clear()).props('flat color=negative')

        def _on_date_pick():
            date_display.set_text(date_picker.value or date.today().isoformat())
            date_dialog.close()
        date_picker.on('update:modelValue', lambda _: _on_date_pick())

        ui.separator()

        # 预约未达成
        with ui.row().classes('items-center gap-sm q-mt-xs'):
            reach_zero_btn = ui.button('生成预约未达成', icon='search').props('flat color=grey-8')
        reach_zero_result_row = ui.row().classes('items-center gap-sm q-pl-xs q-py-xs')

        ui.separator()

        # 步骤列表
        for step in STEPS:
            with ui.column().classes('w-full gap-0'):
                with ui.row().classes('items-center gap-sm w-full'):
                    icon_el = ui.icon(STATUS_ICON['pending'][0], size='sm').classes(STATUS_ICON['pending'][1])
                    ui.label(step['label']).classes('text-body1 flex-1')
                    btn = ui.button('单独运行', icon='play_arrow').props('flat dense color=primary')
                    step_icon_els[step['id']] = icon_el
                    step_btn_els[step['id']] = btn
                step_output_els[step['id']] = ui.column().classes('w-full')

        ui.separator()

        # 日志区
        log_area = ui.log(max_lines=300).classes('w-full h-64 font-mono text-xs')

        ui.separator()

        # 日报状态栏
        with ui.row().classes('items-center gap-sm'):
            md_status_label = ui.label('日报.md：未生成').classes('text-caption text-grey-6')
            open_md_btn = ui.button('打开文件', icon='open_in_new').props('flat dense')
            open_md_btn.visible = False


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

    async def _copy_image(path: Path):
        await copy_image_to_clipboard(path)

    def _load_reach_zero_ids() -> list[str]:
        import pandas as pd
        data_dir = get_feature_dir('daily_report')
        module_dir = data_dir / 'input' / 'yy_pickup_rate'
        reach_file = None
        for f in sorted(module_dir.glob('*.csv')):
            cols = pd.read_csv(f, encoding='utf-16', sep='\t', nrows=0).columns
            if '揽收达成率' in cols:
                reach_file = f
                break
        if reach_file is None:
            raise FileNotFoundError(f'未在 {module_dir} 找到含「揽收达成率」列的 CSV')
        df = pd.read_csv(reach_file, encoding='utf-16', sep='\t')
        def to_float_percent(x):
            if isinstance(x, str) and '%' in x:
                return float(x.replace('%', '')) / 100
            return float(x) if pd.notnull(x) else 0.0
        df['揽收达成率'] = df['揽收达成率'].apply(to_float_percent)
        return df[df['揽收达成率'] == 0.0]['单据号'].astype(str).tolist()

    async def _on_reach_zero_click():
        reach_zero_btn.disable()
        try:
            ids: list[str] = await run.io_bound(_load_reach_zero_ids)
            reach_zero_result_row.clear()
            with reach_zero_result_row:
                ui.label(f'reach_zero: {len(ids)} 单').classes('text-body2 text-grey-8')
                ui.separator().props('vertical').classes('self-stretch mx-1')
                async def _copy_partial(id_list=ids):
                    await copy_text_to_clipboard('\n'.join(id_list[:950]))
                async def _copy_all(id_list=ids):
                    await copy_text_to_clipboard('\n'.join(id_list))
                ui.button('前950单', icon='content_copy', on_click=_copy_partial).props('flat dense size=xs color=grey')
                ui.button('全部', icon='content_copy', on_click=_copy_all).props('flat dense size=xs color=grey')
        except Exception as e:
            ui.notify(f'读取失败: {e}', type='negative')
        finally:
            reach_zero_btn.enable()

    def _get_step_result_text(step_id: str) -> str | None:
        all_stats: dict = {}
        for v in step_stats.values():
            all_stats.update(v)
        gen = _load_report_generator()
        return gen.get_step_results(_get_today(), all_stats).get(step_id)

    def _show_step_outputs(step_id: str, result: dict):
        container = step_output_els.get(step_id)
        if container is None:
            return
        images = [Path(v) for v in result.values() if isinstance(v, str) and v.endswith('.png')]
        result_text = _get_step_result_text(step_id)
        if not images and not result_text:
            return
        container.clear()
        with container:
            if result_text:
                ui.markdown(result_text).classes('text-body2 q-pl-lg q-py-xs text-grey-8')
            for img_path in images:
                with ui.row().classes('items-center gap-sm q-pl-lg q-py-xs'):
                    ui.icon('image', size='xs').classes('text-grey-5')
                    ui.label(img_path.name).classes('text-caption text-grey-7 flex-1 font-mono')
                    ui.button(
                        icon='content_copy',
                        on_click=lambda p=img_path: _copy_image(p),
                    ).props('flat dense size=xs color=grey')

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
            try:
                _stats_file.write_text(json.dumps(step_stats), encoding='utf-8')
            except Exception:
                pass
            _set_status(step_id, 'success')
            _show_step_outputs(step_id, result)
            log(f'[{step_id.upper()}] 完成 ✓')
        except Exception as e:
            import traceback
            _set_status(step_id, 'error')
            log(f'[{step_id.upper()}] 失败: {e}')
            for line in traceback.format_exc().splitlines():
                log(f'  {line}')
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
        step_stats.clear()
        try:
            _stats_file.unlink(missing_ok=True)
        except Exception:
            pass
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
    reach_zero_btn.on('click', _on_reach_zero_click)

    for step in STEPS:
        sid, mod = step['id'], step['module']
        step_btn_els[sid].on('click', lambda s=sid, m=mod: asyncio.create_task(_execute_step(s, m)))

    # 从持久化数据恢复状态图标和图片行
    for step_id, result in step_stats.items():
        _set_status(step_id, 'success')
        _show_step_outputs(step_id, result)
