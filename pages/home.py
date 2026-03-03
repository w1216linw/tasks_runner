from nicegui import ui

from components.layout import sidebar
from utils.paths import get_app_dir


def _clear_all_inputs() -> str:
    """删除 data/*/input/ 下的所有文件，返回摘要文字。"""
    removed = []
    for input_dir in sorted(get_app_dir().glob('*/input')):
        for f in input_dir.iterdir():
            if f.is_file():
                f.unlink()
                removed.append(f'{input_dir.parent.name}/{f.name}')
    return removed


def create() -> None:
    sidebar()

    with ui.column().classes('w-full q-pa-lg'):
        ui.label('TaskRunner').classes('text-h3 text-bold')
        ui.separator()
        ui.label('选择左侧功能开始使用。').classes('text-body1 q-mt-md text-grey-7')

        with ui.row().classes('q-mt-xl gap-md flex-wrap'):
            with ui.card().classes('cursor-pointer').on('click', lambda: ui.navigate.to('/process-counting-order')):
                ui.icon('table_chart', size='2rem').classes('text-primary')
                ui.label('处理计单').classes('text-subtitle1 text-bold')
                ui.label('TEMU + YY 订单去重 → 按州统计').classes('text-caption text-grey-7')

            with ui.card().classes('cursor-pointer').on('click', lambda: ui.navigate.to('/generate-weekly-trends')):
                ui.icon('show_chart', size='2rem').classes('text-primary')
                ui.label('生成周趋势图').classes('text-subtitle1 text-bold')
                ui.label('订单 CSV → 每日趋势折线图').classes('text-caption text-grey-7')

            with ui.card().classes('cursor-pointer').on('click', lambda: ui.navigate.to('/weekly-order-scan')):
                ui.icon('bar_chart', size='2rem').classes('text-primary')
                ui.label('周纬度订单扫描').classes('text-subtitle1 text-bold')
                ui.label('XLSX → 按日统计单量').classes('text-caption text-grey-7')

            with ui.card().classes('cursor-pointer').on('click', lambda: ui.navigate.to('/driver-week-analyze')):
                ui.icon('analytics', size='2rem').classes('text-primary')
                ui.label('司机周数据分析').classes('text-subtitle1 text-bold')
                ui.label('DWA 分析 + 周对周对比报表').classes('text-caption text-grey-7')

            with ui.card().classes('cursor-pointer').on('click', lambda: ui.navigate.to('/driver-missions')):
                ui.icon('person_pin', size='2rem').classes('text-primary')
                ui.label('司机任务点数').classes('text-subtitle1 text-bold')
                ui.label('按司机 + 日期统计任务点数').classes('text-caption text-grey-7')

        ui.separator().classes('q-mt-xl')

        def on_clear_click():
            with ui.dialog() as dlg, ui.card():
                ui.label('确认清理').classes('text-subtitle1 text-bold')
                ui.label('将删除所有功能 input/ 目录下的文件，此操作不可撤销。').classes('text-body2')
                with ui.row().classes('q-mt-md gap-sm justify-end w-full'):
                    ui.button('取消', on_click=dlg.close).props('flat')
                    def confirm():
                        dlg.close()
                        removed = _clear_all_inputs()
                        if removed:
                            ui.notify(f'已删除 {len(removed)} 个文件', type='positive')
                        else:
                            ui.notify('input 目录中没有文件', type='info')
                    ui.button('确认删除', on_click=confirm).props('color=negative')
            dlg.open()

        ui.button('清理所有输入文件', icon='delete_sweep', on_click=on_clear_click).props('flat color=negative').classes('q-mt-md')
