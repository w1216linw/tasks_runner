from nicegui import ui

from components.layout import sidebar


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
                ui.label('每日订单量扫描').classes('text-subtitle1 text-bold')
                ui.label('XLSX → 按日统计单量').classes('text-caption text-grey-7')

            with ui.card().classes('cursor-pointer').on('click', lambda: ui.navigate.to('/driver-week-analyze')):
                ui.icon('analytics', size='2rem').classes('text-primary')
                ui.label('司机周数据分析').classes('text-subtitle1 text-bold')
                ui.label('DWA 分析 + 周对周对比报表').classes('text-caption text-grey-7')

            with ui.card().classes('cursor-pointer').on('click', lambda: ui.navigate.to('/driver-missions')):
                ui.icon('person_pin', size='2rem').classes('text-primary')
                ui.label('司机任务点数').classes('text-subtitle1 text-bold')
                ui.label('按司机 + 日期统计任务点数').classes('text-caption text-grey-7')
