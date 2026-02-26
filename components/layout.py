from nicegui import ui


NAV_ITEMS = [
    ('首页', '/'),
    ('处理计单', '/process-counting-order'),
    ('生成周趋势图', '/generate-weekly-trends'),
    ('每日订单量扫描', '/weekly-order-scan'),
    ('司机周数据分析', '/driver-week-analyze'),
    ('司机任务点数', '/driver-missions'),
]


def sidebar() -> None:
    """渲染左侧导航栏（在每个页面顶部调用）。"""
    with ui.left_drawer(fixed=True).classes('bg-grey-1 q-pa-md'):
        ui.label('TaskRunner').classes('text-h6 text-bold q-mb-md')
        for label, path in NAV_ITEMS:
            ui.link(label, path).classes('block q-py-xs text-body1')


def back_button(target: str = '/') -> None:
    """左上角返回按钮。"""
    ui.button(icon='arrow_back', on_click=lambda: ui.navigate.to(target)).props('flat round')
