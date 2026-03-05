"""根据各模块统计结果生成日报.md"""

from datetime import datetime, timedelta
from pathlib import Path


def generate(today: datetime, stats: dict, output_dir: Path) -> Path:
    """
    stats: 各模块 run() 返回的 dict 合并后的结果
    Returns: 生成的 md 文件路径
    """
    yesterday = (today - timedelta(days=1)).strftime('%m-%d')
    last2day = (today - timedelta(days=2)).strftime('%m-%d')
    today_str = today.strftime('%Y-%m-%d')

    def s(key, default='?'):
        return stats.get(key, default)

    lines = [
        f'# ORD {today_str} 本地揽收',
        '',
        f'## 一、今日下单情况（{yesterday} 08:00 – {today_str} 08:00）',
        '',
        f'ORD TUBT 下单 **{s("total_orders_today")}** 单（含取消 {s("total_cancel")} 单）。',
        f'预约揽收 **{s("df_reserve")}** 单，去重后 **{s("df_reserve_clean")}** 单。',
        '',
        '---',
        '',
        f'## 二、{yesterday} 下单揽收情况',
        '',
        f'{yesterday} 当日 TUBT-ORD 共下单 **{s("t1_order_total")}** 单，截至当前已签入 **{s("t1_signed_in")}** 单，'
        f'签入率为 **{s("t1_signin_rate")}%**，仍有 **{s("t1_pending")}** 单处于已下单状态。',
        '',
        f'当日 TUBT-ORD 操作签入量为 **{s("tubt_operation_signin")}** 单，自打面单操作量为 **{s("sp_operation_count")}** 单。',
        '',
        f'已下单订单共 **{s("orders")}** 单，其中 **{s("reach_rate")}%** 的地址已触达。',
        '',
        '---',
        '',
        f'## 三、{last2day} SHEIN D2D 预约揽收',
        '',
        f'{last2day} SHEIN D2D 共接收预约订单 **{s("total_sn")}** 单，客户取消 {s("cancel_sn")} 单，'
        f'已签入 **{s("sign_in_sn")}** 单，达成率 **{s("sign_in_sn_percent")}%**，'
        f'仍有 **{s("total_yxd")}** 单处于已下单未达成状态。',
        '',
        '---',
        '',
        f'## 四、{last2day} 预约揽收达成率',
        '',
        f'{last2day} 预约揽收总体达成率为 **{s("reach_percent")}%**。',
        f'在已达成订单中，实际揽收签入 **{s("signin_count")}** 单，占达成订单的 **{s("signin_percent")}%**；'
        f'推送揽收失败 **{s("signin_fail_count")}** 单，占 **{s("fail_percent")}%**。',
        '',
        f'共有 **{s("df_reach_zero")}** 单揽收未达成，其中已触达地址且 Hub 班次已推送的订单仍计为未达成。',
    ]

    md_path = output_dir / '日报.md'
    md_path.write_text('\n'.join(lines), encoding='utf-8')
    return md_path
