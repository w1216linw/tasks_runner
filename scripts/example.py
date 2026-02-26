"""示例脚本 — 展示如何使用 Excel 工具。"""

from pathlib import Path

from utils.excel import read_excel, write_excel

DATA_DIR = Path(__file__).resolve().parent.parent / "data"


def run() -> str:
    """生成示例 Excel 并读取，返回结果摘要。"""
    output_path = DATA_DIR / "example_output.xlsx"

    sample_data = [
        {"姓名": "张三", "年龄": 28, "城市": "北京"},
        {"姓名": "李四", "年龄": 34, "城市": "上海"},
        {"姓名": "王五", "年龄": 22, "城市": "广州"},
    ]

    write_excel(output_path, sample_data)
    rows = read_excel(output_path)
    return f"写入 {len(sample_data)} 行，读回 {len(rows)} 行 -> {output_path}"


if __name__ == "__main__":
    print(run())
