from pathlib import Path
import sys


def get_app_dir() -> Path:
    """返回 app 数据目录。打包后在 exe 同级的 app/，开发时用 data/。"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent / 'app'
    return Path(__file__).resolve().parent.parent / 'data'


def get_feature_dir(feature: str) -> Path:
    """返回某个功能的数据目录，并确保子目录存在。"""
    base = get_app_dir() / feature
    base.mkdir(parents=True, exist_ok=True)
    return base
