from pathlib import Path
import subprocess
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


def get_base_dir() -> Path:
    """返回代码根目录。打包后为 _MEIPASS（解压临时目录），开发时为项目根目录。"""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS)  # type: ignore[attr-defined]
    return Path(__file__).resolve().parent.parent


def open_path(path: Path) -> None:
    """在系统文件管理器中打开文件或文件夹（跨平台）。"""
    if sys.platform == 'win32':
        import os
        os.startfile(str(path))
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', str(path)])
