"""跨平台图片复制到剪贴板工具。"""

from __future__ import annotations

import base64
import subprocess
import sys
from pathlib import Path

from nicegui import run, ui


async def copy_image_to_clipboard(path: Path) -> None:
    """将 PNG 文件复制到系统剪贴板，成功后弹出提示。"""
    try:
        if sys.platform == 'darwin':
            script = f'set the clipboard to (read (POSIX file "{path}") as «class PNGf»)'
            await run.io_bound(
                subprocess.run, ['osascript', '-e', script],
                capture_output=True, check=True,
            )
        else:
            data = base64.b64encode(path.read_bytes()).decode()
            await ui.run_javascript(f'''
                const blob = await fetch('data:image/png;base64,{data}').then(r=>r.blob());
                await navigator.clipboard.write([new ClipboardItem({{'image/png':blob}})]);
            ''', timeout=60.0)
        ui.notify('已复制到剪贴板 ✓', type='positive')
    except Exception as e:
        ui.notify(f'复制失败: {e}', type='negative')
