"""跨平台图片复制到剪贴板工具。"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

from nicegui import run, ui


async def copy_text_to_clipboard(text: str) -> None:
    """将文本复制到系统剪贴板，成功后弹出提示。"""
    try:
        if sys.platform == 'darwin':
            await run.io_bound(
                subprocess.run, ['pbcopy'],
                input=text.encode('utf-8'), check=True,
            )
        elif sys.platform == 'win32':
            # clip.exe 需要 UTF-16 LE with BOM 才能正确处理中文/特殊字符
            await run.io_bound(
                subprocess.run, ['clip'],
                input=text.encode('utf-16'), check=True,
            )
        else:
            # Linux: 优先 xclip，fallback xsel
            try:
                await run.io_bound(
                    subprocess.run, ['xclip', '-selection', 'clipboard'],
                    input=text.encode('utf-8'), check=True,
                )
            except FileNotFoundError:
                await run.io_bound(
                    subprocess.run, ['xsel', '--clipboard', '--input'],
                    input=text.encode('utf-8'), check=True,
                )
        ui.notify('已复制到剪贴板 ✓', type='positive')
    except Exception as e:
        ui.notify(f'复制失败: {e}', type='negative')


async def copy_image_to_clipboard(path: Path) -> None:
    """将 PNG 文件复制到系统剪贴板，成功后弹出提示。"""
    try:
        if sys.platform == 'darwin':
            script = f'set the clipboard to (read (POSIX file "{path}") as «class PNGf»)'
            await run.io_bound(
                subprocess.run, ['osascript', '-e', script],
                capture_output=True, check=True,
            )
        elif sys.platform == 'win32':
            ps_script = (
                'Add-Type -Assembly System.Windows.Forms;'
                'Add-Type -Assembly System.Drawing;'
                f'[System.Windows.Forms.Clipboard]::SetImage('
                f'[System.Drawing.Image]::FromFile([System.IO.Path]::GetFullPath("{path}")))'
            )
            await run.io_bound(
                subprocess.run,
                ['powershell', '-NoProfile', '-NonInteractive', '-Command', ps_script],
                capture_output=True, check=True,
            )
        else:
            # Linux: xclip
            await run.io_bound(
                subprocess.run,
                ['xclip', '-selection', 'clipboard', '-t', 'image/png', '-i', str(path)],
                check=True,
            )
        ui.notify('已复制到剪贴板 ✓', type='positive')
    except Exception as e:
        ui.notify(f'复制失败: {e}', type='negative')
