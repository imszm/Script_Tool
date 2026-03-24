from __future__ import annotations

import runpy
import sys
from pathlib import Path


def run_legacy_script(script_path: str | Path, argv: list[str] | None = None) -> None:
    """
    在当前进程内执行历史脚本（保留其 __name__ == '__main__' 行为）。
    这能让 main.py 作为统一入口，而不需要立刻把所有脚本全部重写成模块。
    """
    path = Path(script_path).resolve()
    if not path.exists():
        raise FileNotFoundError(str(path))

    old_argv = sys.argv
    try:
        sys.argv = [str(path)] + (argv or [])
        runpy.run_path(str(path), run_name="__main__")
    finally:
        sys.argv = old_argv

