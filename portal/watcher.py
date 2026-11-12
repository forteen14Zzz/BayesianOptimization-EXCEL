"""portal/watcher.py — watchdog 监听 Excel 文件变化。

设计:
- 监听 projects/*/*.xlsx(项目目录的 xlsx)
- 文件 mtime 变了 + md5 变了 → 触发 pipeline.process(phase=validate)
- 保留 BO 文件,只让 validate 阶段同步,真正 BO 用 [🚀 启动 BO] 手动

⚠️ Windows 注意:
- 忽略 `~$xxx.xlsx` 临时锁文件
- mtime 变了但 md5 一样 → 跳过(避免 Ctrl+S 无改动误触发)
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

from portal.core import excel_io, pipeline


class Handler(FileSystemEventHandler):
    def __init__(self, project_dir: Path):
        self.project_dir = project_dir.resolve()
        self.last_md5: dict[str, str] = {}
        self.cooldown = 1.0  # 秒,1 秒内的多次变更合并到一次
        self.last_triggered: dict[str, float] = {}

    def on_modified(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path).resolve()
        # 忽略临时文件 / cache / log
        name = path.name
        if name.startswith("~$"):
            return
        if name.startswith("."):
            return
        if not name.endswith(".xlsx"):
            return
        if not path.exists():
            return
        # 只关心项目目录里的 xlsx(避免跑到别处)
        if path.parent.resolve() != self.project_dir:
            return
        # cooldown: 1 秒内同文件只触发一次
        now = time.time()
        if now - self.last_triggered.get(str(path), 0) < self.cooldown:
            return
        # md5 过滤
        try:
            h = excel_io.file_md5(path)
        except Exception:
            return
        if self.last_md5.get(str(path)) == h:
            return  # 内容没变
        self.last_md5[str(path)] = h
        self.last_triggered[str(path)] = now

        print(f"[watcher] {path.name} changed, running validate")
        result = pipeline.process(self.project_dir, pipeline.PHASE_VALIDATE)
        print(f"[watcher] result: {result}")


def watch(project_dir: Path):
    project_dir = project_dir.resolve()
    print(f"[watcher] watching: {project_dir}")
    handler = Handler(project_dir)
    obs = Observer()
    obs.schedule(handler, str(project_dir), recursive=False)
    obs.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        obs.stop()
    obs.join()


def main():
    p = argparse.ArgumentParser(description="BoFire Lab Portal — file watcher")
    p.add_argument("--project", type=Path, required=True)
    args = p.parse_args()
    if not args.project.exists():
        print(f"❌ 项目目录不存在: {args.project}")
        return 1
    watch(args.project)
    return 0


if __name__ == "__main__":
    sys.exit(main())