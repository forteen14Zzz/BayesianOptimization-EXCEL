"""portal/run.py — 命令行入口(给 .bat 用)。"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from portal.core import pipeline


def main():
    p = argparse.ArgumentParser(description="BoFire Lab Portal — single BO run")
    p.add_argument("--project", type=Path, required=True)
    p.add_argument("--phase", default="ask", choices=["ask", "validate", "refresh"])
    args = p.parse_args()
    if not args.project.exists():
        print(f"❌ 项目目录不存在: {args.project}")
        return 1
    result = pipeline.process(args.project, args.phase)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())