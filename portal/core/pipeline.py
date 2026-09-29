"""portal/core/pipeline.py — 流程编排。

主入口:
  process(project_dir, phase='ask' | 'validate' | 'refresh')
    phase=ask: 读 config + records → 跑 BO → 写 .cache
    phase=validate: 只校验 config,不动 .cache
    phase=refresh: 拿 .cache 最新文件合并到主 xlsx 的 📊 看板

⚠️ 时序设计:
  用户按 [🚀 启动 BO] 按钮  → run.py called → portal.run_process(phase=ask)
  用户按 [🔄 刷新最新推荐] → pipeline.run_process(phase=refresh)
  watchgod 监听到 mtime 变化 → pipeline.run_process(phase=validate)
"""
from __future__ import annotations

import json
import time
import traceback
from datetime import datetime
from pathlib import Path

from . import excel_io
from . import strategy_runner
from . import recommender


PHASE_ASK = "ask"
PHASE_VALIDATE = "validate"
PHASE_REFRESH = "refresh"


def _project_xlsx(project_dir: Path) -> Path:
    """从项目目录找唯一的 .xlsx 主文件(排除 .cache 和隐藏文件)。"""
    for f in project_dir.glob("*.xlsx"):
        if not f.name.startswith("~"):
            return f
    raise FileNotFoundError(f"项目目录里没找到 .xlsx: {project_dir}")


def process(project_dir: Path, phase: str) -> dict:
    """主流程入口。返回 dict 让上层(watcher / run.py)读 status。"""
    project_dir = project_dir.resolve()
    log_path = project_dir / "LOG.txt"
    started = time.time()
    log = []

    def _log(msg: str):
        line = f"[{datetime.now().isoformat(timespec='seconds')}] {msg}"
        log.append(line)
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(line + "\n")
        except Exception:
            pass

    try:
        xlsx_path = _project_xlsx(project_dir)
        _log(f"phase={phase} project={project_dir.name} xlsx={xlsx_path.name}")

        if phase == PHASE_VALIDATE:
            cfg = excel_io.read_config(xlsx_path)
            _log(f"validate OK: {len(cfg.factors)} factors, {len(cfg.outputs)} outputs")
            excel_io.write_status(
                xlsx_path,
                {
                    "state": "awaiting_bo",
                    "step": "1/4",
                    "next": "点 [🚀 启动 BO]",
                    "project_name": project_dir.name,
                    "project_dir": str(project_dir),
                    "round": 0,
                    "n_data": 0,
                    "last_bo_at": "",
                    "convergence_msg": "⏳ 暂无收敛数据",
                    "error": "",
                },
            )
            return {"ok": True, "phase": phase, "elapsed_s": time.time() - started}

        if phase == PHASE_ASK:
            cfg = excel_io.read_config(xlsx_path)
            factor_keys = [f.key for f in cfg.factors]
            output_keys = [o.key for o in cfg.outputs]
            records = excel_io.read_records(xlsx_path, factor_keys, output_keys)
            valid_records = [r for r in records if r.valid == 1]

            _log(f"config OK, n_records={len(records)}, n_valid={len(valid_records)}")

            # 算 round / next_trial_id
            round_num = 1
            next_trial_id = 1
            existing_tids = sorted(r.trial_id for r in records)
            if existing_tids:
                # 从 T0001 这种找最大编号
                max_n = 0
                for tid in existing_tids:
                    if tid.startswith("T") and tid[1:].isdigit():
                        max_n = max(max_n, int(tid[1:]))
                next_trial_id = max_n + 1
                # round = next_trial_id / batch_size
                round_num = max(1, max_n // max(1, cfg.strategy.batch_size) + 1)

            # 跑 BO
            candidates = strategy_runner.run_ask(
                cfg,
                valid_records,
                next_trial_id_start=next_trial_id,
            )
            # workaround #7: pin 非 trainable 因子
            candidates = strategy_runner.pin_non_trainable(candidates, records, cfg)

            _log(f"BO produced {len(candidates)} candidates")

            # 写 .cache
            cache_path = recommender.write_recommend_cache(
                project_dir,
                candidates,
                factor_keys,
                output_keys,
                round_num,
            )
            _log(f"cache written: {cache_path.name}")

            # 更新 _status
            excel_io.write_status(
                xlsx_path,
                {
                    "state": "bo_done",
                    "step": "4/4",
                    "next": "点 [🔄 刷新最新推荐] 把推荐写到 📊 看板",
                    "project_name": project_dir.name,
                    "project_dir": str(project_dir),
                    "round": round_num,
                    "n_data": len(valid_records),
                    "last_bo_at": datetime.now().isoformat(timespec="seconds"),
                    "convergence_msg": "⏳ 暂无收敛数据" if len(valid_records) < cfg.strategy.n_initial else "✅ 收敛评估中",
                    "error": "",
                },
            )
            return {
                "ok": True,
                "phase": phase,
                "elapsed_s": time.time() - started,
                "n_candidates": len(candidates),
                "cache": str(cache_path),
            }

        if phase == PHASE_REFRESH:
            latest = recommender.find_latest_cache(project_dir)
            if latest is None:
                _log("refresh: no cache file")
                return {"ok": False, "phase": phase, "error": "no cache file"}
            # 算 round
            existing = excel_io.read_records(xlsx_path, [f.key for f in excel_io.read_config(xlsx_path).factors], [o.key for o in excel_io.read_config(xlsx_path).outputs])
            cfg = excel_io.read_config(xlsx_path)
            round_num = (len(existing) // max(1, cfg.strategy.batch_size)) + 1
            recommender.merge_cache_to_dashboard(project_dir, xlsx_path, latest, round_num)
            _log(f"refreshed dashboard from {latest.name}")
            return {"ok": True, "phase": phase, "elapsed_s": time.time() - started, "from": latest.name}

        return {"ok": False, "phase": phase, "error": f"unknown phase {phase!r}"}

    except Exception as e:
        tb = traceback.format_exc()
        _log(f"ERROR phase={phase}: {e}\n{tb}")
        # 写 status 错误
        try:
            excel_io.write_status(
                _project_xlsx(project_dir),
                {
                    "state": "error",
                    "error": str(e)[:200],
                },
            )
        except Exception:
            pass
        return {"ok": False, "phase": phase, "error": str(e), "elapsed_s": time.time() - started}


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--project", required=True, type=Path)
    p.add_argument("--phase", default=PHASE_ASK, choices=[PHASE_ASK, PHASE_VALIDATE, PHASE_REFRESH])
    args = p.parse_args()
    result = process(args.project, args.phase)
    print(json.dumps(result, ensure_ascii=False, indent=2))