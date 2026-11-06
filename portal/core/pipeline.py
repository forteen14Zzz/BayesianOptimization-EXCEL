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

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from . import excel_io
from . import strategy_runner
from . import recommender
from . import sync_headers


PHASE_ASK = "ask"
PHASE_VALIDATE = "validate"
PHASE_REFRESH = "refresh"


# ---------- 🧪 记录自动填充 ----------
def _fill_candidates_to_records(
    xlsx_path: Path,
    cache_path: Path,
    candidates: list,
    factor_keys: list[str],
    output_keys: list[str],
) -> int:
    """把 ask 阶段的候选实验条件自动填到主 xlsx 的 🧪 记录 sheet。

    用户做完实验后,直接填 Yield 列,不用手动抄推荐条件。

    设计:
    - 🧪 记录 row 4 表头 = trial_id, batch_id, 因子..., yield, valid, status, 来源, 录入时间
      因子列顺序 = factor_keys 顺序(从 ⚙️ 配置读)
    - 每个候选写一行:Yield 列留空,valid=1,status=pending,来源=第N轮推荐
    - 找第一个空行追加(不覆盖已有数据)
    - 用 atomic write (同 merge_cache_to_dashboard_safe)
    """
    if not candidates:
        return 0

    tmp_path = xlsx_path.with_suffix(xlsx_path.suffix + ".tmp")
    wb = load_workbook(str(xlsx_path))
    try:
        if "🧪 记录" not in wb.sheetnames:
            return 0
        ws = wb["🧪 记录"]

        # 同步 row 4 表头 — 用 cache headers 顺序
        # 列顺序: trial_id, batch_id, 因子..., yield, valid, status, 来源, 录入时间
        expected_headers = ["trial_id", "batch_id"] + list(factor_keys) + list(output_keys) + ["valid", "status", "来源", "录入时间"]
        for j, h in enumerate(expected_headers):
            target = ws.cell(row=4, column=j + 1)
            if isinstance(target, MergedCell):
                continue
            target.value = h

        # 找第一个空行(从 row 5 起)
        # ⚠️ 跳过模板默认的"💡 提示"合并行(通常是 row 14 / A14:K14)
        next_row = 5
        for r in range(5, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if v is None or (isinstance(v, str) and not v.strip()):
                # 跳过提示行
                if isinstance(v, str) and (v.startswith("💡") or v.startswith("📌") or v.startswith("提示")):
                    continue
                break
            # 跳过合并区子格(A14:K14 那种被锁的)
            if isinstance(v, MergedCell):
                continue
            next_row = r + 1

        # 检查 next_row 是不是在合并区 — 如果是,把数据往下挪到合并区之后
        # 找出 next_row 之后第一个不在合并区的主格行
        max_merged_end = 0
        for rng in ws.merged_cells.ranges:
            if rng.min_row <= next_row <= rng.max_row:
                max_merged_end = max(max_merged_end, rng.max_row)
        if max_merged_end > 0:
            next_row = max_merged_end + 1

        # ⚠️ 删除 row 14 模板默认的"💡 提示"行(和实际表头冲突)
        # 同事只看到 4 个真实列,不需要这个老提示
        from openpyxl.cell.cell import MergedCell as _MC2
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row == 14 and "💡" in str(ws.cell(row=14, column=1).value or ""):
                ws.unmerge_cells(str(rng))
        # 清 row 14 内容
        for c in range(1, 13):
            cell = ws.cell(row=14, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = None

        # 写每个候选一行
        now = datetime.now().isoformat(timespec="seconds")
        batch_id = datetime.now().strftime("B%Y%m%d-%H%M%S")
        n_written = 0
        for cand in candidates:
            row = [cand.trial_id, batch_id]
            for k in factor_keys:
                row.append(cand.factors.get(k))
            for o in output_keys:
                row.append(None)  # 用户做完实验填
            row.extend([1, "pending", "BO推荐", now])
            for j, v in enumerate(row):
                target = ws.cell(row=next_row, column=j + 1)
                if isinstance(target, MergedCell):
                    continue
                target.value = v
            next_row += 1
            n_written += 1

        # atomic write
        import shutil as _sh
        import time as _t
        last_err = None
        for attempt in range(3):
            try:
                wb.save(str(tmp_path))
                _sh.move(str(tmp_path), str(xlsx_path))
                return n_written
            except PermissionError as e:
                last_err = e
                _t.sleep(0.5)
                continue
        # 失败 — 不抛错(只是没填成功,用户可手动)
        _log_safe(f"fill records failed after 3 retries: {last_err}")
        return 0
    finally:
        try:
            wb.close()
        except Exception:
            pass
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass


def _log_safe(msg: str) -> None:
    """最小的 logger — 不依赖 LOG.txt(避免循环 import)。"""
    try:
        from pathlib import Path as _P
        log_path = _P(__file__).resolve().parent.parent.parent / "projects" / "process.log"
        # 简化: 写到 stdlib
        import logging as _l
        _l.warning(msg)
    except Exception:
        pass


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

            # ⚠️ 先同步 🧪 记录 的表头(跟随 ⚙️ 配置)
            # 否则同事填的列顺序和配置对不上 → BO 学到错乱的关系
            try:
                sync_result = sync_headers.sync_record_headers(
                    xlsx_path, factor_keys, output_keys
                )
                if sync_result.get("changed"):
                    _log(
                        f"record headers synced: {sync_result['old_headers']} "
                        f"-> {sync_result['new_headers']} "
                        f"(migrated {sync_result['migrated']} rows)"
                    )
            except Exception as e:
                _log(f"header sync failed (continue anyway): {e}")

            records = excel_io.read_records(
                xlsx_path, factor_keys, output_keys, factor_specs=cfg.factors
            )
            valid_records = [r for r in records if r.valid == 1]

            _log(f"config OK, n_records={len(records)}, n_valid={len(valid_records)}")

            # 算 round / next_trial_id
            round_num = 1
            next_trial_id = 1
            max_n = 0  # ⚠️ 必须在 if 外面初始化,否则 records 为空时 UnboundLocalError
            # 只算 valid=1 的 max trial_id,避免和失败行冲突
            valid_tids = sorted(r.trial_id for r in valid_records)
            if valid_tids:
                for tid in valid_tids:
                    if tid.startswith("T") and tid[1:].isdigit():
                        max_n = max(max_n, int(tid[1:]))
                next_trial_id = max_n + 1

            # ⚠️ round_num 从 .cache 文件数算(累积的),不是从 trial_id // batch_size 算
            # 原因:用户可能调整 batch_size,或者 GP 重跑,导致 trial_id 算不准
            cache_dir = project_dir / ".cache"
            if cache_dir.exists():
                cache_count = len(list(cache_dir.glob("recommend_*.xlsx")))
                round_num = max(1, cache_count + 1)
            else:
                # 没 cache 文件,就是第 1 轮
                round_num = 1

            # ⚠️ trial_id 必须全局唯一:读 .cache 里所有 cache 文件,
            # 找 max trial_id,作为下次起点。否则同事不填回 🧪 记录
            # 时,trial_id 永远从 T0006 开始(冲突)。
            if cache_dir.exists():
                global_max = max_n
                for cache_file in cache_dir.glob("recommend_*.xlsx"):
                    cwb = __import__("openpyxl").load_workbook(
                        str(cache_file), read_only=True, data_only=True
                    )
                    for row in cwb.active.iter_rows(min_row=2, max_col=1, values_only=True):
                        if row and row[0] and str(row[0]).startswith("T") and str(row[0])[1:].isdigit():
                            n = int(str(row[0])[1:])
                            global_max = max(global_max, n)
                    cwb.close()
                next_trial_id = max(next_trial_id, global_max + 1)

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

            # ⚠️ 自动填候选到 🧪 记录 — 用户做完实验直接填 yield
            # 用 cache headers 顺序写 row 4 表头(覆盖旧的)
            # 每个候选写一行:trial_id, batch_id, 因子..., yield(空), valid=1, status=pending
            n_filled = _fill_candidates_to_records(
                xlsx_path, cache_path, candidates,
                factor_keys=factor_keys, output_keys=output_keys,
            )
            _log(f"filled {n_filled} candidates into 🧪 记录")

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
            # ⚠️ 用 safe 版本(atomic write,Excel 开着也能跑)
            result = recommender.merge_cache_to_dashboard_safe(project_dir, xlsx_path, latest, round_num)
            if not result.get("ok"):
                _log(f"refresh failed: {result.get('error')}")
                return {"ok": False, "phase": phase, "error": result.get("error"), "elapsed_s": time.time() - started}
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