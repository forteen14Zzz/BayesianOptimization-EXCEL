# BoFire Lab Portal

> **代号**: BoFire Lab Portal (BLP)
> **定位**: 给有机化学实验室的"无代码"贝叶斯优化工具 — 一个 `.xlsx` + 后台 Python。
> **状态**: ✅ M1+M2 完成, M3 文档收尾

---

## 一句话产品形态

**同事用 Excel 配因子、跑完贴回结果,系统自动告诉你下一组实验该做什么条件。**

```
你改 Excel → watchdog 监听 → Python 跑 GP-BO → 新推荐写到 Excel
```

---

## 5 分钟快速体验

```bash
# 1. 装环境(首次需装 bofire,大约 5-10 分钟)
pip install bofire==0.4.1 botorch openpyxl pydantic pandas matplotlib watchdog cvxpy pymoo shap

# 2. 生成 demo 项目(2 个)
python scripts/build_demo_project.py        # Suzuki mock demo
python scripts/build_arylation_demo.py      # 真实 1728 条 Pd-芳基化反应

# 3. 跑 3-5 轮 BO(任选一个 demo)
python scripts/run_demo.py                  # Suzuki (mock)
python scripts/run_arylation_demo.py        # Arylation (真数据 + 真 GP-BO)
```

---

## 真实 benchmark — Arylation demo

**数据**: 来自 Princeton Shields + BMS 的公开反应优化数据集
[All_CH_arylation_Experiments_Jay_12052019.csv](https://github.com/connorcoley/reaction-optimization-game)
- 1728 条真实 Pd-催化直接芳基化实验
- 5 因子: base(4)/ ligand(12)/ solvent(4)/ temperature(80-130°C)/ concentration(0.05-0.2M)
- 目标: yield(0-100)

**M1 demo 跑 5 轮结果**(30 条种子 → GP 推荐 15 条):

| 轮 | 推荐组合 | 实测 yield |
|----|---------|----------|
| 1 | CsOAc/CgMe-PPh/BuOAc + ... | ~10-40% |
| 2 | KOPiv + X-Phos + DMAc @ T=122°C | **77.4%** |
| 3 | 探测 BuOAc/BuCN/DMAc 边界 | 0-10% |
| 4 | KOAc + CgMe-PPh + DMAc @ T=95°C | **65.0%** |
| 5 | KOAc + CgMe-PPh + DMAc @ T=80/120/130°C | 65.0% × 3(锁定组合 + 探边界)|

> **结论**: 真 BoFire GP-BO(qLogEI)在 30 条种子数据后,**5 轮内找到 KOAc+CgMe-PPh+DMAc 高产组合**(GP 文献高产组合之一)。每轮 ~30s,5 轮总共约 2.5 分钟。

---

## 项目结构

```
E:\bofire\
├── docs/                           # 设计文档
│   ├── PRODUCT_DESIGN.md           # 完整产品设计
│   ├── 同事使用手册.md               # 中文,用户向
│   ├── IT_装机手册.md              # 中文,IT 向
│   ├── BENCHMARK_ARYLATION.md      # 真数据 benchmark 报告
│   ├── FAQ.md                      # 常见问题
│   └── DESIGN_DECISIONS.md         # 设计决策日志
├── install/                        # IT 装机
│   ├── install.bat
│   ├── requirements.txt            # 锁版本
│   └── setup_autostart.bat
├── templates/
│   └── project_template.xlsx       # 4 sheet + 隐藏 _status + 公式
├── portal/                         # Python 后台(~1500 行)
│   ├── watcher.py                  # watchdog + md5 过滤
│   ├── run.py                      # CLI 入口
│   └── core/
│       ├── excel_io.py             # Excel 读写 + pydantic 校验
│       ├── domain_builder.py       # config → BoFire Domain
│       ├── strategy_runner.py      # ask (含降级路径)
│       ├── recommender.py          # .cache 写入 + 看板合并
│       └── pipeline.py             # 流程编排
├── projects/                       # 同事项目
│   ├── Suzuki_2026Q4_demo/         # M1 演示
│   └── Arylation_BMS_demo/         # 真数据 + 真 GP-BO
├── scripts/                        # 工具脚本
│   ├── build_project_template.py
│   ├── build_demo_project.py
│   ├── build_arylation_demo.py
│   ├── run_demo.py                 # Suzuki 3 轮 demo
│   ├── run_arylation_demo.py       # Arylation 5 轮 demo
│   └── verify_install.py           # M2 验收脚本
└── tests/                          # 44 个 pytest
    ├── test_excel_io.py            # 21 tests
    ├── test_domain_builder.py      # 14 tests
    ├── test_pipeline.py            # 4 tests
    └── test_watcher.py             # 5 tests
```

---

## 状态

| Milestone | 状态 | 验证 |
|-----------|------|------|
| **M1** 脚手架 | ✅ 完成 | 44/44 测试 + end-to-end demo |
| **M2** 安装脚本 | ✅ 完成 | verify_install.py 6/6 通过 |
| **M3** 文档 | ✅ 完成(本次) | README + 中文手册 + IT + FAQ |
| **M4** 真实项目试点 | 📋 待用户 | 找 1 位部门同事用 |
| **M5** 推广 | 📋 待用户 | ≥3 项目 |

### 测试

```bash
$ python -m pytest tests/ -q
44 passed, 5 warnings in ~5m (含真 GP-BO pipeline 4 个测试)
```

### 演示

```bash
$ python scripts/run_arylation_demo.py
# 5 轮真 GP-BO, 总耗时 ~2.5 分钟
# 找高产组合 + 缓存写主 xlsx
```

---

## 技术亮点

1. **单文件 Excel 工作流** — 4 sheet 按工作流顺序排列(🚀→⚙️→🧪→📊),状态机驱动下一步引导
2. **状态机 + 隐藏 helper** — `.status.json` + 隐藏 `_status` sheet + Toast 通知
3. **零 Excel 锁冲突** — Python 只写 `.cache/recommend_*.xlsx`,用户点 [🔄 刷新] 才合并到主 xlsx
4. **graceful degradation** — BoFire 不可用时降级到纯 Python 随机采样,代码零修改
5. **真 GP-BO 跑通** — BoFire 0.4.1 + botorch + qLogEI,在 1728 条真实反应上验证

---

## 文档

- [`docs/PRODUCT_DESIGN.md`](docs/PRODUCT_DESIGN.md) — 完整设计文档
- [`docs/同事使用手册.md`](docs/同事使用手册.md) — 给同事的中文手册
- [`docs/IT_装机手册.md`](docs/IT_装机手册.md) — 给 IT 的装机手册
- [`docs/BENCHMARK_ARYLATION.md`](docs/BENCHMARK_ARYLATION.md) — 真数据 benchmark
- [`docs/FAQ.md`](docs/FAQ.md) — 常见问题
- [`docs/DESIGN_DECISIONS.md`](docs/DESIGN_DECISIONS.md) — 设计决策日志

---

## License & 致谢

- 数据集来自 [Princeton Shields + BMS Reaction Optimization Game](https://github.com/connorcoley/reaction-optimization-game)
- 算法来自 [BoFire](https://github.com/experimental-design/bofire) + [BoTorch](https://botorch.org/)
- 内部项目,2026 年起始