# BoFire Lab Portal

> **一句话**:给有机化学实验室的"无代码"贝叶斯优化工具 —— 同事只用一个 `.xlsx`,
> 系统自动算出下一组该做的实验条件。

```text
同事填 Excel  →  点 [🚀 启动 BO]  →  Python 跑 GP-BO  →  推荐写回 📊 看板
```

基于 [BoFire](https://github.com/experimental-design/bofire) + [BoTorch](https://botorch.org/)。

---

## 为什么需要它

做反应优化时,常见流程是「改条件 → 做实验 → 看结果 → 再改」。
问题是:**改哪些条件最有价值,靠经验猜**。贝叶斯优化(BO)能根据已有数据,
算出"下一组最值得做的条件",通常几十次实验就能找到高产组合。

但 BO 的门槛高(要写 Python、配 API)。这个工具把它藏起来:

| | 传统 BO | 这个工具 |
|---|---|---|
| 配置实验空间 | 写 Python | 填 ⚙️ 配置 sheet |
| 看推荐 | 看终端/调 API | 看 📊 看板 |
| 回填结果 | 写 DataFrame | 填 🧪 记录 sheet |
| 表格软件锁冲突 | — | 自动保存/关闭/重开表格 |

---

## 快速开始

### 1. 装环境

需要 **Python 3.11+**(推荐 3.11 / 3.12)。
检查:命令行敲 `python --version`。没有的话去 python.org 装,
安装时**务必勾选 "Add Python to PATH"**。

**方式 A — 不用命令行(推荐)**

1. 在本页面点 **Code → Download ZIP**
2. 解压到任意目录,例如 `D:\BOFIRELabPortal\`
3. 双击 **`install\install.bat`** — 等 5-10 分钟(装依赖)

**方式 B — 有 git 的话**

```bash
git clone https://github.com/forteen14Zzz/BayesianOptimization-EXCEL.git
cd BayesianOptimization-EXCEL
install\install.bat
```

装完会问是否配开机自启动(watcher,可选)。

### 2. 建项目

双击 **`init_new_project.bat`**,输入项目名,比如 `Suzuki_2026Q4_张三`,回车。

### 3. 用

打开 `projects\<项目名>\<项目名>.xlsx`,按 sheet 顺序走:

```
🚀 开始   →  ⚙️ 配置(填因子/目标)  →  🧪 记录(填已有数据)  →  📊 看板(看推荐)
```

**⚙️ 配置** sheet 底部有 `[📋 生成记录表头]` —— 填完配置点它,
系统按你的因子自动生成 🧪 记录 的表头(不用死记列顺序)。

之后每轮:点 `[🚀 启动 BO]` → 表格自动保存关闭 → 跑 BO → 表格自动打开,
新推荐已在 📊 看板。

详细步骤见 **[新手使用指南.md](新手使用指南.md)**。

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
.
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
│   └── project_template.xlsx       # 4 sheet + 隐藏 _status
├── portal/                         # Python 后台
│   ├── watcher.py                  # watchdog 监听 + 内容哈希过滤
│   ├── run.py                      # CLI 入口
│   └── core/
│       ├── excel_io.py             # Excel 读写 + pydantic 校验
│       ├── sync_headers.py         # 按配置重建记录表头 + 数据迁移
│       ├── domain_builder.py       # config → BoFire Domain
│       ├── strategy_runner.py      # ask (含降级路径)
│       ├── convergence.py          # 收敛判定(官方公式)
│       ├── convergence_plot.py     # 收敛趋势图
│       ├── excel_control.py        # WPS/Excel 自动保存/关闭/重开
│       ├── recommender.py          # .cache 写入 + 看板合并
│       └── pipeline.py             # 流程编排
├── projects/                       # demo 项目
│   ├── Suzuki_2026Q4_demo/         # 合成数据 demo
│   └── Arylation_BMS_demo/         # 真数据 + 真 GP-BO
├── scripts/                        # 工具脚本
│   ├── bat_templates.py            # 3 个独立 bat 的模板
│   ├── build_project_template.py
│   ├── build_demo_project.py
│   ├── build_arylation_demo.py
│   ├── init_new_project.py         # 新建项目
│   ├── upgrade_project_bats.py     # 老项目升级
│   ├── run_demo.py                 # Suzuki demo
│   ├── run_arylation_demo.py       # Arylation demo
│   └── verify_install.py           # 装机验收
└── tests/                          # 73 个 pytest
```

---

## 核心特性

### 单文件 Excel 工作流

4 个 sheet 按工作流顺序排列,每步都有"下一步去哪"的引导:

| Sheet | 谁写 | 内容 |
|-------|------|------|
| 🚀 开始 | 系统 | 状态、按钮、收敛提示 |
| ⚙️ 配置 | **用户** | 因子、目标、策略、约束 |
| 🧪 记录 | **用户** | 实验条件 + 实测 Yield |
| 📊 看板 | 系统 | 最新推荐、收敛曲线、日志 |

### 收敛判定用官方判据

用 BoFire 官方的 `ObjectiveImprovementCriterion`
([文档](https://experimental-design.github.io/bofire/docs/userguides/convergence_criteria.html)):

> 最近 `收敛窗口` 个实验里,最好成绩的提升 < `收敛阈值` → 判为收敛

⚠️ 窗口单位是**实验数**不是轮次。看板会自动画趋势图。

### 表格软件自动化

点按钮后自动 **保存 → 关闭 → 跑 BO → 重新打开**。
支持 **WPS 表格**(ProgID `Ket.Application`)和 **Microsoft Excel**。

### Graceful degradation

BoFire 不可用时降级到纯 Python 随机采样,代码零修改,流程照跑。

---

## 开发

```bash
python -m pip install -r install/requirements.txt
python -m pytest tests/ -q
```

**73 个测试**:单元 69 个 + pipeline 4 个(含真 GP-BO)。

⚠️ 依赖版本锁定在 `install/requirements.txt`(`bofire==0.5.0` — 官方
Convergence Criteria 从 0.5.0 才有)。

贡献规范见 [agent.md](agent.md):每次改动后必须补测试并 commit。

---

## 文档

| 文档 | 给谁看 |
|------|--------|
| [新手使用指南.md](新手使用指南.md) | 同事(中文,一步一步) |
| [docs/同事使用手册.md](docs/同事使用手册.md) | 同事(完整手册) |
| [docs/IT_装机手册.md](docs/IT_装机手册.md) | IT(装机 + 排障) |
| [docs/PRODUCT_DESIGN.md](docs/PRODUCT_DESIGN.md) | 设计全貌 |
| [docs/FAQ.md](docs/FAQ.md) | 常见问题 |
| [docs/BENCHMARK_ARYLATION.md](docs/BENCHMARK_ARYLATION.md) | 真数据 benchmark |
| [docs/DESIGN_DECISIONS.md](docs/DESIGN_DECISIONS.md) | 设计决策日志 |

---

## License & 致谢

本项目采用 [MIT License](LICENSE)。

- 算法来自 [BoFire](https://github.com/experimental-design/bofire) + [BoTorch](https://botorch.org/)
- 数据集来自 [Princeton Shields + BMS Reaction Optimization Game](https://github.com/connorcoley/reaction-optimization-game)
- 表格自动化:pywin32(WPS/Excel COM)