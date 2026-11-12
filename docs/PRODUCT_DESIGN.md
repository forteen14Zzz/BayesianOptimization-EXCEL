# 产品设计文档 — BoFire Lab Portal

> **代号**: BoFire Lab Portal(BLP)
> **定位**: 给有机化学实验室的同事用的"无代码"贝叶斯优化工具
> **目标用户**: 部门同事,完全无代码经验,但会用 Excel,公司网络禁止安装 .exe
> **文档状态**: v1.0 草案
> **作者**: Bo 🦞
> **最后更新**: 2026-09-29

---

## 1. 一句话定位

**"用 Excel 配因子、跑完贴回结果,系统自动告诉你下一组实验该做什么条件。"

---

## 2. 用户与场景

### 2.1 目标用户画像

| 维度 | 描述 |
|------|------|
| 角色 | 有机化学实验室研究员 / 工程师 |
| 技术水平 | 熟练 Excel;**不会** Python / 命令行 / 装软件 |
| 设备 | 公司电脑,网络受控,**禁止安装 .exe**(只允许 IT 审批过的白名单软件) |
| 工作模式 | 一个研究员可能同时跑多个反应项目(模板 / 底物 / 催化剂筛选) |

### 2.2 典型使用流程(从用户视角)

```
T0  IT 给我们装好工具(一次性,后面我们负责)
T1  我新建项目 "Suzuki_2026Q4",填入 6 个因子 + 1 个目标(Yield)
T2  系统给我返回 3 组初始条件 → 我去实验室做实验
T3  把 Yield 测出来 → 填回 Excel → 点一个按钮
T4  系统用 GP 拟合 + Bayesian Optimization,再给我 3 组条件
T5  重复 T2-T4 直到收敛(我看着曲线觉得差不多了就停)
```

### 2.3 用户**不**需要关心的事

- Python / pip / venv / 命令行(但 IT 安装的人需要)
- GP / EI / 贝叶斯优化的数学
- 数据备份(系统自动)
- 多用户/多项目隔离的细节(自动)

---

## 3. 硬约束(为什么这么设计)

| 约束 | 来源 | 设计影响 |
|------|------|----------|
| 公司网络禁装 .exe | 用户答 | 不能是 Electron 桌面应用 / 商业 BO 软件(如 JMP / Minitab)/ LabVIEW |
| 用户只懂 Excel | 用户答 | UI 主入口必须是 .xlsx;Python 仅在 IT 装机时用 |
| 多项目隔离 | 用户答 | 每个项目一个独立文件夹,绝不共享 |
| 实验成本/耗时不定 | 用户答 | 不能假设"便宜能跑 1000 次";也要支持"只能跑 10 次"的昂贵场景 |
| 有机化学实验室 | 用户答 | 因子类型以连续(温度/当量)+ 分类(催化剂/溶剂)为主;可能涉及 SMILES(分子式) |

---

## 4. 产品形态

### 4.1 整体架构(分 3 层)

```
┌──────────────────────────────────────────┐
│  Layer 1 · Excel 前端(用户唯一接触层)    │
│  ─ config.xlsx   因子/目标/策略定义      │
│  ─ batch_*.xlsx  实验结果录入            │
│  ─ recommend_*.xlsx 系统返回的推荐条件   │
├──────────────────────────────────────────┤
│  Layer 2 · Excel ↔ Python 桥接           │
│  ─ Excel 解析(openpyxl + 校验)           │
│  ─ Domain 构建(BoFire)                   │
│  ─ GP 训练 + BO ask()(BoFire + botorch)  │
│  ─ 结果回写 Excel                        │
├──────────────────────────────────────────┤
│  Layer 3 · 计算引擎                       │
│  ─ BoFire 0.4.1(Bayesian Optimization)  │
│  ─ botorch 0.x(底层)                     │
│  ─ RDKit(分子输入,可选)                  │
└──────────────────────────────────────────┘
```

### 4.2 关键设计决策

| 决策 | 选择 | 否决方案 | 理由 |
|------|------|----------|------|
| **触发方式** | Excel 内嵌按钮调用 Python 脚本 | Excel VBA 宏 / Power Automate | 公司可能禁宏;Power Automate 不在白名单 |
| **运行方式** | Python 后台常驻监听 Excel 文件变化 | 双击 .bat 触发单步 | 同事**怕**命令行;后台跑对用户透明 |
| **Excel 模板** | 每项目 1 套 config + N 个 batch + 系统生成的 recommend | 通用模板 | 项目隔离 + 用户易理解 |
| **状态保存** | 每个项目独立目录 + history.csv(累积,不覆盖) | 数据库 | 微信扫不到 .db,Excel 能 grep |
| **可视化** | Python 生成 PNG 嵌入 recommend.xlsx | Excel 图表公式 | GP 拟合曲线 Excel 画不出来 |
| **Python 依赖** | 锁版本(pip freeze 到 requirements.txt) | latest | 跨机器/重装稳定性 |

### 4.3 项目目录结构(标准)

```
BoFireLabPortal/
├── README.md                          ← 给同事的快速入门
├── docs/
│   └── PRODUCT_DESIGN.md              ← 本文档
├── install/                           ← IT 装机用,用户不看
│   ├── install.bat
│   ├── requirements.txt              ← 锁定:bofire==0.4.1,botorch,openpyxl,pandas,rdkit
│   └── README_For_IT.md
├── templates/                        ← 用户复制用
│   ├── config_template.xlsx          ← 因子/目标/策略 schema
│   └── batch_template.xlsx
├── projects/                         ← 同事的所有项目
│   ├── Suzuki_2026Q4_张三/
│   │   ├── config.xlsx               ← 用户编辑
│   │   ├── input/
│   │   │   └── batch_2026-09-29.xlsx ← 用户填结果
│   │   ├── output/
│   │   │   └── recommend_2026-09-29_143022.xlsx          ← 系统生成
│   │   ├── data/
│   │   │   └── history.csv           ← 自动维护
│   │   ├── models/                   ← GP 模型快照(高级功能)
│   │   └── LOG.txt                   ← 每次 BO 的运行日志
│   └── Heck_2026Q4_李四/
│       └── ...(同上)
├── portal/                           ← Python 桥接层
│   ├── watcher.py                    ← 文件监听入口
│   ├── core/
│   │   ├── excel_io.py               ← xlsx 读写 + schema 校验
│   │   ├── domain_builder.py         ← config.xlsx → BoFire Domain
│   │   ├── strategy_runner.py        ← ask() / tell()
│   │   ├── recommender.py            ← 推荐结果打包
│   │   └── diagnostics.py            ← 收敛分析、特征重要性
│   ├── tests/                        ← 单元测试(给开发/我)
│   └── config.json                   ← 全局配置
└── AGENTS.md / agent.md              ← Agent 协作规范
```

---

## 5. Excel Schema(用户最关心的设计)

### 5.1 `config.xlsx` — 项目定义

包含 4 张 sheet:

#### Sheet 1: `Factors`(实验因子)

| key(列名) | type | trainable | low | high | categories | notes |
|------------|------|-----------|-----|------|------------|-------|
| 温度 | continuous | TRUE | 0 | 100 |  | °C |
| 催化剂用量 | continuous | TRUE | 0.1 | 5.0 |  | mol% |
| 反应时间 | continuous | TRUE | 1 | 24 |  | h |
| 溶剂 | categorical | TRUE |  |  | 甲苯;THF;DMF | |
| 底物 | categorical_descriptor | FALSE |  |  | 苯硼酸;噻吩硼酸 | 留 trainable=FALSE 表示固定 |

支持的 `type`(覆盖 90% 有机化学实验):

| 类型 | 含义 | 必须填的列 |
|------|------|----------|
| `continuous` | 连续变量(温度、当量、时间) | `low`、`high` |
| `discrete` | 离散数值(如 1/5/10 eq) | `categories` (用 `;` 分隔) |
| `categorical` | 分类变量(溶剂、催化剂) | `categories` |
| `categorical_descriptor` | 分类 + 数值描述符(如不同配体带 cost、logP) | `categories` + `descriptors`(JSON) |
| `categorical_molecular` | SMILES 分子 | `categories`(SMILES `;` 分隔),需要 RDKit |

#### Sheet 2: `Outputs`(实验目标)

| key | direction | target | lower_bound | upper_bound | unit |
|-----|-----------|--------|--------------|-------|------|
| Yield | maximize |  | 0 | 100 | % |
| 转化率 | maximize |  | 0 | 100 | % |
| 选择性 | target | 95 | 0 | 100 | % |

支持的 `direction`:

- `maximize`(越大越好)— Yield / conversion / activity
- `minimize`(越小越好)— impurity / cost / byproduct
- `target`(越接近越好)— selectivity / pH / enantiomeric excess

#### Sheet 3: `Strategy`(BO 策略)

| key | value | options | 说明(给用户看的) |
|-----|-------|---------|------------------|
| 策略类型 | sobo | random;sobo;mobo;doe | 单目标 GP-BO(默认) |
| 采集函数 | qLogEI | qLogEI;qLogNEI;qEI;qUCB | 标准 EI(默认);噪声大选 qLogNEI |
| 最少初始点数 | 5 |  | 数据 < 此值用随机采样 |
| 批大小 | 3 |  | 每轮推荐的实验数 |

**默认值策略**: `sobo` + `qLogEI` + 5 初始点 + 3 批大小 — 适用于 90% 反应筛选场景。用户改坏了不会崩,有 schema 校验。

#### Sheet 4: `Constraints`(可选,高级)

用户一般**不需要填**。复杂场景(温度+时间总和不能>100)用:

| name | type | features | coefficients | rhs |
|------|------|----------|--------------|-----|
| 总能耗 | linear_inequality | 温度;反应时间 | 1;1 | 100 |

### 5.2 `batch_<日期>.xlsx` — 实验记录

| trial_id | batch_id | 温度 | 催化剂用量 | 反应时间 | 溶剂 | Yield | valid_Yield | status |
|----------|----------|------|----------|----------|------|-------|-------------|--------|
| T0001 | 2026-09-29 | 60 | 2.0 | 12 | 甲苯 | 78.5 | 1 | measured |
| T0002 | 2026-09-29 | 80 | 1.5 | 8 | THF | 82.1 | 1 | measured |

- `trial_id` 自动生成 T0001 起,用户可不管
- `valid_*` = 0 表示实验失败,该行不进入训练
- `status` 系统填 `measured`;用户可标 `failed` / `pending`

### 5.3 `recommend_<时间戳>.xlsx` — 系统返回(用户读)

3 张 sheet:

1. **推荐条件** — 下一步要做的 N 组条件 + GP 预测 mean/sd
2. **收敛曲线** — 历轮最佳 Yield,带图(嵌入 PNG)
3. **本轮摘要** — 跑了多少数据点、用了什么策略、用了多少秒

---

## 6. 用户操作手册(给同事的)

### 6.1 一次性 IT 安装

```
1. IT 在 D:\BOFIRELabPortal\ 解压本仓库
2. 双击 install\install.bat
3. 等 5-10 分钟(装 BoFire 等 Python 包,可能要预编译 torch)
4. 安装完成,以后研究员就不需要再碰 Python
```

### 6.2 研究员日常(关键:不碰 Python)

```
1. 在 projects\ 目录下复制 templates\config_template.xlsx
2. 改名 Suzuki_2026Q4_张三\config.xlsx
3. 双击打开 → 在 Factors 表填因子 → Outputs 表填目标 → Strategy 表不改 → 保存关闭
4. 第一次点 config.xlsx 里的 [🚀 启动 BO] 按钮
   → 系统生成第一组推荐条件(input 目录有 output 链接)
5. 去实验室做这 3 组实验,测 Yield
6. 复制 templates\batch_template.xlsx → 改名为 batch_2026-09-29.xlsx
   → 把推荐条件复制过来 → 在 Yield 列填实测值 → 保存关闭
7. 点 batch_*.xlsx 里的 [📊 提交数据] 按钮
   → 系统自动跑 BO → 生成新的 recommend_*.xlsx
8. 回到第 5 步,直到满意
```

> Excel 内嵌按钮如何调用 Python?见 §9.4 待解决的技术问题。

---

## 7. 非功能性需求

| 维度 | 目标 |
|------|------|
| **正确性** | 所有 BO 推荐数学结果可复现(锁定 bofire==0.4.1,seed 固定) |
| **可用性** | 研究员从拿到模板到第一次"启动 BO"≤ 10 分钟 |
| **可靠性** | 用户误填 config.xlsx → 报错信息友好,不崩 |
| **性能** | 单轮 BO(10-50 数据点)≤ 60 秒 |
| **可观测** | 每次 BO 写 LOG.txt:输入/输出/异常/耗时 |
| **可移植** | 整个 projects/ 目录可整体拷给同事,在另一台机器恢复 |

---

## 8. 安全与合规

| 项 | 处理 |
|---|------|
| 公司网络禁 .exe | **不依赖 .exe**;Python 解释器走 IT 白名单或 portable Python |
| Excel 文件 | 公式安全(无 VBA 宏,避免被宏拦截);用 `=HYPERLINK()` + 文件监听实现"按钮" |
| 数据隐私 | projects/ 默认放本地 D 盘;不联网,数据不出机器 |
| 数据备份 | 建议 IT 配置 projects/ 目录的每日快照;不在产品里内置 |
| 凭据 | 无外部 API key;完全离线运行 |

---

## 9. 待解决问题 / Open Questions

| # | 问题 | 候选方案 |
|---|------|---------|
| 9.1 | **Python 环境怎么不靠 .exe 装?** | (a) IT 预先装好 Python in user-space;(b) Portable Python zip;(c) Excel VBA 调 wscript |
| 9.2 | **Excel "按钮"如何触发 Python?** | (a) Excel `=HYPERLINK()` 打开 .py 脚本;(b) 后台 `watchdog` 监听文件 mtime;(d) Excel COM 接口(需 .exe) |
| 9.3 | **同事的 config.xlsx 改坏了** | schema 校验 → 友好的红色错误提示 + 不保存推荐 |
| 9.4 | **多用户隔离到机器级别** | 现阶段:每台机器一份 projects/;多人共享机器靠文件夹权限 |
| 9.5 | **模型快照(models/)** | 默认不存;高级用户可勾选"存模型"以复现 |
| 9.6 | **多目标 MOBO 的 UI** | mobo 策略 UI 怎么做帕累托前沿可视化?推荐 sheet 加 1 张散点图 |
| 9.7 | **收敛判断** | 默认 10 轮无显著提升 → 提示"可以停了";阈值用户可调 |
| 9.8 | **RDKit 装不上时降级** | molecular 输入检测到没装 RDKit → 降级为 categorical + 警告 |

---

## 10. 验收标准(Milestone)

| 里程碑 | 内容 | 验收 |
|--------|------|------|
| **M1**: 脚手架 | `E:\bofire\` 含 config_template / batch_template / 1 个 demo project / 能跑通 end-to-end | demo_reaction 跑 3 轮,recommend 合理 |
| **M2**: 安装脚本 | `install.bat` 双击能在新机器完成 venv 装包 | IT 同事盲测,10 分钟内完成 |
| **M3**: 文档 | 给同事的 README(中文)+ 给 IT 的 README | 同事不看文档能跟着用 |
| **M4**: 真实项目试点 | 1 位部门同事用本工具跑了 1 个真实反应 | 同事反馈"易用、不出错" |
| **M5**: 推广 | 工具在部门内有 ≥ 3 个项目在用 | 部门领导认可 |

---

## 11. 不做(Non-Goals)

为了让项目能在合理时间内交付,**明确不做**:

- ❌ 跨机器同步(云端、GitHub 推送)— IT 网络管控不允许
- ❌ 多用户并发协作(项目加锁)— 现阶段每项目 1 个 owner
- ❌ 移动端 / Web UI — Excel 是唯一入口
- ❌ 商业 BO 算法(如 JMP / Minitab)— 完全 Python
- ❌ 主动学习(Active Learning) / 理论指导(Theory-Guided)— 只做标准 BO
- ❌ 实验仪器直连(无需对接 HPLC / 反应釜)— 同事自己把数字抄到 Excel

---

## 12. 参考 / 灵感

| 来源 | 学到的 |
|------|--------|
| experimental-design/bofire(GitHub)| 核心算法库;已锁 0.4.1 |
| 之前的 `bo_portal_package_v5.2` | Excel schema 和工作流设计(本设计在此基础上重构) |
| organic-chemistry BO 文献(Shields 2021 等)| 真实 benchmark 数据 |

---

## 13. 下一步

确认本设计后:

1. 在 `E:\bofire\` 下创建项目骨架(milestone 1)
2. 先做 1 个 end-to-end demo(不追求完美,先跑通)
3. 真实同事试用,迭代