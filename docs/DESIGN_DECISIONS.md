# BoFire Lab Portal — 设计决策日志

> 记录每个关键设计决策 + 当时的备选 + 为什么最终这样选。
> 给未来的我自己(以及后来人)看 — 为什么代码是现在这样。

---

## 1. 单文件 4 sheet, 而不是多文件

**决定**: 同事操作的项目 = 一个 `.xlsx` 文件,4 个 sheet(🚀→⚙️→🧪→📊)

**备选**:
- (a) 多文件:`config.xlsx` + `batch_*.xlsx` + `recommend_*.xlsx` + `data/history.csv`(初版)
- (b) Web UI(Gradio / Flask)
- (c) 命令行工具(用户填 JSON)

**为什么选单文件**:
- 用户明确说"只想用 Excel"
- 4 sheet 物理隔离用户编辑区 vs 系统写入区,不会相互覆盖
- 状态机 + Toast 提示驱动下一步,降低认知负担
- 离线 + 无网 + 不需装软件,符合禁 .exe / 禁云端要求

**关键 trade-off**:
- Excel 黄条"外部修改"问题 → 用 `.cache/` 中间文件 + 用户点 [🔄 刷新] 合并
- 用户编辑区 vs 系统区要**物理隔离**(`⚙️🎯🧪` 你编辑;`📊📜📋` Python 写)

---

## 2. watchdog + md5, 而不是 Excel COM / xlwings

**决定**: `watchdog.observers.Observer` 监听文件 mtime + 内容 md5 变化,触发 BO

**备选**:
- (a) Excel COM / `xlwings` / PyXLL — 需要装 .exe,公司禁
- (b) Excel VBA 宏 `Shell("py ...")` — 公司可能禁宏
- (c) 定时轮询(每 30s 扫一次)
- (d) 用户点按钮手动触发

**为什么选 watchdog + md5**:
- 公司网络禁 .exe / 禁宏,所以排除 COM/VBA
- watchdog 是事件驱动,CPU ≈ 0,比轮询省电
- md5 过滤避免 Ctrl+S 无改动误触发
- 用户也可以用 `.bat` 应急手动触发(方案 H,设计文档里)

**关键 trick**:
- 忽略 `~$xxx.xlsx` Excel 临时锁文件(标准做法)
- 1 秒 cooldown 防多事件合并误触
- 内容 md5 比较,而不是 mtime(因为 Excel 保存改 mtime)

---

## 3. 写 `.cache/`, 而不是直接写主 xlsx

**决定**: Python 只写 `.cache/recommend_<时间戳>.xlsx`,用户点 [🔄 刷新] 才合并到主 xlsx 的 📊 看板 sheet。

**备选**:
- (a) 直接写主 xlsx(经典方案)
- (b) 写新文件 `recommend_<时间戳>.xlsx`(初版)

**为什么选 cache + 手动刷新**:
- **避免 Excel 锁**:Python 用 openpyxl 写,Excel 开着会弹"外部修改"或拒绝
- **避免数据丢失**:用户编辑过程中,Python 覆盖可能丢用户数据
- **用户控制**:用户主动 [🔄 刷新] 时知道"系统在写"

**代价**:多一次点击。但这点点击远比"数据被覆盖"的灾难轻。

---

## 4. 状态机 + 隐藏 _status sheet, 而不是改 8 sheet 表头

**决定**: 用 `.status.json` + 隐藏 `_status` sheet + 各 sheet 顶部公式引用 `_status!B2` 等

**备选**:
- (a) 每个 sheet 表头都加状态文字(动态)
- (b) 单 sheet 内分块(像 form 表单)

**为什么选隐藏 helper**:
- 用户看 4 个 sheet 简单,不会出现 8 sheet 认知负担
- 隐藏 `_status` sheet 的 cell 是真值源,其他 sheet 公式引用
- Excel 重算公式自动刷新(不用重启)
- 不污染用户编辑区

**关键 trick**:
- `_status!B2` = `state`(awaiting_config / awaiting_bo / bo_running / bo_done / error)
- 各 sheet 顶部 **不要放公式**,而是**直接 =_status!B2**(看上去是静态文字)
- sheet tab 颜色标识当前 step(绿/蓝/橙/黄)

---

## 5. graceful degradation — BoFire 不可用时降级

**决定**: 代码里 try/except 让 BoFire 不可用 → 纯 Python 随机采样,代码不改一行

**备选**:
- (a) 强行要求 BoFire,装不上不让跑
- (b) BoFire 不可用直接抛错

**为什么选降级**:
- **125 MB torch 下载超时**的真实案例(M1 早期就遇到)
- 用户机器环境差异大,不强求一致
- 降级后功能还在(只是没 GP-BO 那种智能)
- 装上 BoFire 后**自动**用真 GP-BO(无需改代码)

**实现位置**:
- `domain_builder.build_domain()` — 装不上返回 None
- `strategy_runner._pure_python_ask()` — 终极 fallback
- 测试 `tests/test_pipeline.py` — 跑 fallback 路径也通过

---

## 6. BoFire 0.4.1, 不是最新版

**决定**: 锁 `bofire==0.4.1`

**备选**:
- (a) `bofire>=0.5`(新版)
- (b) latest(无锁)

**为什么选 0.4.1**:
- v5.2 老版本跑过,踩过的坑都熟(详见 memory)
- 0.5+ 重写 API,我们 workaround 失效
- 锁版本 → 可复现,跨机器一致
- 升级是大工作,等真有需求再说

**0.4.1 的已知坑**(集中在 `domain_builder.py`):
1. `DiscreteOutput` 会让 MOBO 失败 → 用 `ContinuousOutput`
2. `CategoricalOutput` + MOBO 被拒 → 降级 numeric 编码
3. `NonlinearConstraint` abstract → 用 `NonlinearInequalityConstraint`,无 `rhs` 参数
4. `Domain` 不要传 `constraints=None`
5. MOBO acf 必须是 EHVI 家族(qLogEHVI / qNEHVI)
6. 没有 `CategoricalTaskInput` → 用 `CategoricalInput`
7. 没有 `trainable_inputs` 参数 → `pin_non_trainable` 后处理
8. `ContinuousOutput` 没 `lower_bound` / `upper_bound` → 用 `MaximizeObjective` 等
9. `Domain.inputs` / `.outputs` 是 Inputs/Outputs 对象 → `.features` 是真正的 list

---

## 7. qLogEI, 不是 qEI / qUCB

**决定**: 默认 `qLogEI`(log Expected Improvement)

**备选**:
- (a) `qEI`(标准 Expected Improvement)
- (b) `qLogEI`(log EI — 数值更稳)
- (c) `qUCB`(Upper Confidence Bound)
- (d) `qNEI`(Noisy EI — 数据有噪声)

**为什么选 qLogEI**:
- BoFire 默认
- 数值稳定(避免 EI → 0 时梯度消失)
- 大部分反应筛选场景都适用
- 噪声大的场景用户可改 qNEI

---

## 8. 默认 30 数据点 / 3 批大小 / 5 初始点

**决定**: `n_initial=5, batch_size=3`

**备选**:
- (a) n_initial=20, batch_size=5(贪心)
- (b) n_initial=3, batch_size=1(保守)
- (c) 用户可调

**为什么选 5 / 3**:
- 5 初始点 + 3 推荐 = 8 数据,够 GP 起步
- 3 批大小 → 同事一轮做 3 个反应,工作量大但能跑
- 用户可在 Strategy 表里改

---

## 9. 把 cache 合并到主 xlsx 时保护 MergedCell

**决定**: `recommender.merge_cache_to_dashboard` 里加 `isinstance(target, MergedCell)` 检查

**备选**:
- (a) 直接 `ws.cell(...).value = ...`(会抛 AttributeError)
- (b) 把模板里的合并都解开

**为什么选 isinstance 检查**:
- 模板里 row 47/61 的章节标题 `_style_section` 用了 `merge_cells`
- 历轮推荐 row 49-60 写到 row 49 时如果它属于合并区就崩
- 解开合并会破坏模板结构(用户看到 📊 看板 不再整齐)

---

## 10. categorical 保留原值字符串,不是 int 编码

**决定**: `prepare_data` 把 categorical 列保留原值字符串给 BoFire

**备选**:
- (a) categorical 转 int 编码(经典 ML 做法)
- (b) 保留字符串(BoFire 要求)

**为什么选保留**:
- BoFire 0.4.1 `tell(experiments=...)` 内部会做 `validate`,要求 categorical 值在 domain 的 categories 列表里
- int 编码(0, 1, 2)会被 BoFire 拒绝:"invalid values for `溶剂`, allowed are [...]"
- 必须保留原值字符串

---

## 11. 真 BoFire 跑通后,降级路径仍保留

**决定**: 即便 BoFire 装上,代码仍走 `try/except` 降级路径

**为什么**:
- 不同环境(Windows / Mac / Linux)行为不同
- 同事的电脑可能 BoFire 装一半失败
- 降级路径让架构更鲁棒
- 测试同时跑两个路径都通过

**这是有意为之的冗余**,不是疏忽。

---

## 12. 没做 MOBO / 主动学习 / 仪器对接(Non-Goals)

参考 `docs/PRODUCT_DESIGN.md` §11 明确列出:
- ❌ 跨机器同步(云端、GitHub 推送)
- ❌ 多用户并发协作(项目加锁)
- ❌ 移动端 / Web UI
- ❌ 商业 BO 算法(JMP / Minitab)
- ❌ 主动学习 / 理论指导
- ❌ 实验仪器直连(无需对接 HPLC / 反应釜)

**为什么不做**:部门级小工具,M1-M5 范围内交付就够了。任何扩展都要先列到 M2+ 计划里。

---

## 13. 每个文件先 commit,不是攒一起 commit

参考 `agent.md`:
- 每次改完 git commit
- 不累积
- commit 前 working tree 干净

**为什么**: 写完代码发现 bug,git history 能精确定位哪次 commit 引入。攒 5 个文件一起 commit,出 bug 时 debug 信息密度低。

---

## 14. Arylation 5 轮 demo = 项目内部"烟雾测试"

**决定**: 真 BoFire 跑通后,用真数据(1728 条 Pd-反应)跑 5 轮 demo 验证

**为什么**:
- Mock 数据感觉不到 GP 的智能(随机就是 19% 平均 yield)
- 真数据下 GP 推荐出的高产组合和文献一致 → 验证算法 + 路径都 OK
- 给老板汇报有具体数字(77.4%, 65%)
- 找 bug(连续 categorical 编码失败就是从这里发现的)

**trade-off**:
- 跑一次 ~2.5 分钟
- 数据完全枚举(1728 个组合),BO 没有"超出"已知最优的可能
- 但 GP 5 轮找到的 77.4% 接近最优,验证工程 OK

---

## 总结

| 决策 | trade-off |
|------|-----------|
| 单文件 4 sheet | 认知清晰 vs 文件功能丰富度低 |
| watchdog + md5 | 事件驱动 vs 需要后台常驻 |
| `.cache/` + 手动刷新 | 零锁冲突 vs 多一步操作 |
| 隐藏 `_status` sheet | 4 sheet 视觉 vs 公式重算 |
| graceful degradation | 永远能跑 vs 偶尔有 silent degrade |
| BoFire 0.4.1 锁版 | 可复现 vs 用不上新版 |
| qLogEI 默认 | 数值稳 vs 偶尔欠探索 |
| 5 初始 + 3 批 | 起步快 vs 早期数据稀 |
| MergedCell 保护 | 模板稳 vs 偶尔丢数据 |
| categorical 原值 | BoFire 验证通过 vs 不能用 ML 编码 |
| 降级路径冗余 | 鲁棒 vs 代码冗余 |
| 内部 benchmark | 验证算法 vs 跑一次 2.5 分钟 |
| 每文件 commit | debug 精准 vs commit 多 |

写代码时回头看,大部分决策是为了**鲁棒性 + 用户体验**,而不是为了"技术酷炫"。