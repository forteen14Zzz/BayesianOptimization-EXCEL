# BoFire Lab Portal — 常见问题(FAQ)

> 用户、研究员、给 IT 都可能问到。

---

## 给用户(化学家)的 FAQ

### Q:我不会写代码,真的能用吗?

**A:能**。你的全部操作就是:
- 填 Excel 表(因子 / 实验结果)
- 点按钮(`[🚀 启动 BO]`)
- 看推荐条件

Python、GP-BO、贝叶斯优化全在后台,你完全不用懂。

### Q:Python 装包错误会不会影响我?

**A:不会**。如果 BoFire 没装好,系统会**自动降级**到"纯随机推荐" — 仍然能给你 3 组新条件,只是不会像 GP-BO 那样越跑越准。

⚠️ 这种情况请联系 IT 修 Python 环境,长期降级会浪费实验次数。

### Q:我的实验 fail 了怎么办?

**A:在 🧪 记录 sheet 里,把 `valid` 列填 `0` 即可**。BO 训练时会自动过滤掉这些行,不影响。

你也可以在 `status` 列填 `failed`,只是标记。

### Q:怎么知道 BO 还要不要继续跑?

**A:看「收敛提示」**。🚀 开始 sheet 底部和 📊 看板 里会显示:

| 显示 | 意思 |
|------|------|
| `⏳ 收敛评估中 — 已积累 N 个实验` | 数据还不够判,继续跑 |
| `⏳ 未收敛 — 最近 5 个实验内最好成绩提升了 0.32` | 还在进步,值得继续 |
| `✅ 已收敛 — 最近 5 个实验内最好成绩只提升了 0.002` | 涨不动了,可以停 |

判据用的是 **BoFire 官方 `ObjectiveImprovementCriterion`**:
最近 `收敛窗口` 个实验里,最好成绩的提升 < `收敛阈值` → 判为收敛。

⚠️ `收敛窗口` 的单位是**实验数**不是轮次 — batch_size=3 时一轮 = 3 个实验。

官方文档:https://experimental-design.github.io/bofire/docs/userguides/convergence_criteria.html

### Q:看板上的收敛图怎么看?

**A:📊 看板 的「📈 收敛曲线」区**,左边是数据表,右边是趋势图。

| 图上看到 | 意思 |
|----------|------|
| 绿线(最好成绩)还在往上爬 | 有进展,继续跑 |
| 绿线变平了(平台期) | 涨不动了,考虑收工 |
| 橙色标题「已收敛」 | 系统判定收敛,建议停止 |
| 黄色带子(±阈值) | 提升量在阈值带里,属于"基本没动" |

每次点 `[🚀 启动 BO]` 会自动重画。图存在项目的 `.cache\convergence.png`。

### Q:推荐的条件我不知道能不能做(没有原料、没有设备)怎么办?

**A**:
- 直接忽略那条推荐,在 🧪 记录 sheet 里把它的 `valid` 填 `0`(或者干脆删掉那行)
- BO 下轮会基于"实际可做"的实验重新推荐

### Q:BO 推荐的条件看起来很离谱(比如温度 0°C)怎么办?

**A**:
- 检查 ⚙️ 配置 sheet 的因子 `low` / `high` 设置 — 是不是设错了
- 默认 80-130°C,如果你改成了 0-200°C,GP 会在边界外探索
- 改回正确范围后,点 [🚀 启动 BO] 重跑

### Q:为什么推荐每次都不一样?

**A**:GP 内部有随机性(即使 seed 固定,BoFire 默认有少量随机)。如果你想完全一致:
- 改 `install\requirements.txt` 里 torch 的 seed
- 或在 Strategy 表里改 `策略类型` 为 `doe`(确定性的 design of experiments)

### Q:跑了 10 轮还是没收敛怎么办?

**A**:
- 可能你的反应空间太大(因子多 / 范围大) — 缩小范围
- 可能实验重复太多噪声大 — 加更多 valid=0 标记失败点
- 可能反应没有明显梯度(随机就够) — GP 没优势,考虑 DOE

---

## 给 IT / DevOps 的 FAQ

### Q:Python 3.14 装不上 BoFire 怎么办?

**A**:降级到 Python 3.11 或 3.12。BoFire 0.4.1 不支持 3.13+(Python 3.13 偶尔能跑但有 warning,3.14 必崩)。

```cmd
# 卸 3.14 装 3.12
python-3.12.exe /quiet InstallAllUsers=1 PrependPath=1
```

### Q:torch 下载超时(125MB 卡 5 分钟)?

**A**:用清华源 / 阿里源:
```cmd
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple torch
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r install\requirements.txt
```

### Q:用户说"BO 没反应",怎么排查?

**A**:看 `projects\<项目>\LOG.txt`:
```
[2026-09-29 14:30:00] phase=ask project=...
[2026-09-29 14:30:30] ERROR: No module named 'cvxpy'
```

最常见错:
- `No module named 'cvxpy'` / `pymoo` / `shap` → 装包不全
- `RuntimeError: CUDA not available` → torch 装了 GPU 版,改成 CPU
- `ValidationError` → 用户 ⚙️ 配置 填错了
- `MergedCell` → 模板被用户改坏了,从 `templates\` 重拷一份

### Q:多用户怎么隔离?

**A**:
- **现阶段**:每个项目一个子目录,靠 Windows 文件夹 ACL 隔离
- **不**支持并发编辑(Non-Goals 已声明)
- 多用户共用一台机器 OK(互不干扰)

### Q:怎么备份?

**A**:复制 `projects\` 整个目录。这是**所有数据 + Excel + cache + 模型**的来源,丢了全没了。

### Q:升级 BoFire 版本(比如 0.5)会破坏兼容吗?

**A**:**会**。BoFire 0.4.1 有大量 workaround 在 `domain_builder.py`,升级到 0.5+ 需要重写:
- `acquisition_function` 接受 qLogEI 字符串(0.5+)
- `SoboStrategy` 默认有 `trainable_inputs` 参数(0.5+)
- `Domain.outputs` 有 `.features`(0.5+)

不升级到 0.5 也能跑 — 项目锁了 `bofire==0.4.1`。什么时候想升级再说。

---

## 给 Bo / 我(开发者)的 FAQ

### Q:为什么不用更新的 BoFire (0.5+)?

**A**:本项目锁定 `bofire==0.4.1` 因为:
- 0.5+ 重写了一大堆 API,我们的 workaround 失效
- 0.4.1 是我们之前 v5.2 跑过的稳定版
- 升级是大工作,等真有需求再做

### Q:为什么不用多目标 MOBO?

**A**:M1 范围里是 "do standard BO",MOBO 是 v2 才做。设计文档 §11 明确列 Non-Goals。

但 BoFire 0.4.1 路径已经留好(`resolve_acf_function` 已支持 mobo + EHVI 家族)。要做直接加 Outputs 列。

### Q:为什么用 BoFire 不用 Ax / Optuna / Hyperopt?

**A**:
- **BoFire + botorch**:专为化学反应设计的 categorical + continuous,文档好,Lock 0.4.1
- **Optuna**:适合 ML hyperparameter,但 categorical + continuous 边界控制弱
- **Ax**:Meta 出品,但 API 复杂,需要 GPU 训练时启动慢
- **Hyperopt**:太老,GP 实现弱

BoFire + botorch 是化学反应场景的最佳选择。

### Q:为什么不用 Docker / 云端?

**A**:Non-Goals 列了:❌ 跨机器同步 + ❌ 多用户并发 + ❌ Web UI。公司网络禁 .exe,纯 .bat + Python 是最佳兼容方案。

### Q:跑 5 轮 30 数据点 / 秒够快吗?

**A**:**够**。BoFire 0.4.1 默认 fit GP 全数据 + optimization 300 重启,在 30-50 点场景下 30s 正常。

未来如果数据点 > 100,可改 `num_restarts=50`,等高负载再调。

### Q:为什么不用 GP 预测的 mean/sd?

**A**:M1 范围里 yield_mean / yield_sd 是 None — 因为我没写 predict 步骤。这是 v2 工作(给老板看 "GP 觉得这一组大概是 X%")。

实际逻辑:
```python
from bofire.surrogates.botorch import BotorchSurrogates
preds = strat.predict(candidates_df)
```

加一行就完了,30 行代码量,但需要 v2 评估需求。

---

## 总结

**有其他问题?**

- **同事/化学家**:看 `docs\同事使用手册.md`
- **IT/运维**:看 `docs\IT_装机手册.md`
- **设计/开发**:看 `docs\PRODUCT_DESIGN.md` + `docs\DESIGN_DECISIONS.md`
- **真数据表现**:看 `docs\BENCHMARK_ARYLATION.md`