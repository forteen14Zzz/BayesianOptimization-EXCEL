# BoFire Lab Portal — IT 装机手册

> 装机时间:约 5-10 分钟(主要花在 BoFire/torch 下载,~700 MB)
> 网络要求:能访问 PyPI + PyTorch CPU 源

---

## 系统要求

| 项 | 要求 |
|---|------|
| OS | Windows 10 / 11 |
| Python | **3.11 或 3.12**(BoFire 0.4.1 不支持 3.13+ / 3.14)|
| 磁盘 | ~3 GB(venv + BoFire + torch + botorch + cvxpy + pymoo + shap)|
| 内存 | 4 GB+ |
| 网络 | PyPI + pytorch.org 可访问(能下载 ~125 MB 的 torch CPU) |

⚠️ **不要装 Python 3.14+** — BoFire 0.4.1 不兼容;3.13 也能跑但有些兼容性 warning。
⚠️ **Python 必须 `Add to PATH`** — 否则 install.bat 找不到 python。

---

## 安装步骤

### 方法 1:标准装机(推荐)

1. 把整个 `E:\bofire\` 目录(或者打包的 zip)复制到目标机器,推荐 `D:\BOFIRELabPortal\`
2. 双击 `install\install.bat`
3. 等 5-10 分钟(进度:`[2/4] 装依赖...`,最耗时间)
4. 装完后会问"是否配开机自启动 watchdog?" — 推荐 `y`
5. 安装完成

### 方法 2:手动装机(当 install.bat 卡住用)

```cmd
cd /d D:\BOFIRELabPortal
python -m venv venv
venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install -r install\requirements.txt
```

如果 torch 下载慢:
```cmd
python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple torch
python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r install\requirements.txt
```

---

## 表格软件:WPS 还是 Excel?

**都支持**,程序会自动识别:
- **WPS 表格**(同事实际用的)→ COM ProgID `Ket.Application`
- **Microsoft Excel** → COM ProgID `Excel.Application`

自动保存/关闭/重开依赖 `pywin32`,它在 `requirements.txt` 里,正常装机就会装上。

⚠️ **WPS 需要开启 COM 支持**(默认是开的):
> WPS 表格 → 菜单「文件」→「选项」→「其他选项」→「高级」→ 勾选 **「启用 WPS 表格 COM 组件」**(有的版本在「系统」标签页)

没开的话:点按钮后表格不会自动关闭,程序会报 `Permission denied`,让同事**手工关掉表格再点按钮**即可(不影响其他功能)。

---

## ⚠️ 关于按钮的三个坑(2026-11 实测)

给 IT 的说明,避免重复排查:

| 坑 | 现象 | 本项目怎么绕的 |
|----|------|----------------|
| **相对路径失效** | 点了按钮没反应 | 所有按钮用**绝对路径**(`E:\bofire\projects\<项目>\*.bat`) |
| **bat 参数丢失** | 点 [生成记录表头] 却跑了 BO | 每个按钮一个**独立 bat 文件**,路径里不带参数 |
| **只写 `=HYPERLINK()` 公式** | 表格软件不认这是链接(`Hyperlinks.Count=0`) | 同时写真 `cell.hyperlink` 对象 |

3 个独立 bat:

| bat | 对应按钮 | 干什么 |
|-----|---------|--------|
| `跑_BO_点这里.bat` | `[🚀 启动 BO]` | ask(算推荐) + refresh(合并到看板) |
| `刷新看板.bat` | `[🔄 单独刷新]` | 只把 cache 合并到 📊 看板 |
| `生成记录表头.bat` | `[📋 生成记录表头]` | 只按 ⚙️ 配置 重建 🧪 记录 表头 |

按钮位置(每个项目 xlsx 里):
- **🚀 开始** sheet row 15-20 — 全局入口
- **⚙️ 配置** sheet 底部 row 37-42 — `🎬 操作` 区域(同事主要用这个:填完配置直接点)

---

## 验证安装

### 方法 1:看 demo Excel

打开 `D:\BOFIRELabPortal\projects\Suzuki_2026Q4_demo\Suzuki_2026Q4_demo.xlsx`

应该看到:
- 默认落在 `🚀 开始` sheet
- 切到 `🧪 记录` sheet 看到 5 条初始数据(1 条 valid=0)
- 切到 `⚙️ 配置` sheet 看到 4 个因子 + Yield 目标预填好

### 方法 2:跑 verify_install.py(M2 验收)

```cmd
cd /d D:\BOFIRELabPortal
python scripts\verify_install.py
```

应该看到 6/6 通过:
- ✅ 模板生成
- ✅ pytest 全通过(44 tests)
- ✅ 3 轮 end-to-end demo
- ✅ install.bat / setup_autostart.bat 语法
- ✅ start_watcher.bat 生成逻辑
- ✅ requirements.txt 锁版本

### 方法 3:手动验证 BoFire

```cmd
python -c "import bofire, botorch, torch; print('bofire', bofire.__version__); print('torch', torch.__version__)"
```

应输出:
```
bofire 0.4.1
torch 2.x.x+cpu
```

---

## 给用户的工作机部署

每个研究员:
1. 在 `projects\` 下新建自己的子目录(命名:`<项目名>_<姓名>`)
2. 复制 `templates\project_template.xlsx` 到子目录
3. 改名 `<项目名>_<姓名>.xlsx`
4. 告诉研究员打开,跟着 `🚀 开始` sheet 引导填

---

## 故障排查

| 现象 | 解决 |
|------|------|
| `python` 找不到 | 装 Python 时勾 "Add to PATH" |
| `pip install` 卡住(timeout) | 用清华源:`pip install -i https://pypi.tuna.tsinghua.edu.cn/simple <pkg>` |
| `venv` 创建失败 | 没用管理员运行 → 右键 → 以管理员身份运行 |
| `powershell` 创建快捷方式失败 | Windows 7 / Server 2012 不支持 `New-Object -ComObject WScript.Shell` |
| 装机完了但 Excel 打不开 | 看 `projects\<项目>\LOG.txt` 最后报错 |
| 用户看不到推荐 | 看 `_status!B2` 是不是 `error`,B11 有错误信息 |
| `No module named 'bofire'` | install.bat 没装完;手动跑 `pip install bofire==0.4.1` |
| `No module named 'cvxpy'` / `pymoo` / `shap` | requirements.txt 没列全;手动装 |
| BO 跑 70+ 秒 | 正常,BoFire 0.4.1 默认全数据 fit GP,30 数据点 ~30s |
| Python 3.14 警告 torch.jit.script | 不影响功能,可忽略;下次升级 BoFire ≥0.5 会修 |

---

## 日常运维

| 操作 | 怎么做 |
|------|--------|
| **备份** | 复制 `projects\` 整个目录(所有项目数据 + Excel + cache + 模型)|
| **升级** | 改 `install\requirements.txt` 里包版本号 → 重跑 `install\install.bat`(venv 已存在会跳过创建,只升级包)|
| **监控** | 看 `projects\<项目>\LOG.txt` 的最后 50 行 |
| **重置 demo** | 删 `projects\<项目>\.cache`,重跑 `scripts\run_demo.py` 或 `run_arylation_demo.py` |
| **查看真数据 demo** | `projects\Arylation_BMS_demo\` — 真实 1728 条 Pd-芳基化反应 |
| **跑测试** | `python -m pytest tests\ -q`(~5 分钟)|
| **添加新研究员** | 复制 `templates\project_template.xlsx` 到 `projects\<新项目>_<姓名>\`,改名,让用户填 |

---

## 文件清单(IT 视角)

```
D:\BOFIRELabPortal\
├── install\           ← 装机脚本
├── portal\            ← Python 代码(不要动)
├── templates\         ← 用户复制起点
├── projects\          ← 同事的项目数据
│   ├── Suzuki_2026Q4_demo\
│   └── Arylation_BMS_demo\   ← 真数据 benchmark
├── tests\             ← 单元测试(给 Bo / 我)
└── scripts\           ← 工具脚本
```

---

## 卸载/清理

1. 删除 `D:\BOFIRELabPortal\` 目录
2. 删除开机启动项里的 `BoFireLabPortal_watcher.lnk`(`shell:startup` 文件夹)
3. (可选)卸载 Python — 但其他程序可能用,谨慎

---

## 安全注意

- **无 VBA 宏** — 所有 `.xlsx` 文件纯公式 + 值,过 IT 宏扫描
- **数据全本地** — 无外部 API key,无云同步
- **依赖 pip 来源** — 默认 PyPI,可改清华源
- **不开新端口** — 后台 watchdog 只读本地文件,不监听网络

---

## 联系

- **同事遇到 Excel 操作问题** → 转给研究员 I(T 负责的同事对一线)
- **架构 / 代码 / 设计问题** → 转给 Bo
- **任何问题先看** `projects\<项目>\LOG.txt`

---

## 装机验收脚本

跑 `python scripts\verify_install.py` 输出:

```
=== M2 验收:模拟 IT 装机流程 ===
=== Step 1:依赖检查(轻量子集) ===
  ✓ openpyxl
  ✓ pydantic
  ✓ watchdog
  ⚠️ bofire (可选,未装时走降级)
=== Step 2:逐项验收 ===
  ✅ 模板生成
  ✅ pytest 全通过
  ✅ 3 轮 end-to-end demo
  ✅ install.bat / setup_autostart.bat 语法
  ✅ start_watcher.bat 生成逻辑
  ✅ requirements.txt 锁版本
=== M2 验收通过 ✅
```