# BoFire Lab Portal — IT 装机手册

> 装机时间:约 5-10 分钟(主要花在 BoFire/torch 下载)
> 网络要求:能访问 PyPI + PyTorch CPU 源

---

## 系统要求

| 项 | 要求 |
|---|------|
| OS | Windows 10 / 11 |
| Python | **3.11 或 3.12**(BoFire 0.4.1 不支持 3.13+,需要 ≥ 3.9) |
| 磁盘 | ~3 GB(venv + BoFire + torch) |
| 网络 | PyPI + pytorch.org 可访问 |

⚠️ **不要装 Python 3.14+** — BoFire 0.4.1 不兼容。

## 安装步骤

1. 把 `E:\bofire\` 整个目录复制到目标机器,推荐 `D:\BOFIRELabPortal\`
2. 双击 `install\install.bat`
3. 等 5-10 分钟(进度:`[2/4] 装依赖...`)
4. 装完后会问"是否配开机自启动 watchdog?" — 推荐输 `y`
5. 安装完成

## 验证安装

打开 `D:\BOFIRELabPortal\projects\Suzuki_2026Q4_demo\Suzuki_2026Q4_demo.xlsx`

应该看到:
- 默认落在 `🚀 开始` sheet
- 切到 `🧪 记录` sheet 看到 5 条初始数据
- 切到 `⚙️ 配置` sheet 看到 4 个因子 + Yield 目标预填好

## 卸载/清理

- 删除 `D:\BOFIRELabPortal\` 目录
- 删除开机启动项里的 `BoFireLabPortal_watcher.lnk`
  (`shell:startup` 文件夹)

## 给用户的工作机部署

每个研究员:
1. 在 `projects\` 下新建自己的子目录(命名:`<项目名>_<姓名>`)
2. 复制 `templates\project_template.xlsx` 到子目录
3. 改名 `<项目名>_<姓名>.xlsx`
4. 告诉研究员打开,跟着 `🚀 开始` sheet 引导填

## 故障排查

| 现象 | 解决 |
|------|------|
| `python` 找不到 | 装 Python 时勾 "Add to PATH" |
| `pip install` 卡住 | 检查网络;手动 `pip install bofire==0.4.1 -i https://pypi.tuna.tsinghua.edu.cn/simple` |
| `venv` 创建失败 | 没用管理员运行 → 右键 → 以管理员身份运行 |
| `powershell` 创建快捷方式失败 | Windows 7 / Server 2012 不支持 `New-Object -ComObject WScript.Shell` |
| 装机完了但 Excel 打不开 | 看 `projects\<项目>\LOG.txt` 最后报错 |
| 用户看不到推荐 | 看 `_status!B2` 是不是 `error`,B11 有错误信息 |

## 日常运维

| 操作 | 怎么做 |
|------|--------|
| 备份 | 复制 `projects\` 整个目录(所有项目数据 + Excel + cache + 模型) |
| 升级 | 改 `install\requirements.txt` 里包版本号 → 重跑 `install\install.bat`(venv 已存在会跳过创建,只升级包) |
| 监控 | 看 `projects\<项目>\LOG.txt` 的最后 50 行 |
| 重置 demo | 删 `projects\Suzuki_2026Q4_demo\.cache`,重跑 `scripts\run_demo.py` |

## 文件清单(IT 视角)

```
D:\BOFIRELabPortal\
├── install\           ← 装机脚本
├── portal\            ← Python 代码(不要动)
├── templates\         ← 用户复制起点
├── projects\          ← 同事的项目数据
├── tests\             ← 单元测试(给 Bo / 我)
└── scripts\           ← 工具脚本(给 Bo 也跑)
```

## 联系

任何问题先看 `projects\<项目>\LOG.txt` → 转给 Bo 调试。