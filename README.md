# BoFire Lab Portal

> **代号**: BoFire Lab Portal (BLP)
> **定位**: 给有机化学实验室的"无代码"贝叶斯优化工具 — 一个 `.xlsx` + 后台 Python。

## 项目结构

```
E:\bofire\
├── docs/                # 设计文档
├── install/             # IT 装机脚本(一次性)
├── templates/           # 用户复制起步用的 xlsx 模板
├── portal/              # Python 代码(后台服务 + BO 引擎)
│   ├── watcher.py       # watchdog 监听 Excel 文件
│   ├── run.py           # 命令行入口(.bat 应急用)
│   └── core/            # 核心模块
│       ├── excel_io.py
│       ├── domain_builder.py
│       ├── strategy_runner.py
│       ├── recommender.py
│       ├── pipeline.py
│       └── convergence.py
├── projects/            # 同事的所有项目(每个项目一个子目录)
│   └── Suzuki_2026Q4_demo/
└── tests/               # pytest 测试
```

## 快速开始(同事视角)

1. **一次性 IT 装机**:双击 `install\install.bat`,等 5-10 分钟
2. **新建项目**:复制 `templates\project_template.xlsx` → `projects\<项目名>\<项目名>.xlsx`
3. **跟着 🚀 开始 sheet 引导**:填 ⚙️ 配置 → 🧪 记录 → 📊 看推荐,后台自动跑 BO

## 快速开始(开发视角)

```bash
# 1. 创建 venv 并装依赖
python -m venv venv
.\venv\Scripts\activate
pip install -r install\requirements.txt

# 2. 跑测试
pytest tests\ -v

# 3. 跑 demo
python portal\run.py --project projects\Suzuki_2026Q4_demo --phase ask
```

## 设计文档

详见 `docs/PRODUCT_DESIGN.md`。

## 状态

🚧 **M1 脚手架进行中** — 见 git log。