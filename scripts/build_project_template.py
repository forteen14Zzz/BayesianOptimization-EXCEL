"""生成 templates/project_template.xlsx — 4 sheet + 隐藏 _status + 公式预设 + 按钮。

这是用户复制的模板起点,所有 sheet 结构、公式、按钮、单元格样式在这里定义。

设计: 4 个用户可见 sheet + 1 个隐藏 _status helper sheet
  1. 🚀 开始     — 状态面板 + 按钮(只读 + 公式)
  2. ⚙️ 配置   — 因子 + 目标 + 策略(用户编辑,Python 不写)
  3. 🧪 记录   — 实验记录(用户编辑,Python 只追加 status 列)
  4. 📊 看板   — 推荐/曲线/日志(系统写,用户只读)

隐藏:
  _status        — Python 写,各 sheet 公式引用
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName


# ---------- 样式常量 ----------
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="FFFFFF")
TITLE_FILL = PatternFill("solid", fgColor="305496")  # 深蓝
SECTION_FONT = Font(name="微软雅黑", size=12, bold=True, color="1F4E78")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")  # 浅蓝
TABLE_HEADER_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
TABLE_HEADER_FILL = PatternFill("solid", fgColor="4472C4")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")  # 浅黄 — 用户填的区域
HINT_FONT = Font(name="微软雅黑", size=9, italic=True, color="808080")
BUTTON_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
BUTTON_FILL = PatternFill("solid", fgColor="70AD47")  # 绿
THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)


def _style_title(ws, row, col, text, span=8):
    """sheet 顶部大标题栏"""
    ws.cell(row=row, column=col, value=text)
    c = ws.cell(row=row, column=col)
    c.font = TITLE_FONT
    c.fill = TITLE_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=row, end_row=row, start_column=col, end_column=col + span - 1)
    ws.row_dimensions[row].height = 28


def _style_section(ws, row, col, text, span=8):
    """章节标题"""
    ws.cell(row=row, column=col, value=text)
    c = ws.cell(row=row, column=col)
    c.font = SECTION_FONT
    c.fill = SECTION_FILL
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(start_row=row, end_row=row, start_column=col, end_column=col + span - 1)
    ws.row_dimensions[row].height = 22


def _style_table_header(ws, row, headers, start_col=1):
    """数据表表头"""
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + i, value=h)
        c.font = TABLE_HEADER_FONT
        c.fill = TABLE_HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER
    ws.row_dimensions[row].height = 32


def _style_input(ws, row, col, hint=None):
    """用户输入单元格(浅黄底)"""
    c = ws.cell(row=row, column=col)
    c.fill = INPUT_FILL
    c.border = THIN_BORDER
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    if hint:
        ws.cell(row=row, column=col + 1, value=hint).font = HINT_FONT


def _button(ws, row, col, label, target_sheet):
    """按钮: 跳到指定 sheet 的 A1(内部链接,不需要绝对路径)。"""
    from openpyxl.worksheet.hyperlink import Hyperlink

    ref = f"#'{target_sheet}'!A1"
    c = ws.cell(row=row, column=col, value=f'=HYPERLINK("{ref}","{label}")')
    # ⚠️ 同时写真 hyperlink 对象,让 Excel 认出这是链接(否则点了没反应)
    c.hyperlink = Hyperlink(
        ref=c.coordinate,
        location=f"'{target_sheet}'!A1",
        tooltip=label,
    )
    c.font = BUTTON_FONT
    c.fill = BUTTON_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = THIN_BORDER


def _button_cmd(ws, row, col, label, command):
    """按钮: 点击执行 .bat。

    ⚠️ 两个关键点(2026-10-08 实测):

    1. **绝对路径** — Excel 解析相对路径用 Excel 当前工作目录,
       不是 xlsx 所在目录 → 相对路径必失效。用占位符 {{BO_BAT}},
       init_new_project.py 生成项目时替换成绝对路径。

    2. **真 hyperlink 对象** — openpyxl 只写 `=HYPERLINK()` 公式时,
       Excel 打开后 `Range.Hyperlinks.Count` = 0,不认这是超链接。
       必须同时设 `cell.hyperlink = Hyperlink(...)`,Excel 才会
       把它当可点击链接处理。
    """
    from openpyxl.worksheet.hyperlink import Hyperlink

    safe_label = label.replace('"', '""')
    safe_href = command.replace('"', '""')

    c = ws.cell(row=row, column=col, value=f'=HYPERLINK("{safe_href}","{safe_label}")')
    # ⚠️ 同时写真 hyperlink 对象,让 Excel 认出这是链接
    c.hyperlink = Hyperlink(
        ref=c.coordinate,
        target=safe_href,
        tooltip=label,
    )
    c.font = BUTTON_FONT
    c.fill = BUTTON_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = THIN_BORDER


def build_start_sheet(wb: Workbook) -> None:
    """Sheet 1: 🚀 开始 — 状态面板 + 按钮 + 下一步引导"""
    ws = wb["🚀 开始"]
    # 列宽
    for col, w in enumerate([18, 18, 18, 18, 18, 18, 18, 18], start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    _style_title(ws, 1, 1, "🚀 BoFire Lab Portal — 项目仪表板", span=8)

    # 项目信息区
    _style_section(ws, 3, 1, "📋 项目信息", span=8)
    ws["A4"] = "项目名:"
    ws["B4"] = "=MID(CELL('FILENAME_LEN'),1,0)&\"\""  # 占位,实际用 _status 引用
    # 项目名用 _status!B5
    ws["B4"] = "=_status!B5"
    ws["A4"].font = Font(bold=True)

    ws["A5"] = "项目目录:"
    ws["B5"] = "=_status!B6"
    ws["A5"].font = Font(bold=True)

    ws["A6"] = "当前轮次:"
    ws["B6"] = "=_status!B7"
    ws["A6"].font = Font(bold=True)

    ws["A7"] = "数据点数:"
    ws["B7"] = "=_status!B8"
    ws["A7"].font = Font(bold=True)

    # 状态区
    _style_section(ws, 9, 1, "📡 当前状态", span=8)
    ws["A10"] = "状态:"
    ws["B10"] = "=_status!B2"
    ws["A10"].font = Font(bold=True)

    ws["A11"] = "当前步骤:"
    ws["B11"] = "=_status!B3"
    ws["A11"].font = Font(bold=True)

    ws["A12"] = "下一步:"
    ws["B12"] = "=_status!B4"
    ws["A12"].font = Font(bold=True)
    ws["B12"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[12].height = 32

    # 操作按钮
    _style_section(ws, 14, 1, "🎬 操作", span=8)
    _button(ws, 15, 1, "⚙️ 打开配置 sheet", "⚙️ 配置")
    _button(ws, 15, 4, "🧪 打开记录 sheet", "🧪 记录")
    _button(ws, 15, 7, "📊 打开看板 sheet", "📊 看板")
    ws.row_dimensions[15].height = 28

    # 启动 BO 按钮 — 点击一键跑 ask + refresh
    # ⚠️ 用占位符 {{BO_BAT}},init_new_project.py 生成项目时替换成绝对路径
    # (Excel 解析相对路径用的是 Excel 当前工作目录,不是 xlsx 目录 → 相对路径必失效)
    _button_cmd(ws, 17, 1, "🚀 启动 BO", "{{BO_BAT}}")
    ws["A18"] = "(一键跑 BO: 算新推荐 + 合并到 📊 看板,cmd 窗口会显示日志)"
    ws["A18"].font = HINT_FONT
    ws.merge_cells("A18:H18")

    # 单独 refresh 按钮 — 应急用(如果 📊 看板 没更新)
    _button_cmd(ws, 17, 4, "🔄 单独刷新", "{{BO_BAT}} refresh")
    ws["A19"] = "(仅把 cache 合并到 📊 看板,不跑 BO)"
    ws["A19"].font = HINT_FONT
    ws.merge_cells("A19:H19")

    # 生成记录表头按钮 — 填完 ⚙️ 配置后点这个,🧪 记录 的表头就按配置生成好了
    # (不用等 watchdog 触发,点了立刻生效)
    _button_cmd(ws, 20, 1, "📋 生成记录表头", "{{BO_BAT}} sync")
    ws["A21"] = "(按 ⚙️ 配置 的因子/目标重建 🧪 记录 表头 — 填完配置先点这个,再去 🧪 记录 填数据)"
    ws["A21"].font = HINT_FONT
    ws.merge_cells("A21:H21")

    # 收敛提示(下移到 row 23 — row 21 被"生成记录表头"按钮占用)
    _style_section(ws, 23, 1, "📈 收敛提示", span=8)
    ws["A24"] = "=_status!B10"
    ws["A24"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells("A24:H24")
    ws.row_dimensions[22].height = 40


def build_config_sheet(wb: Workbook) -> None:
    """Sheet 2: ⚙️ 配置 — 因子 / 目标 / 策略 / 约束(用户编辑)"""
    ws = wb["⚙️ 配置"]
    for col, w in enumerate([20, 14, 12, 10, 10, 28, 18, 18], start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    _style_title(ws, 1, 1, "⚙️ 项目配置 — 因子 / 目标 / 策略", span=8)

    # 面包屑
    ws["A2"] = "步骤 2/4 · 当前:配置"
    ws["A2"].font = HINT_FONT

    # --- Factors ---
    _style_section(ws, 4, 1, "📐 实验因子 (Factors)", span=8)
    _style_table_header(ws, 5, ["key", "type", "trainable", "low", "high", "categories", "unit", "notes"])
    # 4 行示例(可删可加)
    for r in range(6, 10):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = THIN_BORDER
            ws.cell(row=r, column=c).fill = INPUT_FILL
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="left", vertical="center")
    # type 列加数据验证
    # ⚠️ 范围必须限制在 Factors 区域内(row 6-11)。
    # 之前写 B6:B105 会一路跨到 Outputs(row 14+) / Strategy(row 21+) /
    # Constraints(row 33+),同事在那些区域的 B 列点单元格也会看到
    # continuous/discrete/categorical —— 这就是"三个表格下拉菜单不对"的根因。
    from openpyxl.worksheet.datavalidation import DataValidation
    dv_type = DataValidation(
        type="list",
        formula1='"continuous,discrete,categorical,categorical_descriptor,categorical_molecular"',
        allow_blank=True,
    )
    ws.add_data_validation(dv_type)
    dv_type.add("B6:B11")
    dv_tr = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws.add_data_validation(dv_tr)
    dv_tr.add("C6:C11")

    # 示例填 3 个因子(Suzuki 反应常用)
    ws["A6"] = "温度"
    ws["B6"] = "continuous"
    ws["C6"] = "TRUE"
    ws["D6"] = 20
    ws["E6"] = 100
    ws["G6"] = "°C"
    ws["A7"] = "催化剂用量"
    ws["B7"] = "continuous"
    ws["C7"] = "TRUE"
    ws["D7"] = 0.5
    ws["E7"] = 5.0
    ws["G7"] = "mol%"
    ws["A8"] = "反应时间"
    ws["B8"] = "continuous"
    ws["C8"] = "TRUE"
    ws["D8"] = 1
    ws["E8"] = 24
    ws["G8"] = "h"
    ws["A9"] = "溶剂"
    ws["B9"] = "categorical"
    ws["C9"] = "TRUE"
    ws["F9"] = "甲苯;THF;DMF;二氧六环"

    # --- Outputs ---
    _style_section(ws, 12, 1, "🎯 实验目标 (Outputs)", span=8)
    _style_table_header(ws, 13, ["key", "direction", "target", "lower_bound", "upper_bound", "unit", "", ""])
    for r in range(14, 17):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = THIN_BORDER
            ws.cell(row=r, column=c).fill = INPUT_FILL
    # ⚠️ 下拉范围必须限制在本区域内(row 14-16),
    # 不能像之前那样写 B14:B105 — 那会跨到 Strategy / Constraints 区域,
    # 同事在别的地方点单元格也会看到 maximize/minimize/target。
    dv_dir = DataValidation(type="list", formula1='"maximize,minimize,target"', allow_blank=True)
    ws.add_data_validation(dv_dir)
    dv_dir.add("B14:B16")
    # 示例 Yield
    ws["A14"] = "Yield"
    ws["B14"] = "maximize"
    ws["D14"] = 0
    ws["E14"] = 100
    ws["F14"] = "%"

    # --- Strategy ---
    _style_section(ws, 19, 1, "🧠 BO 策略 (Strategy)", span=8)
    _style_table_header(ws, 20, ["key", "value", "options", "说明", "", "", "", ""])
    strategy_rows = [
        ("策略类型", "sobo", "random;sobo;mobo;doe", "单目标 GP-BO(默认)"),
        ("采集函数", "qLogEI", "qLogEI;qLogNEI;qEI;qUCB", "标准 EI(默认)"),
        ("最少初始点数", 5, "≥3", "数据 < 此值用随机采样"),
        ("批大小", 3, "1-10", "每轮推荐的实验数"),
        ("收敛阈值", 0.01, "0.001-0.1", "最近 N 轮 rolling best 提升 < 此值视为收敛"),
        ("收敛窗口", 5, "3-10", "看最近几轮"),
        ("存模型", "FALSE", "TRUE;FALSE", "勾选后 models/ 留 GP 快照"),
    ]
    for i, (k, v, opt, desc) in enumerate(strategy_rows):
        r = 21 + i
        ws.cell(row=r, column=1, value=k).border = THIN_BORDER
        ws.cell(row=r, column=2, value=v).border = THIN_BORDER
        ws.cell(row=r, column=2).fill = INPUT_FILL
        ws.cell(row=r, column=3, value=opt).border = THIN_BORDER
        ws.cell(row=r, column=4, value=desc).border = THIN_BORDER
        ws.cell(row=r, column=4).font = HINT_FONT

    # ⚠️ Strategy 的 value 列下拉 — 每一行只对自己的 value 格生效
    # (之前 Strategy 完全没有下拉,同事只能手打,容易打错)
    from openpyxl.worksheet.datavalidation import DataValidation
    dv_strategy = DataValidation(type="list", formula1='"random,sobo,mobo,doe"', allow_blank=True)
    ws.add_data_validation(dv_strategy)
    dv_strategy.add("B21")  # 策略类型

    dv_acqf = DataValidation(type="list", formula1='"qLogEI,qLogNEI,qEI,qUCB"', allow_blank=True)
    ws.add_data_validation(dv_acqf)
    dv_acqf.add("B22")  # 采集函数

    dv_bool = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws.add_data_validation(dv_bool)
    dv_bool.add("B27")  # 存模型

    # --- Constraints (可选) ---
    _style_section(ws, 30, 1, "🔒 约束 (Constraints) — 一般不填", span=8)
    ws["A31"] = "高级用户用。复杂场景(如 温度+时间 总和不能 > 100)在这里加 linear_inequality。"
    ws["A31"].font = HINT_FONT
    ws.merge_cells("A31:H31")
    _style_table_header(ws, 32, ["name", "type", "features", "coefficients", "rhs", "", "", ""])
    for r in range(33, 36):
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = THIN_BORDER
            ws.cell(row=r, column=c).fill = INPUT_FILL

    # ⚠️ Constraints 的 type 列下拉 — 只在 row 33-35 生效
    dv_constraint = DataValidation(type="list", formula1='"linear_inequality,linear_equality"', allow_blank=True)
    ws.add_data_validation(dv_constraint)
    dv_constraint.add("B33:B35")


def build_record_sheet(wb: Workbook) -> None:
    """Sheet 3: 🧪 记录 — 实验数据(用户编辑)"""
    ws = wb["🧪 记录"]
    for col, w in enumerate([10, 12, 12, 12, 12, 12, 12, 10, 10, 12, 14], start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    _style_title(ws, 1, 1, "🧪 实验记录", span=11)
    ws["A2"] = "步骤 3/4 · 当前:记录 — 你填实验条件 + 实测 Yield"
    ws["A2"].font = HINT_FONT

    # 表头: trial_id | batch_id | <factors> | Yield | valid | source | created_at
    headers = ["trial_id", "batch_id", "温度", "催化剂用量", "反应时间", "溶剂", "Yield", "valid", "status", "来源", "录入时间"]
    _style_table_header(ws, 4, headers)
    # 8 行空白(可加)
    for r in range(5, 13):
        for c in range(1, 12):
            ws.cell(row=r, column=c).border = THIN_BORDER
            ws.cell(row=r, column=c).fill = INPUT_FILL
    # 数据验证
    from openpyxl.worksheet.datavalidation import DataValidation
    dv_valid = DataValidation(type="list", formula1='"0,1"', allow_blank=True)
    ws.add_data_validation(dv_valid)
    dv_valid.add("H5:H105")
    dv_status = DataValidation(type="list", formula1='"measured,pending,failed"', allow_blank=True)
    ws.add_data_validation(dv_status)
    dv_status.add("J5:J105")

    # 提示
    ws["A14"] = "💡 提示: Yield 单位 %; valid=0 表示失败,BO 不训练; status 通常 'measured'"
    ws["A14"].font = HINT_FONT
    ws.merge_cells("A14:K14")


def build_dashboard_sheet(wb: Workbook) -> None:
    """Sheet 4: 📊 看板 — 推荐/曲线/日志(系统写,用户只读)"""
    ws = wb["📊 看板"]
    for col, w in enumerate([12, 12, 12, 12, 12, 12, 14, 14, 14, 14], start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    _style_title(ws, 1, 1, "📊 看板 — 推荐 / 收敛 / 日志(系统写,你只读)", span=10)
    ws["A2"] = "步骤 4/4 · 当前:看板"
    ws["A2"].font = HINT_FONT

    # 本轮推荐区 — 支持最多 10 个候选(batch_size ≤ 10)
    _style_section(ws, 4, 1, "🎯 最新推荐", span=10)
    headers = ["trial_id", "温度", "催化剂用量", "反应时间", "溶剂", "Yield_pred_mean", "Yield_pred_sd", "acq_value", "round", "asked_at"]
    _style_table_header(ws, 5, headers)
    # 占位 10 行(row 6-15)— 同事 batch_size=10 时能完整展示
    for r in range(6, 16):
        for c in range(1, 11):
            ws.cell(row=r, column=c).border = THIN_BORDER
            ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="E2EFDA")  # 浅绿
    ws["A16"] = "(GP 预测仅供参考,实际 Yield 以你测的为准)"
    ws["A16"].font = HINT_FONT
    ws.merge_cells("A16:J16")

    # 收敛曲线区(数据范围 + 留白给 PNG)
    _style_section(ws, 18, 1, "📈 收敛曲线", span=10)
    _style_table_header(ws, 19, ["round", "best_so_far", "mean_of_best3", "n_measured", "n_failed", "", "", "", "", ""])
    for r in range(20, 50):
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = THIN_BORDER
    ws["G19"] = "(右侧留白区放 matplotlib 收敛 PNG)"
    ws["G19"].font = HINT_FONT

    # 历轮推荐区(累积,系统追加)
    _style_section(ws, 52, 1, "📜 历轮推荐(累积,只追加不删)", span=10)
    _style_table_header(ws, 53, ["round", "asked_at", "trial_id", "温度", "催化剂用量", "反应时间", "溶剂", "Yield_pred_mean", "Yield_pred_sd", "actual_Yield"])
    for r in range(54, 70):
        for c in range(1, 11):
            ws.cell(row=r, column=c).border = THIN_BORDER
    ws["A72"] = "💡 每轮 BO 推荐会追加到这里;实际 Yield 填到最后列 `actual_Yield`"
    ws["A72"].font = HINT_FONT
    ws.merge_cells("A72:J72")

    # 运行日志区
    _style_section(ws, 74, 1, "📋 运行日志(累积)", span=10)
    _style_table_header(ws, 75, ["at", "phase", "round", "elapsed_s", "strategy", "n_measured", "n_train", "status", "msg", ""])
    for r in range(76, 100):
        for c in range(1, 11):
            ws.cell(row=r, column=c).border = THIN_BORDER


def build_status_sheet(wb: Workbook) -> None:
    """隐藏 _status — Python 写,各 sheet 公式引用它"""
    ws = wb["_status"]
    # 默认值
    defaults = {
        "B1": "_status",  # 标识
        "B2": "awaiting_config",  # 状态
        "B3": "1/4",  # 当前步骤
        "B4": "切到 ⚙️ 配置,填写因子和目标",  # 下一步
        "B5": "<未命名项目>",  # 项目名
        "B6": "<未命名目录>",  # 项目目录
        "B7": 0,  # 当前轮次
        "B8": 0,  # 数据点数
        "B9": "",  # 最近 BO 时间
        "B10": "⏳ 暂无收敛数据(数据 ≥ 5 后开始评估)",  # 收敛提示
        "B11": "",  # 最近错误
    }
    for k, v in defaults.items():
        ws[k] = v
    # 命名区域(方便公式)
    wb.defined_names["status_state"] = DefinedName("status_state", attr_text="_status!$B$2")
    wb.defined_names["status_step"] = DefinedName("status_step", attr_text="_status!$B$3")
    wb.defined_names["status_next"] = DefinedName("status_next", attr_text="_status!$B$4")
    wb.defined_names["status_round"] = DefinedName("status_round", attr_text="_status!$B$7")
    wb.defined_names["status_ndata"] = DefinedName("status_ndata", attr_text="_status!$B$8")


def build_workbook() -> Workbook:
    wb = Workbook()
    # 默认 Sheet → 删掉,自己建
    wb.remove(wb.active)
    wb.create_sheet("🚀 开始")
    wb.create_sheet("⚙️ 配置")
    wb.create_sheet("🧪 记录")
    wb.create_sheet("📊 看板")
    wb.create_sheet("_status")

    build_start_sheet(wb)
    build_config_sheet(wb)
    build_record_sheet(wb)
    build_dashboard_sheet(wb)
    build_status_sheet(wb)

    # _status 隐藏
    wb["_status"].sheet_state = "hidden"

    # tab 顺序: 🚀 开始 → ⚙️ 配置 → 🧪 记录 → 📊 看板 → _status(隐藏)
    wb.active = 0  # 打开时落在 🚀 开始
    wb["🚀 开始"].sheet_properties.tabColor = "70AD47"  # 绿
    wb["⚙️ 配置"].sheet_properties.tabColor = "4472C4"  # 蓝
    wb["🧪 记录"].sheet_properties.tabColor = "ED7D31"  # 橙
    wb["📊 看板"].sheet_properties.tabColor = "FFC000"  # 黄

    return wb


def main():
    out = Path(__file__).resolve().parent.parent / "templates" / "project_template.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb = build_workbook()
    wb.save(out)
    print(f"✅ 生成: {out}")
    print(f"   sheet 数: {len(wb.sheetnames)} (含 1 个隐藏 _status)")
    print(f"   可见 sheet: {[s for s in wb.sheetnames if s != '_status']}")


if __name__ == "__main__":
    main()