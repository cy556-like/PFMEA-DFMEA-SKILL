#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PFMEA/DFMEA 报告生成器
======================

基于 AIAG & VDA FMEA 手册（2019 版）七步法，生成 FMEA 分析报告（.xlsx + .docx）。

依赖：
    - openpyxl（生成 Excel）
    - python-docx（生成 Word）

未安装时自动 pip install。

用法：
    python3 generate_fmea.py \\
        --fmea-type PFMEA \\
        --product "前保险杠总成（注塑件）" \\
        --customer "XX 汽车有限公司" \\
        --process-name "注塑成型" \\
        --process-steps "原料干燥,注塑成型,去毛刺,外观检验,包装" \\
        --template generic-fmea \\
        --auto-fill \\
        --output-dir ~/Desktop

⚠️ 评分基准（详见 references/sod_scoring_tables.md）：
    - S（严重度）：1-10，仅 S=10 用于"影响行车安全"
    - O（频度）：1-10，基于预防控制有效性
    - D（探测度）：1-10，基于探测方法成熟度
    - AP（行动优先级）：H/M/L，基于 S/O/D 组合查表
    - 严禁使用 RPN（=S×O×D），2019 版已废弃
"""

import argparse
import json
import os
import subprocess
import sys
import datetime
from pathlib import Path

# ============================================================
# 依赖检查与自动安装
# ============================================================

REQUIRED_PACKAGES = {
    "openpyxl": "openpyxl",
    "docx": "python-docx",  # 注意：导入名是 docx，包名是 python-docx
}


def ensure_packages():
    """检查并自动安装所需的 Python 包。"""
    import importlib

    for import_name, pip_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(import_name)
        except ImportError:
            print(f"[INFO] {pip_name} 未安装，正在自动安装...")
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", pip_name, "--quiet"]
                )
                print(f"[INFO] {pip_name} 安装完成。")
            except subprocess.CalledProcessError as e:
                print(f"[ERROR] {pip_name} 安装失败：{e}")
                print(f"        请手动执行：pip install {pip_name}")
                sys.exit(1)


ensure_packages()

import openpyxl
from openpyxl.styles import (
    Alignment,
    Border,
    Side,
    PatternFill,
    Font,
)
from openpyxl.utils import get_column_letter

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================
# AP 行动优先级矩阵（基于 S/O/D 查询）
# ============================================================

def get_ap_priority(s: int, o: int, d: int) -> str:
    """根据 S/O/D 评分查询 AP 行动优先级（H/M/L）。

    基于 AIAG & VDA FMEA 手册（2019 版）AP 矩阵实现。
    优先级顺序：S → O → D。
    """
    # S=1 必为 L
    if s == 1:
        return "L"

    # S=9-10: H（除非 O=1 且 D=1 才为 M）
    if s >= 9:
        if o == 1 and d == 1:
            return "M"
        return "H"

    # S=6-8
    if 6 <= s <= 8:
        if o >= 6:
            return "H"
        if 4 <= o <= 5:
            if d >= 7:
                return "H"
            return "M"
        if 2 <= o <= 3:
            return "M"
        if o == 1:
            return "L"

    # S=4-5
    if 4 <= s <= 5:
        if o >= 6:
            if d >= 7:
                return "H"
            return "M"
        if 4 <= o <= 5:
            return "M"
        if o <= 3:
            return "L"

    # S=2-3
    if 2 <= s <= 3:
        if o >= 6:
            if d >= 7:
                return "M"
            return "L"
        if 4 <= o <= 5:
            if s == 3:
                return "M"
            return "L"
        return "L"

    return "L"  # 兜底


# ============================================================
# 模板加载
# ============================================================

def load_template(template_slug: str, templates_dir: Path) -> dict:
    """加载指定模板的 template.json。"""
    template_path = templates_dir / template_slug / "template.json"
    if not template_path.exists():
        print(f"[WARN] 模板 {template_slug} 不存在，使用 generic-fmea")
        template_path = templates_dir / "generic-fmea" / "template.json"
    with open(template_path, "r", encoding="utf-8") as f:
        return json.load(f)


def auto_fill_failure_chain(chain: dict) -> dict:
    """自动填充失效链的 S/O/D 评分与 AP。

    当用户启用 auto_fill 模式时，从 chain 中的 _hint 字段读取建议值，
    并自动计算 AP。
    """
    s = chain.get("s_hint") if isinstance(chain.get("s_hint"), int) else chain.get("s", 5)
    o = chain.get("o_hint") if isinstance(chain.get("o_hint"), int) else chain.get("o", 5)
    d = chain.get("d_hint") if isinstance(chain.get("d_hint"), int) else chain.get("d", 5)
    ap = chain.get("ap_hint") or get_ap_priority(s, o, d)

    chain["s"] = s
    chain["o"] = o
    chain["d"] = d
    chain["ap"] = ap
    return chain


def substitute_placeholders(text: str, context: dict) -> str:
    """替换文本中的占位符 {product_name} / {customer} 等。"""
    if not isinstance(text, str):
        return text
    for k, v in context.items():
        text = text.replace("{" + k + "}", str(v))
    return text


# ============================================================
# Excel 生成
# ============================================================

def create_excel_report(data: dict, output_path: Path):
    """生成 FMEA Excel 报告（7 Sheet）。"""
    wb = openpyxl.Workbook()

    # 样式定义
    header_fill = PatternFill("solid", fgColor="305496")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    cell_font = Font(size=10)
    border = Border(
        left=Side(style="thin", color="B4B4B4"),
        right=Side(style="thin", color="B4B4B4"),
        top=Side(style="thin", color="B4B4B4"),
        bottom=Side(style="thin", color="B4B4B4"),
    )
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # AP 颜色
    ap_colors = {
        "H": PatternFill("solid", fgColor="FF6B6B"),  # 红
        "M": PatternFill("solid", fgColor="FFD93D"),  # 黄
        "L": PatternFill("solid", fgColor="6BCB77"),  # 绿
    }

    # ---------- Sheet 1: 表头 ----------
    ws1 = wb.active
    ws1.title = "1.表头"
    ws1.column_dimensions["A"].width = 22
    ws1.column_dimensions["B"].width = 50

    header_info = [
        ("FMEA 类型", data.get("fmea_type", "")),
        ("公司名称", data.get("company", "（待填写）")),
        ("顾客名称", data.get("customer", "")),
        ("项目名称", data.get("product_name", "")),
        ("项目编号", data.get("project_no", "（待填写）")),
        ("系统层级", data.get("system_level", "—")),
        ("设计责任", data.get("design_responsibility", "—")),
        ("工艺名称", data.get("process_name", "—")),
        ("制造地址", data.get("manufacturing_site", "（待填写）")),
        ("团队成员", data.get("team_members", "（待填写）")),
        ("FMEA 开始日期", data.get("start_date", datetime.date.today().isoformat())),
        ("FMEA 修订日期", datetime.date.today().isoformat()),
        ("FMEA 编号", data.get("fmea_no", f"FMEA-{datetime.date.today().strftime('%Y%m%d')}-001")),
        ("使用模板", data.get("template", "generic-fmea")),
        ("参考标准", "AIAG & VDA FMEA Handbook 2019"),
    ]

    ws1.cell(row=1, column=1, value="FMEA 表头信息").font = Font(bold=True, size=14)
    ws1.cell(row=1, column=1).alignment = align_center
    ws1.merge_cells("A1:B1")
    ws1.cell(row=1, column=1).fill = header_fill
    ws1.cell(row=1, column=1).font = header_font

    for i, (k, v) in enumerate(header_info, start=2):
        ws1.cell(row=i, column=1, value=k).font = Font(bold=True, size=10)
        ws1.cell(row=i, column=1).fill = PatternFill("solid", fgColor="D9E1F2")
        ws1.cell(row=i, column=1).border = border
        ws1.cell(row=i, column=1).alignment = align_left
        ws1.cell(row=i, column=2, value=v).font = cell_font
        ws1.cell(row=i, column=2).border = border
        ws1.cell(row=i, column=2).alignment = align_left

    # ---------- Sheet 2: 结构分析 ----------
    ws2 = wb.create_sheet("2.结构分析")
    ws2.column_dimensions["A"].width = 15
    ws2.column_dimensions["B"].width = 35
    ws2.column_dimensions["C"].width = 50

    ws2.cell(row=1, column=1, value="步骤二：结构分析").font = Font(bold=True, size=14)
    ws2.merge_cells("A1:C1")
    ws2.cell(row=1, column=1).fill = header_fill
    ws2.cell(row=1, column=1).font = header_font
    ws2.cell(row=1, column=1).alignment = align_center

    headers = ["层级", "要素名称", "说明"]
    for col, h in enumerate(headers, start=1):
        c = ws2.cell(row=2, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = border

    # 填充结构树
    structure = data.get("structure_tree", [])
    for i, item in enumerate(structure, start=3):
        ws2.cell(row=i, column=1, value=item.get("level", "")).border = border
        ws2.cell(row=i, column=1).alignment = align_center
        ws2.cell(row=i, column=2, value=item.get("name", "")).border = border
        ws2.cell(row=i, column=2).alignment = align_left
        ws2.cell(row=i, column=3, value=item.get("description", "")).border = border
        ws2.cell(row=i, column=3).alignment = align_left

    # ---------- Sheet 3: 功能分析 ----------
    ws3 = wb.create_sheet("3.功能分析")
    ws3.column_dimensions["A"].width = 15
    ws3.column_dimensions["B"].width = 25
    ws3.column_dimensions["C"].width = 55

    ws3.cell(row=1, column=1, value="步骤三：功能分析").font = Font(bold=True, size=14)
    ws3.merge_cells("A1:C1")
    ws3.cell(row=1, column=1).fill = header_fill
    ws3.cell(row=1, column=1).font = header_font
    ws3.cell(row=1, column=1).alignment = align_center

    headers = ["层级", "要素名称", "功能描述"]
    for col, h in enumerate(headers, start=1):
        c = ws3.cell(row=2, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = border

    functions = data.get("function_tree", [])
    for i, item in enumerate(functions, start=3):
        ws3.cell(row=i, column=1, value=item.get("level", "")).border = border
        ws3.cell(row=i, column=1).alignment = align_center
        ws3.cell(row=i, column=2, value=item.get("name", "")).border = border
        ws3.cell(row=i, column=2).alignment = align_left
        ws3.cell(row=i, column=3, value=item.get("function", "")).border = border
        ws3.cell(row=i, column=3).alignment = align_left

    # ---------- Sheet 4: 失效分析 ----------
    ws4 = wb.create_sheet("4.失效分析")
    ws4.column_dimensions["A"].width = 8
    ws4.column_dimensions["B"].width = 45
    ws4.column_dimensions["C"].width = 35
    ws4.column_dimensions["D"].width = 45

    ws4.cell(row=1, column=1, value="步骤四：失效分析（失效链 FE→FM→FC）").font = Font(bold=True, size=14)
    ws4.merge_cells("A1:D1")
    ws4.cell(row=1, column=1).fill = header_fill
    ws4.cell(row=1, column=1).font = header_font
    ws4.cell(row=1, column=1).alignment = align_center

    headers = ["序号", "失效影响（FE）", "失效模式（FM）", "失效起因（FC）"]
    for col, h in enumerate(headers, start=1):
        c = ws4.cell(row=2, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = border

    chains = data.get("failure_chains", [])
    for i, chain in enumerate(chains, start=3):
        ws4.cell(row=i, column=1, value=i - 2).border = border
        ws4.cell(row=i, column=1).alignment = align_center
        ws4.cell(row=i, column=2, value=chain.get("fe", "")).border = border
        ws4.cell(row=i, column=2).alignment = align_left
        ws4.cell(row=i, column=3, value=chain.get("fm", "")).border = border
        ws4.cell(row=i, column=3).alignment = align_left
        ws4.cell(row=i, column=4, value=chain.get("fc", "")).border = border
        ws4.cell(row=i, column=4).alignment = align_left
        # 自适应行高
        ws4.row_dimensions[i].height = max(60, len(chain.get("fe", "")) / 2)

    # ---------- Sheet 5: 风险分析 ----------
    ws5 = wb.create_sheet("5.风险分析")
    col_widths = [6, 30, 25, 35, 35, 6, 6, 6, 6, 8]
    for i, w in enumerate(col_widths, start=1):
        ws5.column_dimensions[get_column_letter(i)].width = w

    ws5.cell(row=1, column=1, value="步骤五：风险分析（S/O/D 评级 + AP）").font = Font(bold=True, size=14)
    ws5.merge_cells("A1:J1")
    ws5.cell(row=1, column=1).fill = header_fill
    ws5.cell(row=1, column=1).font = header_font
    ws5.cell(row=1, column=1).alignment = align_center

    headers = ["序号", "失效影响（FE）", "失效模式（FM）", "预防控制（PC）", "探测控制（DC）", "S", "O", "D", "AP", "特殊特性"]
    for col, h in enumerate(headers, start=1):
        c = ws5.cell(row=2, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = border

    for i, chain in enumerate(chains, start=3):
        s = chain.get("s", 5)
        o = chain.get("o", 5)
        d = chain.get("d", 5)
        ap = chain.get("ap", get_ap_priority(s, o, d))

        ws5.cell(row=i, column=1, value=i - 2).border = border
        ws5.cell(row=i, column=1).alignment = align_center
        ws5.cell(row=i, column=2, value=chain.get("fe", "")).border = border
        ws5.cell(row=i, column=2).alignment = align_left
        ws5.cell(row=i, column=3, value=chain.get("fm", "")).border = border
        ws5.cell(row=i, column=3).alignment = align_left
        ws5.cell(row=i, column=4, value=chain.get("pc", "")).border = border
        ws5.cell(row=i, column=4).alignment = align_left
        ws5.cell(row=i, column=5, value=chain.get("dc", "")).border = border
        ws5.cell(row=i, column=5).alignment = align_left
        ws5.cell(row=i, column=6, value=s).border = border
        ws5.cell(row=i, column=6).alignment = align_center
        ws5.cell(row=i, column=7, value=o).border = border
        ws5.cell(row=i, column=7).alignment = align_center
        ws5.cell(row=i, column=8, value=d).border = border
        ws5.cell(row=i, column=8).alignment = align_center

        ap_cell = ws5.cell(row=i, column=9, value=ap)
        ap_cell.border = border
        ap_cell.alignment = align_center
        ap_cell.font = Font(bold=True)
        if ap in ap_colors:
            ap_cell.fill = ap_colors[ap]

        # 特殊特性识别
        sc = ""
        if s >= 9:
            sc = "CC"
        elif s == 8 and ap in ("H", "M"):
            sc = "SC"
        sc_cell = ws5.cell(row=i, column=10, value=sc)
        sc_cell.border = border
        sc_cell.alignment = align_center
        if sc:
            sc_cell.font = Font(bold=True, color="FF0000")

        ws5.row_dimensions[i].height = max(60, len(chain.get("fe", "")) / 2)

    # ---------- Sheet 6: 优化措施 ----------
    ws6 = wb.create_sheet("6.优化措施")
    col_widths = [6, 30, 12, 50, 15, 15, 12, 8, 8, 8, 8]
    for i, w in enumerate(col_widths, start=1):
        ws6.column_dimensions[get_column_letter(i)].width = w

    ws6.cell(row=1, column=1, value="步骤六：优化措施").font = Font(bold=True, size=14)
    ws6.merge_cells("A1:K1")
    ws6.cell(row=1, column=1).fill = header_fill
    ws6.cell(row=1, column=1).font = header_font
    ws6.cell(row=1, column=1).alignment = align_center

    headers = ["序号", "失效模式", "措施类型", "措施描述", "责任人", "截止日期", "状态", "措施后 S", "措施后 O", "措施后 D", "措施后 AP"]
    for col, h in enumerate(headers, start=1):
        c = ws6.cell(row=2, column=col, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = border

    measures = data.get("optimization_measures", [])
    for i, m in enumerate(measures, start=3):
        ws6.cell(row=i, column=1, value=i - 2).border = border
        ws6.cell(row=i, column=1).alignment = align_center
        ws6.cell(row=i, column=2, value=m.get("fm", "")).border = border
        ws6.cell(row=i, column=2).alignment = align_left
        ws6.cell(row=i, column=3, value=m.get("type", "")).border = border
        ws6.cell(row=i, column=3).alignment = align_center
        ws6.cell(row=i, column=4, value=m.get("description", "")).border = border
        ws6.cell(row=i, column=4).alignment = align_left
        ws6.cell(row=i, column=5, value=m.get("owner", "（待指定）")).border = border
        ws6.cell(row=i, column=5).alignment = align_center
        ws6.cell(row=i, column=6, value=m.get("due_date", "（待指定）")).border = border
        ws6.cell(row=i, column=6).alignment = align_center
        ws6.cell(row=i, column=7, value=m.get("status", "已建议")).border = border
        ws6.cell(row=i, column=7).alignment = align_center
        # 措施后 S/O/D/AP（待实施后填写）
        for col in [8, 9, 10, 11]:
            ws6.cell(row=i, column=col, value="—").border = border
            ws6.cell(row=i, column=col).alignment = align_center

    # ---------- Sheet 7: 风险矩阵 ----------
    ws7 = wb.create_sheet("7.风险矩阵")
    ws7.cell(row=1, column=1, value="步骤七：风险矩阵（S×O 热力图）").font = Font(bold=True, size=14)
    ws7.merge_cells("A1:L1")
    ws7.cell(row=1, column=1).fill = header_fill
    ws7.cell(row=1, column=1).font = header_font
    ws7.cell(row=1, column=1).alignment = align_center

    # S×O 矩阵（10x10）
    ws7.cell(row=3, column=1, value="S\\O").font = Font(bold=True)
    ws7.cell(row=3, column=1).fill = PatternFill("solid", fgColor="305496")
    ws7.cell(row=3, column=1).font = header_font
    ws7.cell(row=3, column=1).alignment = align_center
    ws7.cell(row=3, column=1).border = border

    for o in range(1, 11):
        c = ws7.cell(row=3, column=o + 1, value=o)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = border

    for s in range(10, 0, -1):
        row_idx = 14 - s
        c = ws7.cell(row=row_idx, column=1, value=s)
        c.font = header_font
        c.fill = header_fill
        c.alignment = align_center
        c.border = border
        for o in range(1, 11):
            ap = get_ap_priority(s, o, 5)  # D=5 中位数
            cell = ws7.cell(row=row_idx, column=o + 1, value=ap)
            cell.alignment = align_center
            cell.border = border
            cell.font = Font(bold=True)
            if ap in ap_colors:
                cell.fill = ap_colors[ap]

    # 图例
    ws7.cell(row=16, column=1, value="图例：").font = Font(bold=True)
    ws7.cell(row=16, column=2, value="H（高）").fill = ap_colors["H"]
    ws7.cell(row=16, column=2).alignment = align_center
    ws7.cell(row=16, column=3, value="M（中）").fill = ap_colors["M"]
    ws7.cell(row=16, column=3).alignment = align_center
    ws7.cell(row=16, column=4, value="L（低）").fill = ap_colors["L"]
    ws7.cell(row=16, column=4).alignment = align_center

    # 风险统计
    ws7.cell(row=18, column=1, value="风险统计").font = Font(bold=True, size=12)
    h_count = sum(1 for c in chains if c.get("ap") == "H")
    m_count = sum(1 for c in chains if c.get("ap") == "M")
    l_count = sum(1 for c in chains if c.get("ap") == "L")
    cc_count = sum(1 for c in chains if c.get("s", 0) >= 9)
    sc_count = sum(1 for c in chains if c.get("s", 0) == 8 and c.get("ap") in ("H", "M"))

    stats = [
        ("失效链总数", len(chains)),
        ("AP=H（高优先级）", h_count),
        ("AP=M（中优先级）", m_count),
        ("AP=L（低优先级）", l_count),
        ("CC（关键特性，S≥9）", cc_count),
        ("SC（特殊特性，S=8 且 AP=H/M）", sc_count),
    ]
    for i, (k, v) in enumerate(stats, start=19):
        ws7.cell(row=i, column=1, value=k).font = Font(bold=True)
        ws7.cell(row=i, column=2, value=v).alignment = align_center
        ws7.cell(row=i, column=2).font = Font(bold=True, size=12, color="305496")

    # 保存
    wb.save(output_path)
    print(f"[OK] Excel 报告已生成：{output_path}")


# ============================================================
# Word 生成
# ============================================================

def create_docx_report(data: dict, output_path: Path):
    """生成 FMEA Word 报告（7 章）。"""
    doc = Document()

    # 设置默认字体
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style.font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    # 标题
    title = doc.add_heading(f"{data.get('fmea_type', 'FMEA')} 分析报告", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 副标题
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = sub.add_run(f"产品：{data.get('product_name', '')}\n客户：{data.get('customer', '')}\n日期：{datetime.date.today().isoformat()}")
    run.font.size = Pt(11)

    doc.add_paragraph()

    # ---------- 第 1 章：概述 ----------
    doc.add_heading("第一章 概述", level=1)

    doc.add_paragraph(
        f"本报告基于 AIAG & VDA FMEA 手册（2019 版）七步法编制，覆盖"
        f"{data.get('fmea_type', '')} 分析的全过程。分析对象为"
        f"{data.get('product_name', '')}，涉及客户 {data.get('customer', '')}。"
    )

    p = doc.add_paragraph()
    p.add_run("项目基本信息：").bold = True

    info_table = doc.add_table(rows=8, cols=2)
    info_table.style = "Light Grid Accent 1"
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    info_items = [
        ("FMEA 类型", data.get("fmea_type", "")),
        ("产品名称", data.get("product_name", "")),
        ("客户名称", data.get("customer", "")),
        ("项目编号", data.get("project_no", "（待填写）")),
        ("系统层级", data.get("system_level", "—")),
        ("工艺名称", data.get("process_name", "—")),
        ("团队成员", data.get("team_members", "（待填写）")),
        ("FMEA 编号", data.get("fmea_no", f"FMEA-{datetime.date.today().strftime('%Y%m%d')}-001")),
    ]

    for i, (k, v) in enumerate(info_items):
        info_table.cell(i, 0).text = k
        info_table.cell(i, 1).text = str(v)
        info_table.cell(i, 0).paragraphs[0].runs[0].bold = True if info_table.cell(i, 0).paragraphs[0].runs else False

    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("参考标准：").bold = True
    p.add_run("AIAG & VDA FMEA Handbook 2019、IATF 16949:2016")

    # ---------- 第 2 章：结构分析 ----------
    doc.add_heading("第二章 结构分析", level=1)

    doc.add_paragraph(
        "结构分析将分析对象分解为可视化的结构，识别系统要素及其相互关系。"
        "本步骤为后续功能分析和失效分析提供基础。"
    )

    structure = data.get("structure_tree", [])
    if structure:
        doc.add_paragraph("结构树：")
        for item in structure:
            indent = "  " * (int(item.get("level_indent", 0)) or 0)
            doc.add_paragraph(f"{indent}• {item.get('name', '')}：{item.get('description', '')}")

    # ---------- 第 3 章：功能分析 ----------
    doc.add_heading("第三章 功能分析", level=1)

    doc.add_paragraph(
        "功能分析明确每个系统要素/工序的功能及其要求。"
        "功能的完整定义会使失效分析更全面。功能定义遵循「主动动词 + 可测量名词」的格式。"
    )

    functions = data.get("function_tree", [])
    if functions:
        doc.add_paragraph("功能清单：")
        func_table = doc.add_table(rows=1 + len(functions), cols=3)
        func_table.style = "Light Grid Accent 1"
        func_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr = func_table.rows[0].cells
        hdr[0].text = "层级"
        hdr[1].text = "要素名称"
        hdr[2].text = "功能描述"
        for c in hdr:
            for p in c.paragraphs:
                for r in p.runs:
                    r.bold = True

        for i, item in enumerate(functions, start=1):
            func_table.cell(i, 0).text = str(item.get("level", ""))
            func_table.cell(i, 1).text = str(item.get("name", ""))
            func_table.cell(i, 2).text = str(item.get("function", ""))

    # ---------- 第 4 章：失效分析 ----------
    doc.add_heading("第四章 失效分析", level=1)

    doc.add_paragraph(
        "失效分析识别每个功能的潜在失效影响（FE）、失效模式（FM）和失效起因（FC），"
        "形成失效链 FE→FM→FC。失效模式包括 7 种类型：功能丧失、功能退化、功能间歇、"
        "部分功能丧失、非预期功能、功能超范围、功能延迟。"
    )

    chains = data.get("failure_chains", [])
    if chains:
        doc.add_paragraph(f"共识别 {len(chains)} 条失效链：")

        for i, chain in enumerate(chains, start=1):
            doc.add_heading(f"失效链 #{i}", level=2)

            t = doc.add_table(rows=3, cols=2)
            t.style = "Light Grid Accent 1"
            t.cell(0, 0).text = "失效影响（FE）"
            t.cell(0, 1).text = chain.get("fe", "")
            t.cell(1, 0).text = "失效模式（FM）"
            t.cell(1, 1).text = chain.get("fm", "")
            t.cell(2, 0).text = "失效起因（FC）"
            t.cell(2, 1).text = chain.get("fc", "")

            for r in t.rows:
                r.cells[0].paragraphs[0].runs[0].bold = True if r.cells[0].paragraphs[0].runs else False

    # ---------- 第 5 章：风险分析 ----------
    doc.add_heading("第五章 风险分析", level=1)

    doc.add_paragraph(
        "风险分析通过评估严重度（S）、频度（O）、探测度（D）来估计风险，"
        "并基于 2019 版 AP 行动优先级矩阵判定 H/M/L 三档优先级。"
        "本步骤不再使用 RPN（=S×O×D），改用 AP 矩阵。"
    )

    if chains:
        doc.add_paragraph("风险评级表：")

        risk_table = doc.add_table(rows=1 + len(chains), cols=7)
        risk_table.style = "Light Grid Accent 1"
        risk_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr = risk_table.rows[0].cells
        headers = ["序号", "失效模式", "S", "O", "D", "AP", "特殊特性"]
        for i, h in enumerate(headers):
            hdr[i].text = h
            for p in hdr[i].paragraphs:
                for r in p.runs:
                    r.bold = True

        for i, chain in enumerate(chains, start=1):
            s = chain.get("s", 5)
            o = chain.get("o", 5)
            d = chain.get("d", 5)
            ap = chain.get("ap", get_ap_priority(s, o, d))

            sc = ""
            if s >= 9:
                sc = "CC"
            elif s == 8 and ap in ("H", "M"):
                sc = "SC"

            risk_table.cell(i, 0).text = str(i)
            risk_table.cell(i, 1).text = chain.get("fm", "")
            risk_table.cell(i, 2).text = str(s)
            risk_table.cell(i, 3).text = str(o)
            risk_table.cell(i, 4).text = str(d)
            risk_table.cell(i, 5).text = ap
            risk_table.cell(i, 6).text = sc

        # 风险统计
        doc.add_paragraph()
        p = doc.add_paragraph()
        p.add_run("风险统计：").bold = True

        h_count = sum(1 for c in chains if c.get("ap") == "H")
        m_count = sum(1 for c in chains if c.get("ap") == "M")
        l_count = sum(1 for c in chains if c.get("ap") == "L")
        cc_count = sum(1 for c in chains if c.get("s", 0) >= 9)
        sc_count = sum(1 for c in chains if c.get("s", 0) == 8 and c.get("ap") in ("H", "M"))

        doc.add_paragraph(f"• 失效链总数：{len(chains)}")
        doc.add_paragraph(f"• AP=H（高优先级）：{h_count} 项")
        doc.add_paragraph(f"• AP=M（中优先级）：{m_count} 项")
        doc.add_paragraph(f"• AP=L（低优先级）：{l_count} 项")
        doc.add_paragraph(f"• CC（关键特性，S≥9）：{cc_count} 项")
        doc.add_paragraph(f"• SC（特殊特性，S=8 且 AP=H/M）：{sc_count} 项")

    # ---------- 第 6 章：优化措施 ----------
    doc.add_heading("第六章 优化措施", level=1)

    doc.add_paragraph(
        "优化措施分为三类：预防控制（PC）改进、探测控制（DC）改进、设计/过程变更。"
        "对于 S=9-10 且 AP=H/M 的失效影响，建议至少由管理层评审。"
    )

    measures = data.get("optimization_measures", [])
    if measures:
        doc.add_paragraph(f"共 {len(measures)} 项优化措施：")

        m_table = doc.add_table(rows=1 + len(measures), cols=5)
        m_table.style = "Light Grid Accent 1"
        m_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr = m_table.rows[0].cells
        headers = ["序号", "措施类型", "措施描述", "责任人", "截止日期"]
        for i, h in enumerate(headers):
            hdr[i].text = h
            for p in hdr[i].paragraphs:
                for r in p.runs:
                    r.bold = True

        for i, m in enumerate(measures, start=1):
            m_table.cell(i, 0).text = str(i)
            m_table.cell(i, 1).text = m.get("type", "")
            m_table.cell(i, 2).text = m.get("description", "")
            m_table.cell(i, 3).text = m.get("owner", "（待指定）")
            m_table.cell(i, 4).text = m.get("due_date", "（待指定）")

    # ---------- 第 7 章：结论与建议 ----------
    doc.add_heading("第七章 结论与建议", level=1)

    doc.add_paragraph(
        "本 FMEA 分析基于 AIAG & VDA FMEA 手册（2019 版）七步法编制，"
        f"共识别 {len(chains)} 条失效链，其中高优先级（AP=H）{h_count} 项、"
        f"中优先级（AP=M）{m_count} 项、低优先级（AP=L）{l_count} 项。"
    )

    doc.add_paragraph("建议：")
    doc.add_paragraph("1. 优先处理 AP=H 的失效链，制定详细措施并跟踪实施；")
    doc.add_paragraph("2. 对于 S=9-10 的失效影响（CC 关键特性），必须由管理层评审；")
    doc.add_paragraph("3. 将识别出的 CC/SC 同步到控制计划（CP）；")
    doc.add_paragraph("4. 措施实施后重新评估 S/O/D，验证 AP 是否降低；")
    doc.add_paragraph("5. FMEA 文档应定期评审和更新，特别是在产品设计/工艺变更时。")

    # 保存
    doc.save(output_path)
    print(f"[OK] Word 报告已生成：{output_path}")


# ============================================================
# 主流程
# ============================================================

def main():
    parser = argparse.ArgumentParser(description="PFMEA/DFMEA 报告生成器")
    parser.add_argument("--fmea-type", required=True, choices=["DFMEA", "PFMEA"], help="FMEA 类型")
    parser.add_argument("--product", required=True, help="产品名称")
    parser.add_argument("--customer", required=True, help="客户名称")
    parser.add_argument("--project-no", default="", help="项目编号")
    parser.add_argument("--system-level", default="", help="系统层级（DFMEA）")
    parser.add_argument("--design-responsibility", default="", help="设计责任（DFMEA）")
    parser.add_argument("--process-name", default="", help="工艺名称（PFMEA）")
    parser.add_argument("--process-steps", default="", help="工序清单（逗号分隔，PFMEA）")
    parser.add_argument("--manufacturing-site", default="", help="制造地址（PFMEA）")
    parser.add_argument("--team", default="", help="团队成员")
    parser.add_argument("--template", default="generic-fmea", help="模板 slug")
    parser.add_argument("--auto-fill", action="store_true", help="自动填充 S/O/D 评分")
    parser.add_argument("--output-dir", default=".", help="输出目录")
    parser.add_argument("--failure-chains-json", default="", help="自定义失效链 JSON 文件路径")

    args = parser.parse_args()

    # 模板目录（脚本所在目录的上级 templates/）
    script_dir = Path(__file__).resolve().parent
    templates_dir = script_dir.parent / "templates"

    # 加载模板
    template = load_template(args.template, templates_dir)
    print(f"[INFO] 使用模板：{template.get('name', args.template)}")

    # 构建上下文
    context = {
        "product_name": args.product,
        "customer": args.customer,
        "process_name": args.process_name,
    }

    # 失效链
    chains = template.get("failure_chains_template", [])
    if args.failure_chains_json:
        with open(args.failure_chains_json, "r", encoding="utf-8") as f:
            chains = json.load(f)

    # 占位符替换
    chains = [
        {k: substitute_placeholders(v, context) if isinstance(v, str) else v for k, v in c.items()}
        for c in chains
    ]

    # 自动填充
    if args.auto_fill:
        chains = [auto_fill_failure_chain(c) for c in chains]
        print(f"[INFO] 已自动填充 {len(chains)} 条失效链的 S/O/D 评分")
    else:
        # 仍计算 AP（如果 S/O/D 已提供）
        for c in chains:
            s = c.get("s", c.get("s_hint", 5))
            o = c.get("o", c.get("o_hint", 5))
            d = c.get("d", c.get("d_hint", 5))
            if isinstance(s, int) and isinstance(o, int) and isinstance(d, int):
                c["s"] = s
                c["o"] = o
                c["d"] = d
                c["ap"] = c.get("ap_hint") or get_ap_priority(s, o, d)

    # 构建报告数据
    data = {
        "fmea_type": args.fmea_type,
        "company": "（待填写）",
        "customer": args.customer,
        "product_name": args.product,
        "project_no": args.project_no,
        "system_level": args.system_level,
        "design_responsibility": args.design_responsibility,
        "process_name": args.process_name,
        "manufacturing_site": args.manufacturing_site,
        "team_members": args.team,
        "template": args.template,
        "structure_tree": [],  # 可扩展：从模板提取
        "function_tree": [],   # 可扩展：从模板提取
        "failure_chains": chains,
        "optimization_measures": [
            {
                "fm": c.get("fm", ""),
                "type": "PC+DC 改进",
                "description": c.get("pc", "") + " | " + c.get("dc", ""),
                "owner": "（待指定）",
                "due_date": "（待指定）",
                "status": "已建议",
            }
            for c in chains
            if c.get("ap") in ("H", "M")
        ],
    }

    # 输出路径
    output_dir = Path(args.output_dir).expanduser()
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx_path = output_dir / f"FMEA_{args.fmea_type}_{timestamp}.xlsx"
    docx_path = output_dir / f"FMEA_{args.fmea_type}_{timestamp}.docx"

    # 生成报告
    create_excel_report(data, xlsx_path)
    create_docx_report(data, docx_path)

    print(f"\n{'=' * 60}")
    print(f"✅ FMEA 报告生成完成！")
    print(f"{'=' * 60}")
    print(f"  Excel: {xlsx_path}")
    print(f"  Word:  {docx_path}")
    print(f"\n  失效链数: {len(chains)}")
    h_count = sum(1 for c in chains if c.get("ap") == "H")
    m_count = sum(1 for c in chains if c.get("ap") == "M")
    l_count = sum(1 for c in chains if c.get("ap") == "L")
    print(f"  AP=H: {h_count}  AP=M: {m_count}  AP=L: {l_count}")


if __name__ == "__main__":
    main()
