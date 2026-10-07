# -*- coding: utf-8 -*-
"""Generate per-talent Excel: each talent gets their own sheet with contract details.

Reads combined_results.json (existing 7 contracts) + 猴哥 structured data,
produces a 9-sheet .xlsx with per-talent card layout.
"""

import json
import os
from datetime import datetime
from pathlib import Path
import config

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Styling constants (from gen_brokerage_report.py / make_report.py) ─────────
GREEN = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFEB9C")
RED = PatternFill("solid", fgColor="FFC7CE")
GRAY = PatternFill("solid", fgColor="D9D9D9")
BLUE_HINT = PatternFill("solid", fgColor="DAEEF3")
HEADER_FILL = PatternFill("solid", fgColor="2F5496")
SUBHDR_FILL = PatternFill("solid", fgColor="BDD7EE")
COLHDR_FILL = PatternFill("solid", fgColor="1F4E79")
SECTION_FILL = PatternFill("solid", fgColor="2F5496")

thin = Side(style="thin", color="AAAAAA")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
bottom_border = Border(left=thin, right=thin, top=thin, bottom=Side(style="medium", color="2F5496"))

HEADER_FONT = Font(name="微软雅黑", size=14, bold=True, color="FFFFFF")
SUBHDR_FONT = Font(name="微软雅黑", size=10, color="2F5496")
COLHDR_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
DATA_FONT = Font(name="微软雅黑", size=10)
LABEL_FONT = Font(name="微软雅黑", size=10, bold=True, color="1F4E79")
TITLE_FONT = Font(name="微软雅黑", size=16, bold=True, color="2F5496")
SECTION_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
CHECK_FONT = Font(name="微软雅黑", size=14, bold=True, color="006100")
CROSS_FONT = Font(name="微软雅黑", size=14, bold=True, color="9C0006")

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)
LABEL_ALIGN = Alignment(horizontal="right", vertical="center", wrap_text=True)


def _sc(ws, row, col, value, fill=None, font=None, align=None):
    """Set cell value + style."""
    c = ws.cell(row=row, column=col, value=value)
    if fill:
        c.fill = fill
    if font:
        c.font = font
    if align:
        c.alignment = align
    c.border = border
    return c


def _sc_merge(ws, row1, col1, row2, col2, value, fill=None, font=None, align=None):
    """Merge cells and set value + style."""
    ws.merge_cells(start_row=row1, start_column=col1, end_row=row2, end_column=col2)
    c = ws.cell(row=row1, column=col1, value=value)
    if fill:
        c.fill = fill
    if font:
        c.font = font
    if align:
        c.alignment = align
    # Apply border to all cells in merge range
    for r in range(row1, row2 + 1):
        for cl in range(col1, col2 + 1):
            ws.cell(row=r, column=cl).border = border
    return c


def write_section_header(ws, row, title):
    """Write a blue section header spanning columns A-B."""
    _sc_merge(ws, row, 1, row, 2, title, fill=SECTION_FILL, font=SECTION_FONT, align=CENTER)
    ws.row_dimensions[row].height = 26


def write_field_row(ws, row, label, value, value_fill=None, value_font=None):
    """Write a label (col A, right-aligned bold) + value (col B, left-aligned)."""
    _sc(ws, row, 1, label, font=LABEL_FONT, align=LABEL_ALIGN)
    vf = value_font if value_font else DATA_FONT
    _sc(ws, row, 2, value, fill=value_fill, font=vf, align=LEFT_TOP)
    ws.row_dimensions[row].height = max(22, 16 * (str(value).count('\n') + 1))


def build_talent_sheet(wb, contract, metadata):
    """Build one talent sheet with card layout (sections A-D)."""
    name = metadata.get("talent_name", contract.get("contract_name", ""))
    # Sheet name: max 31 chars for Excel
    sheet_name = name[:31]

    ws = wb.create_sheet(sheet_name)

    # Column widths
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 72

    row = 1

    # ── Title ──
    _sc_merge(ws, row, 1, row, 2, f"📋 {name} — 经纪合同关键条款摘录",
              font=TITLE_FONT, align=LEFT)
    ws.row_dimensions[row].height = 36
    row += 1
    _sc_merge(ws, row, 1, row, 2,
              f"提取日期: {datetime.now().strftime('%Y-%m-%d')}  |  合同文件: {metadata.get('source_file', contract.get('contract_name', ''))}",
              font=SUBHDR_FONT, align=LEFT)
    ws.row_dimensions[row].height = 20
    row += 2

    # ═══════════════ Section A: 基本信息 ═══════════════
    write_section_header(ws, row, "一、基本信息")
    row += 1
    info_fields = [
        ("达人名称", metadata.get("talent_name", "")),
        ("签约公司", metadata.get("company", "")),
        ("签约账号", metadata.get("account", "")),
        ("签订日期", metadata.get("signing_date", "")),
        ("合同文件名", contract.get("contract_name", "")),
    ]
    for label, value in info_fields:
        write_field_row(ws, row, label, value)
        row += 1

    row += 1  # blank row between sections

    # ═══════════════ Section B: 合同期限 ═══════════════
    write_section_header(ws, row, "二、合同期限")
    row += 1

    ct = contract.get("contract_term", {})
    if ct.get("found"):
        term_fields = [
            ("起算日期", ct.get("start_date", "—")),
            ("终止日期", ct.get("end_date", "—")),
            ("初始期限", f"{ct.get('term_years', '—')}年" if ct.get("term_years") else "—"),
            ("期限类型", ct.get("term_type", "—")),
            ("提前终止条件", ct.get("early_termination", "—")),
            ("原文摘录", ct.get("source_text_raw", "—")),
        ]
    else:
        term_fields = [("状态", "⚠ 未在合同中找到明确约定")]

    for label, value in term_fields:
        vf = DATA_FONT
        vfill = None
        if label == "原文摘录":
            vf = Font(name="微软雅黑", size=9, italic=True, color="555555")
            vfill = GRAY
        if label == "状态" and "未在" in str(value):
            vfill = RED
        write_field_row(ws, row, label, value, value_fill=vfill, value_font=vf)
        if label == "原文摘录":
            ws.row_dimensions[row].height = 60
        row += 1

    row += 1

    # ═══════════════ Section C: 续约条件 ═══════════════
    write_section_header(ws, row, "三、续约条件")
    row += 1

    rc = contract.get("renewal_conditions", {})
    if rc.get("found"):
        is_unconditional = rc.get("is_unconditional_renewal", False)
        # Also detect unconditional from text
        renewal_text = (rc.get("renewal_type", "") + rc.get("source_text_raw", ""))
        if "无条件" in renewal_text:
            is_unconditional = True

        unconditional_display = "✓ 是（无条件自动续约）" if is_unconditional else "✗ 否（有附加条件）"
        ufill = GREEN if is_unconditional else YELLOW
        ufont = CHECK_FONT if is_unconditional else CROSS_FONT

        renewal_fields = [
            ("续约类型", rc.get("renewal_type", "—")),
            ("无条件续约", unconditional_display),
            ("通知期限", f"{rc.get('notice_period_days', '—')}天" if rc.get("notice_period_days") else "未约定（自动续约无需通知）"),
            ("通知方式", rc.get("notice_method", "—")),
            ("续约条件摘要", rc.get("renewal_conditions_summary", "—")),
            ("退出/不续约机制", rc.get("opt_out_mechanism", "—")),
            ("原文摘录", rc.get("source_text_raw", "—")),
        ]
    else:
        renewal_fields = [("状态", "⚠ 未在合同中找到明确约定")]

    for label, value in renewal_fields:
        vf = DATA_FONT
        vfill = None
        if label == "无条件续约":
            vf = ufont
            vfill = ufill
        if label == "原文摘录":
            vf = Font(name="微软雅黑", size=9, italic=True, color="555555")
            vfill = GRAY
        if label == "状态" and "未在" in str(value):
            vfill = RED
        write_field_row(ws, row, label, value, value_fill=vfill, value_font=vf)
        if label == "原文摘录":
            ws.row_dimensions[row].height = 60
        row += 1

    row += 1

    # ═══════════════ Section D: 附件收益分配 ★ ═══════════════
    write_section_header(ws, row, "四、附件收益分配 ★重点★")
    row += 1

    ps = contract.get("profit_sharing", {})
    if ps.get("found"):
        adjustments = ps.get("split_adjustments", [])
        # Build gradient display
        gradients = []
        for i, adj in enumerate(adjustments):
            income = adj.get("income_type", "").replace("全部商业合作", "").strip("（）()")
            ratio = adj.get("ratio", "")
            gradients.append(f"第{i+1}梯度 ({income}): {ratio}" if income else f"第{i+1}梯度: {ratio}")

        profit_fields = [
            ("基础分成比", ps.get("base_split_ratio", "—")),
            ("分成梯度详情", "\n".join(gradients) if gradients else ps.get("base_split_ratio", "—")),
            ("成本扣除方式", ps.get("cost_deduction_policy", "—")),
            ("结算周期", ps.get("settlement_cycle", "—")),
            ("结算截止日", f"每月{ps.get('settlement_deadline_days', '—')}日前" if ps.get("settlement_deadline_days") else "—"),
            ("税费承担", ps.get("tax_handling", "—")),
            ("保底条款", ps.get("minimum_guarantee", "—")),
            ("附件引用", ps.get("appendix_reference", "无")),
            ("附件是否在合同文件中", "✓ 是（已在docx/pdf中找到）" if ps.get("appendix_found_in_docx") else "✗ 否（需获取纸质版附件）"),
            ("原文摘录", ps.get("source_text_raw", "—")),
        ]
    else:
        profit_fields = [("状态", "⚠ 未在合同中找到明确约定")]

    for label, value in profit_fields:
        vf = DATA_FONT
        vfill = None
        if label == "分成梯度详情":
            ws.row_dimensions[row].height = max(22, 18 * len(gradients)) if gradients else 22
        if label == "原文摘录":
            vf = Font(name="微软雅黑", size=9, italic=True, color="555555")
            vfill = GRAY
        if label == "附件是否在合同文件中":
            if "否" in str(value):
                vfill = YELLOW
                vf = Font(name="微软雅黑", size=10, bold=True, color="9C6500")
            else:
                vfill = GREEN
                vf = Font(name="微软雅黑", size=10, color="006100")
        if label == "状态" and "未在" in str(value):
            vfill = RED
        write_field_row(ws, row, label, value, value_fill=vfill, value_font=vf)
        if label == "原文摘录":
            ws.row_dimensions[row].height = 60
        row += 1

    # Freeze top rows
    ws.freeze_panes = "A1"
    return ws


def build_overview_sheet(wb, all_contracts, all_metadata):
    """Sheet 1: 综合对比总览 — horizontal comparison matrix."""
    ws = wb.create_sheet("综合对比总览")

    headers = [
        "达人名称", "签约公司", "签订日期", "初始期限(年)", "实际总期限",
        "无条件续约", "续约类型", "第一梯度分成", "后续梯度分成", "结算周期", "备注"
    ]
    col_count = len(headers)

    # Title
    _sc_merge(ws, 1, 1, 1, col_count,
              "演艺经纪合同 — 达人关键条款综合对比总览",
              font=TITLE_FONT, align=LEFT)
    ws.row_dimensions[1].height = 36
    _sc_merge(ws, 2, 1, 2, col_count,
              f"提取日期: {datetime.now().strftime('%Y-%m-%d')}  |  共 {len(all_contracts)} 位达人",
              font=SUBHDR_FONT, align=LEFT)
    ws.row_dimensions[2].height = 22

    # Headers
    for ci, h in enumerate(headers, 1):
        _sc(ws, 3, ci, h, fill=COLHDR_FILL, font=COLHDR_FONT, align=CENTER)
    ws.row_dimensions[3].height = 28
    ws.freeze_panes = "A4"

    for ri in range(len(all_contracts)):
        row = 4 + ri
        contract = all_contracts[ri]
        meta = all_metadata[ri]
        ct = contract.get("contract_term", {})
        rc = contract.get("renewal_conditions", {})
        ps = contract.get("profit_sharing", {})

        # Detect unconditional renewal
        is_uncond = False
        renewal_text = (rc.get("renewal_type", "") + rc.get("source_text_raw", ""))
        if "无条件" in renewal_text:
            is_uncond = True
        uncond_str = "✓ 无条件" if is_uncond else ("✗ 有附加条件" if rc.get("found") else "—")

        # Profit sharing summary
        adjustments = ps.get("split_adjustments", [])
        first_gradient = adjustments[0].get("ratio", "—") if adjustments else ps.get("base_split_ratio", "—")
        later_gradients = " → ".join(a.get("ratio", "") for a in adjustments[1:]) if len(adjustments) > 1 else "—"

        # Notes
        notes = []
        if ps.get("appendix_reference") and not ps.get("appendix_found_in_docx"):
            notes.append("⚠ 附件未在文件中找到")
        if ct.get("confidence") == "low":
            notes.append("期限需人工确认")
        if rc.get("confidence") == "low":
            notes.append("续约条款不完整")

        values = [
            meta.get("talent_name", contract.get("contract_name", "")),
            meta.get("company", ""),
            meta.get("signing_date", ""),
            ct.get("term_years", "—"),
            meta.get("actual_total_years", "—"),
            uncond_str,
            rc.get("renewal_type", "—"),
            first_gradient,
            later_gradients,
            ps.get("settlement_cycle", "—"),
            "; ".join(notes) if notes else "",
        ]
        aligns = [LEFT, LEFT, CENTER, CENTER, CENTER, CENTER, LEFT, CENTER, CENTER, CENTER, LEFT]
        for ci, v in enumerate(values, 1):
            al = aligns[ci - 1]
            _sc(ws, row, ci, v, font=DATA_FONT, align=al)
        ws.row_dimensions[row].height = 32

        # Color-code conditional cells
        if is_uncond:
            ws.cell(row=row, column=6).fill = GREEN
        else:
            ws.cell(row=row, column=6).fill = YELLOW
        if notes:
            ws.cell(row=row, column=11).fill = YELLOW

    # Auto-width
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells:
            if cell.value:
                text = str(cell.value)
                length = sum(2 if ord(c) > 127 else 1 for c in text)
                max_len = max(max_len, length)
        ws.column_dimensions[col_letter].width = min(max(max_len * 1.1 + 2, 8), 40)

    # Widen some columns
    ws.column_dimensions['A'].width = 20
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['G'].width = 28
    ws.column_dimensions['H'].width = 20
    ws.column_dimensions['I'].width = 24
    ws.column_dimensions['K'].width = 22

    return ws


# ── 猴哥 contract data (structured from PDF) ────────────────────────────────
HOUGE_CONTRACT = {
    "contract_name": "猴哥说车-侯焜山-南京车节奏-2024年签",
    "contract_term": {
        "found": True,
        "source_paragraphs": [],
        "source_text_raw": "本协议有效期及双方合作期限为二十年，自2024年2月1日起，至2044年1月31日止。协议到期后，无条件自动延续二十年至2064年1月31日。",
        "start_date": "2024年2月1日",
        "end_date": "2044年1月31日（无条件自动延续至2064年1月31日）",
        "term_years": 20,
        "term_type": "固定期限（实际40年：20+20无条件自动延续）",
        "early_termination": "①甲方无故未付收益超3个月或严重损害名誉/身心健康→乙方书面解约\n②连续3次评估不达标→甲方无责解约\n③身体/心理疾病、信誉/社会评价严重降低→甲方书面解约\n④乙方私下承接演艺活动超3次→甲方有权解约",
        "confidence": "high"
    },
    "renewal_conditions": {
        "found": True,
        "source_paragraphs": [],
        "source_text_raw": "协议到期后，无条件自动延续二十年至2064年1月31日。乙方承诺合同期限届满后优先与甲方续约，否则按照第八条第3款承担违约责任（合作总收益100%-150%违约金）。",
        "renewal_type": "无条件自动续约20年 + 优先续约权",
        "is_unconditional_renewal": True,
        "notice_period_days": None,
        "notice_method": "未约定（自动续约无需通知）",
        "renewal_conditions_summary": "首次20年届满后无条件自动延续20年至2064年；全部届满后乙方须优先与甲方续约，否则按合作总收益100%-150%支付违约金",
        "opt_out_mechanism": "无条件自动延续阶段无退出机制；40年届满后若拒绝优先续约→违约金（合作总收益100%-150%）",
        "confidence": "high"
    },
    "profit_sharing": {
        "found": True,
        "source_paragraphs": [],
        "source_tables": [],
        "source_text_raw": "附件1：2024年2月1日-2025年2月28日 甲50%:乙50%；2025年3月1日起 甲40%:乙60%。净利收益=甲方实际收到合作账号商业合作客户支付费用-账号经营产生的所有显性成本。月度结算，每月8日前出结算单。甲方纳税后按比例支付乙方，乙方个税由甲方代扣代缴。",
        "base_split_ratio": "甲50%:乙50%（第1年）→ 甲40%:乙60%（第2年起）",
        "split_adjustments": [
            {"income_type": "全部商业合作收入（2024.2.1-2025.2.28）", "ratio": "甲50%:乙50%"},
            {"income_type": "全部商业合作收入（2025.3.1起）", "ratio": "甲40%:乙60%"}
        ],
        "cost_deduction_policy": "净利收益模式：甲方实际到账收入 - 账号经营所有显性成本 = 净利润基数",
        "settlement_cycle": "月度",
        "settlement_deadline_days": 8,
        "tax_handling": "甲方按法律法规纳税→税后按附件比例支付乙方；乙方个人所得税自行承担，甲方代扣代缴",
        "minimum_guarantee": "无保底条款（甲方未收到客户款→不对乙方承担结算义务）",
        "appendix_reference": "附件1《短视频平台艺人经纪协议之附件1》（含收益分配比例表、合作账号信息：抖音猴哥说车3679.2万粉、小红书猴哥说车36.2万粉）",
        "appendix_found_in_docx": True,
        "confidence": "high"
    }
}

# ── Talent metadata mapping ──────────────────────────────────────────────────
TALENT_META = [
    {
        "talent_name": "大黄",
        "company": "南京水火土网络科技有限公司",
        "account": "抖音「大黄」",
        "signing_date": "2023年8月1日",
        "actual_total_years": "40年（20+20自动延续）",
        "source_file": "大黄经纪约.docx",
    },
    {
        "talent_name": "良田",
        "company": "南京水火土网络科技有限公司",
        "account": "抖音「良田」（483.9万粉）",
        "signing_date": "2023年3月1日",
        "actual_total_years": "3年（收入>1000万则自动延至5年）",
        "source_file": "良田经纪约.docx",
    },
    {
        "talent_name": "阿飞",
        "company": "南京车节奏文化传媒有限公司",
        "account": "—（未在提取中体现）",
        "signing_date": "2024年10月28日",
        "actual_total_years": "40年（20+20自动延续）",
        "source_file": "阿飞-20年.docx",
    },
    {
        "talent_name": "小鱼海棠（徐翔宇）",
        "company": "南京水火土网络科技有限公司",
        "account": "抖音「小鱼海棠」等",
        "signing_date": "2021年5月7日",
        "actual_total_years": "5年（已到期/即将到期）",
        "source_file": "小鱼海棠21年签的经纪约（徐翔宇）—水火土(1).docx",
    },
    {
        "talent_name": "小鱼",
        "company": "南京水火土网络科技有限公司",
        "account": "抖音「小鱼海棠」(1986.1w) / 小红书(428.8w) 等",
        "signing_date": "2023年8月1日",
        "actual_total_years": "40年（20+20自动延续）",
        "source_file": "小鱼经纪约.docx",
    },
    {
        "talent_name": "刘冠宇",
        "company": "北京车节奏文化传媒有限公司",
        "account": "—（未在提取中体现）",
        "signing_date": "2023年8月1日",
        "actual_total_years": "40年（20+20自动延续）",
        "source_file": "北京车节奏-刘冠宇-新.docx",
    },
    {
        "talent_name": "顾猛",
        "company": "车节奏文化传媒有限公司",
        "account": "—（未在提取中体现）",
        "signing_date": "2023年7月1日",
        "actual_total_years": "40年（20+20自动延续）",
        "source_file": "顾猛-车节奏最新pdf.docx",
    },
    {
        "talent_name": "猴哥说车（侯焜山）",
        "company": "南京车节奏文化传媒有限公司",
        "account": "抖音「猴哥说车」(3679.2万粉) / 小红书「猴哥说车」(36.2万粉)",
        "signing_date": "2024年2月1日",
        "actual_total_years": "40年（20+20自动延续）",
        "source_file": "猴哥说车-侯焜山-南京车节奏-2024年签.pdf",
    },
]


def generate_report(existing_json: str, output_path: str):
    """Generate the per-talent Excel workbook."""
    # Load existing 7 contracts
    with open(existing_json, "r", encoding="utf-8") as f:
        existing_data = json.load(f)

    if isinstance(existing_data, list):
        existing_contracts = existing_data
    elif isinstance(existing_data, dict):
        existing_contracts = existing_data.get("contracts", existing_data.get("data", list(existing_data.values())))
    else:
        raise ValueError(f"Unexpected data format: {type(existing_data)}")

    # Combine: 8 contracts
    all_contracts = list(existing_contracts) + [HOUGE_CONTRACT]
    all_metadata = TALENT_META

    # Verify counts match
    if len(all_contracts) != len(all_metadata):
        print(f"WARNING: {len(all_contracts)} contracts but {len(all_metadata)} metadata entries")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Sheet 1: Overview
    build_overview_sheet(wb, all_contracts, all_metadata)

    # Sheets 2-9: Per talent
    for i, (contract, meta) in enumerate(zip(all_contracts, all_metadata)):
        build_talent_sheet(wb, contract, meta)
        print(f"  Sheet {i+2}: {meta.get('talent_name', contract.get('contract_name', '?'))}")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    print(f"\nReport saved to: {output_path}")
    print(f"  Sheets: {wb.sheetnames}")
    print(f"  Total talents: {len(all_contracts)}")


if __name__ == "__main__":
    import sys

    input_json = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.join(config.BASE, "_brokerage_extract", "_combined_results.json")
    output_path = sys.argv[2] if len(sys.argv) > 2 else \
        os.path.join(config.BASE, "brokerage_excels", "经纪约关键条款提取对比_达人视图_20260710.xlsx")

    generate_report(input_json, output_path)
