# -*- coding: utf-8 -*-
"""Generate PPT report: 演艺经纪合同关键条款分析 — 8位达人对比.

Reads combined_results.json (7 contracts) + 猴哥 structured data,
produces a 12-slide .pptx with per-talent detail pages + overview.
"""

import json
import os
from datetime import datetime
import config

from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── Color palette (consistent with Excel styling) ──────────────────────────
BLUE_DARK = RGBColor(0x2F, 0x54, 0x96)
BLUE_MID = RGBColor(0xBD, 0xD7, 0xEE)
BLUE_DEEP = RGBColor(0x1F, 0x4E, 0x79)
GREEN = RGBColor(0xC6, 0xEF, 0xCE)
GREEN_DARK = RGBColor(0x00, 0x61, 0x00)
YELLOW = RGBColor(0xFF, 0xEB, 0x9C)
YELLOW_DARK = RGBColor(0x9C, 0x65, 0x00)
RED = RGBColor(0xFF, 0xC7, 0xCE)
RED_DARK = RGBColor(0x9C, 0x00, 0x06)
GRAY_LIGHT = RGBColor(0xF2, 0xF2, 0xF2)
GRAY_MID = RGBColor(0xD9, 0xD9, 0xD9)
GRAY_TEXT = RGBColor(0x55, 0x55, 0x55)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
TABLE_HDR_BG = BLUE_DEEP
TABLE_ROW_ALT = RGBColor(0xE8, 0xF0, 0xF8)

# ── Font helpers ───────────────────────────────────────────────────────────
FONT_FAMILY = "微软雅黑"


def _run(para, text, size=Pt(11), bold=False, color=BLACK, italic=False):
    """Add a run to a paragraph."""
    run = para.add_run()
    run.text = str(text)
    run.font.name = FONT_FAMILY
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.italic = italic
    return run


def _add_textbox(slide, left, top, width, height, fill_color=None):
    """Add a textbox shape with optional solid fill."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    if fill_color:
        tb.fill.solid()
        tb.fill.fore_color.rgb = fill_color
    else:
        tb.fill.background()
    tb.line.fill.background()  # no border
    return tb


def _add_paragraph(tf, text, size=Pt(11), bold=False, color=BLACK,
                   align=PP_ALIGN.LEFT, space_before=Pt(2), space_after=Pt(2),
                   italic=False, level=0):
    """Add a paragraph to a text frame, return it."""
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = space_before
    p.space_after = space_after
    p.level = level
    _run(p, text, size=size, bold=bold, color=color, italic=italic)
    return p


def _add_section_title(slide, left, top, width, text, size=Pt(13)):
    """Add a blue-bg section title bar (like a mini header)."""
    tb = _add_textbox(slide, left, top, width, Inches(0.32), fill_color=BLUE_DARK)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _run(p, text, size=size, bold=True, color=WHITE)


def _add_field(tf, label, value, label_size=Pt(10), value_size=Pt(10),
               label_color=BLUE_DARK, value_color=BLACK, value_bold=False,
               italic=False):
    """Add a label: value line to a text frame."""
    p = tf.add_paragraph()
    p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(1)
    p.space_after = Pt(1)
    _run(p, f"{label}：", size=label_size, bold=True, color=label_color)
    _run(p, str(value) if value else "—", size=value_size, bold=value_bold,
         color=value_color, italic=italic)


def _add_color_tag(tf, label, is_positive, true_text="✓ 是（无条件自动续约）",
                   false_text="✗ 否（有附加条件）"):
    """Add a color-coded tag line (green=good, yellow/red=warning)."""
    p = tf.add_paragraph()
    p.alignment = PP_ALIGN.LEFT
    p.space_before = Pt(1)
    p.space_after = Pt(1)
    _run(p, f"{label}：", size=Pt(10), bold=True, color=BLUE_DARK)
    if is_positive:
        _run(p, true_text, size=Pt(12), bold=True, color=GREEN_DARK)
    else:
        _run(p, false_text, size=Pt(12), bold=True, color=RED_DARK)


def _style_table_cell(cell, text, font_size=Pt(9), bold=False, color=BLACK,
                      fill=None, align=PP_ALIGN.CENTER):
    """Style a single table cell."""
    cell.text = ""
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    _run(p, str(text), size=font_size, bold=bold, color=color)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


# ── Slide builders ─────────────────────────────────────────────────────────


def build_cover(prs):
    """Slide 1: Cover / title slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

    # Full-slide dark blue background
    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = BLUE_DARK

    # Title
    tb = _add_textbox(slide, Inches(1.5), Inches(2.0), Inches(10), Inches(1.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _run(p, "演艺经纪合同关键条款分析", size=Pt(36), bold=True, color=WHITE)

    # Subtitle
    tb2 = _add_textbox(slide, Inches(1.5), Inches(3.5), Inches(10), Inches(0.8))
    tf2 = tb2.text_frame
    p2 = tf2.paragraphs[0]
    p2.alignment = PP_ALIGN.CENTER
    _run(p2, "8位短视频达人经纪合同对比研究", size=Pt(20), bold=False, color=RGBColor(0xDA, 0xEE, 0xF3))

    # Date line
    tb3 = _add_textbox(slide, Inches(1.5), Inches(4.5), Inches(10), Inches(0.5))
    tf3 = tb3.text_frame
    p3 = tf3.paragraphs[0]
    p3.alignment = PP_ALIGN.CENTER
    _run(p3, f"提取日期：{datetime.now().strftime('%Y年%m月%d日')}",
         size=Pt(12), bold=False, color=RGBColor(0xA0, 0xC0, 0xE0))

    # Decorative line
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(3), Inches(5.0), Inches(7), Pt(3)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = WHITE
    line.line.fill.background()


def build_overview_table(prs, all_contracts, all_metadata):
    """Slide 2: 8-dalent horizontal comparison table."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Title bar
    _add_section_title(slide, Inches(0.5), Inches(0.25), Inches(12.33),
                       "一、达人合同关键条款综合对比总览", size=Pt(16))

    headers = [
        "达人名称", "签约公司", "签订日期", "合同期限",
        "无条件\n续约", "第一梯度分成", "后续梯度", "结算周期"
    ]
    col_widths = [Inches(1.8), Inches(2.2), Inches(1.2), Inches(1.5),
                  Inches(1.0), Inches(1.5), Inches(1.8), Inches(1.0)]

    rows = len(all_contracts) + 1  # +1 for header
    cols = len(headers)
    tbl_left = Inches(0.3)
    tbl_top = Inches(0.8)
    tbl_width = sum(w for w in col_widths)
    tbl_height = Inches(0.35) * rows

    table = slide.shapes.add_table(rows, cols, tbl_left, tbl_top,
                                   tbl_width, tbl_height).table

    # Set column widths
    for ci, w in enumerate(col_widths):
        table.columns[ci].width = w

    # Header row
    for ci, h in enumerate(headers):
        _style_table_cell(table.cell(0, ci), h, font_size=Pt(9), bold=True,
                          color=WHITE, fill=TABLE_HDR_BG)

    # Data rows
    for ri in range(len(all_contracts)):
        contract = all_contracts[ri]
        meta = all_metadata[ri]
        ct = contract.get("contract_term", {})
        rc = contract.get("renewal_conditions", {})
        ps = contract.get("profit_sharing", {})

        # Unconditional renewal detection
        is_uncond = False
        renewal_text = (rc.get("renewal_type", "") + rc.get("source_text_raw", ""))
        if "无条件" in renewal_text:
            is_uncond = True

        # Term display
        term_years = ct.get("term_years", "—")
        actual = meta.get("actual_total_years", f"{term_years}年" if term_years else "—")

        # Profit sharing
        adjustments = ps.get("split_adjustments", [])
        first_grad = adjustments[0].get("ratio", "—") if adjustments else ps.get("base_split_ratio", "—")
        later = " → ".join(a.get("ratio", "") for a in adjustments[1:]) if len(adjustments) > 1 else "—"

        uncond_str = "✓ 无条件" if is_uncond else ("✗ 附条件" if rc.get("found") else "—")

        row_data = [
            meta.get("talent_name", contract.get("contract_name", "")),
            meta.get("company", ""),
            meta.get("signing_date", ""),
            str(actual),
            uncond_str,
            first_grad,
            later,
            ps.get("settlement_cycle", "—"),
        ]

        row_fill = GRAY_LIGHT if ri % 2 == 0 else None
        for ci, val in enumerate(row_data):
            align = PP_ALIGN.CENTER if ci >= 2 else PP_ALIGN.LEFT
            _style_table_cell(table.cell(ri + 1, ci), val, font_size=Pt(8),
                              color=BLACK, fill=row_fill, align=align)

        # Color-code unconditional column
        uc_cell = table.cell(ri + 1, 4)
        if is_uncond:
            uc_cell.fill.solid()
            uc_cell.fill.fore_color.rgb = GREEN
        elif rc.get("found"):
            uc_cell.fill.solid()
            uc_cell.fill.fore_color.rgb = YELLOW

    # Footnote
    tb = _add_textbox(slide, Inches(0.3), Inches(6.5), Inches(12), Inches(0.4))
    tf = tb.text_frame
    _add_paragraph(tf, "注：绿色=无条件自动续约；黄色=有附加条件续约/协商续约。收益分成为甲（公司）:乙（达人）比例。",
                   size=Pt(8), color=GRAY_TEXT)


def build_talent_detail(prs, contract, metadata):
    """Build one talent detail slide (two-column layout)."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    name = metadata.get("talent_name", contract.get("contract_name", ""))
    ct = contract.get("contract_term", {})
    rc = contract.get("renewal_conditions", {})
    ps = contract.get("profit_sharing", {})

    # ── Top title bar ──
    tb = _add_textbox(slide, Inches(0.3), Inches(0.15), Inches(12.7), Inches(0.5),
                      fill_color=BLUE_DARK)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    _run(p, f"  {name}  —  经纪合同关键条款", size=Pt(20), bold=True, color=WHITE)

    # ── Left column: 基本信息 + 合同期限 + 续约条件 ──
    left_x = Inches(0.3)
    left_w = Inches(5.8)
    row_y = Inches(0.8)

    # -- 基本信息 --
    _add_section_title(slide, left_x, row_y, left_w, "基本信息")
    row_y += Inches(0.4)
    tb_info = _add_textbox(slide, left_x, row_y, left_w, Inches(0.9))
    tf_info = tb_info.text_frame
    tf_info.word_wrap = True
    p0 = tf_info.paragraphs[0]
    _run(p0, "", size=Pt(2))  # dummy first paragraph
    _add_field(tf_info, "达人名称", metadata.get("talent_name", ""))
    _add_field(tf_info, "签约公司", metadata.get("company", ""))
    _add_field(tf_info, "签约账号", metadata.get("account", ""), value_size=Pt(9))
    _add_field(tf_info, "签订日期", metadata.get("signing_date", ""))
    row_y += Inches(1.1)

    # -- 合同期限 --
    _add_section_title(slide, left_x, row_y, left_w, "合同期限")
    row_y += Inches(0.4)
    tb_term = _add_textbox(slide, left_x, row_y, left_w, Inches(1.4))
    tf_term = tb_term.text_frame
    tf_term.word_wrap = True
    p0 = tf_term.paragraphs[0]
    _run(p0, "", size=Pt(2))

    if ct.get("found"):
        _add_field(tf_term, "起算日期", ct.get("start_date", "—"))
        _add_field(tf_term, "终止日期", ct.get("end_date", "—"))
        _add_field(tf_term, "实际总期限", metadata.get("actual_total_years", f"{ct.get('term_years', '—')}年"))
        _add_field(tf_term, "期限类型", ct.get("term_type", "—"))
        early_term = ct.get("early_termination", "—")
        if early_term and len(early_term) > 80:
            early_term = early_term[:80] + "…"
        _add_field(tf_term, "提前终止", early_term, value_size=Pt(8), italic=True)
    else:
        _add_field(tf_term, "状态", "⚠ 未在合同中找到明确约定", value_color=RED_DARK)
    row_y += Inches(1.6)

    # -- 续约条件 --
    _add_section_title(slide, left_x, row_y, left_w, "续约条件")
    row_y += Inches(0.4)
    tb_renewal = _add_textbox(slide, left_x, row_y, left_w, Inches(1.2))
    tf_renewal = tb_renewal.text_frame
    tf_renewal.word_wrap = True
    p0 = tf_renewal.paragraphs[0]
    _run(p0, "", size=Pt(2))

    if rc.get("found"):
        # Detect unconditional
        is_uncond = False
        renewal_text = (rc.get("renewal_type", "") + rc.get("source_text_raw", ""))
        if "无条件" in renewal_text:
            is_uncond = True

        _add_field(tf_renewal, "续约类型", rc.get("renewal_type", "—"), value_bold=True)
        _add_color_tag(tf_renewal, "无条件续约", is_uncond)
        _add_field(tf_renewal, "续约条件摘要", rc.get("renewal_conditions_summary", "—"),
                   value_size=Pt(8), italic=True)
        opt_out = rc.get("opt_out_mechanism", "—")
        if opt_out and len(opt_out) > 60:
            opt_out = opt_out[:60] + "…"
        _add_field(tf_renewal, "退出/不续约机制", opt_out, value_size=Pt(8), italic=True)
    else:
        _add_field(tf_renewal, "状态", "⚠ 未在合同中找到明确约定", value_color=RED_DARK)

    # ── Right column: 收益分配 ★ ──
    right_x = Inches(6.4)
    right_w = Inches(6.6)
    row_y = Inches(0.8)

    _add_section_title(slide, right_x, row_y, right_w, "附件收益分配 ★重点★")
    row_y += Inches(0.4)
    tb_profit = _add_textbox(slide, right_x, row_y, right_w, Inches(5.2))
    tf_profit = tb_profit.text_frame
    tf_profit.word_wrap = True
    p0 = tf_profit.paragraphs[0]
    _run(p0, "", size=Pt(2))

    if ps.get("found"):
        # Build gradient display
        adjustments = ps.get("split_adjustments", [])
        gradients = []
        for i, adj in enumerate(adjustments):
            income = adj.get("income_type", "").replace("全部商业合作", "").strip("（）()")
            ratio = adj.get("ratio", "")
            notes = adj.get("notes", "")
            if income:
                line = f"第{i+1}梯度 ({income}): {ratio}"
            else:
                line = f"第{i+1}梯度: {ratio}"
            if notes:
                line += f" [{notes}]"
            gradients.append(line)

        _add_field(tf_profit, "基础分成比", ps.get("base_split_ratio", "—"), value_bold=True,
                   value_size=Pt(11), value_color=BLUE_DEEP)
        for g_text in gradients:
            _add_field(tf_profit, f"  ├ 梯度详情", g_text, value_size=Pt(9))

        _add_field(tf_profit, "成本扣除方式", ps.get("cost_deduction_policy", "—"), value_size=Pt(9))
        _add_field(tf_profit, "结算周期", ps.get("settlement_cycle", "—"))
        if ps.get("settlement_deadline_days"):
            _add_field(tf_profit, "结算截止日", f"每月{ps.get('settlement_deadline_days')}日前")
        _add_field(tf_profit, "税费承担", ps.get("tax_handling", "—"), value_size=Pt(8), italic=True)
        _add_field(tf_profit, "保底条款", ps.get("minimum_guarantee", "—"))
        _add_field(tf_profit, "附件引用", ps.get("appendix_reference", "无"), value_size=Pt(8))

        # Appendix found status
        if ps.get("appendix_found_in_docx"):
            _add_field(tf_profit, "附件是否在合同中", "✓ 是（已在合同文件中找到）",
                       value_color=GREEN_DARK, value_bold=True)
        else:
            _add_field(tf_profit, "附件是否在合同中", "✗ 否（需获取纸质版附件）",
                       value_color=RED_DARK, value_bold=True)

        # Source text excerpt
        src = ps.get("source_text_raw", "—")
        if src and len(src) > 150:
            src = src[:150] + "…"
        _add_field(tf_profit, "原文摘录", src, value_size=Pt(8), italic=True,
                   value_color=GRAY_TEXT)
    else:
        _add_field(tf_profit, "状态", "⚠ 未在合同中找到明确约定", value_color=RED_DARK)

    # ── Footer line ──
    tb_footer = _add_textbox(slide, Inches(0.3), Inches(7.0), Inches(12), Inches(0.3))
    tf_footer = tb_footer.text_frame
    _add_paragraph(tf_footer, f"合同文件：{metadata.get('source_file', contract.get('contract_name', ''))}",
                   size=Pt(7), color=GRAY_TEXT)


def build_summary(prs, all_contracts, all_metadata):
    """Slide 11: Key findings summary."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    _add_section_title(slide, Inches(0.5), Inches(0.25), Inches(12.33),
                       "二、关键发现与风险提示", size=Pt(16))

    # Statistics
    total = len(all_contracts)
    unconditional_count = 0
    conditional_count = 0
    term_40yr = 0
    term_short = 0

    for contract in all_contracts:
        rc = contract.get("renewal_conditions", {})
        ct = contract.get("contract_term", {})
        renewal_text = (rc.get("renewal_type", "") + rc.get("source_text_raw", ""))
        if "无条件" in renewal_text:
            unconditional_count += 1
        elif rc.get("found"):
            conditional_count += 1

        actual = ct.get("term_years")
        if actual and actual >= 20:
            term_40yr += 1
        elif actual and actual <= 5:
            term_short += 1

    # Left panel: statistics
    left_x = Inches(0.5)
    left_w = Inches(5.8)
    row_y = Inches(0.8)

    # Stat card
    _add_section_title(slide, left_x, row_y, left_w, "合同期限分布")
    row_y += Inches(0.4)
    tb_stat = _add_textbox(slide, left_x, row_y, left_w, Inches(1.5))
    tf = tb_stat.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    _run(p0, "", size=Pt(2))
    _add_field(tf, "40年实际期限（20+20自动续）", f"{term_40yr}/{total} 人", value_bold=True,
               value_size=Pt(14), value_color=BLUE_DARK)
    _add_field(tf, "5年及以下短期", f"{term_short}/{total} 人", value_size=Pt(11))
    _add_field(tf, "附条件延期", f"{conditional_count}/{total} 人（良田：收入>1000万→延2年）",
               value_size=Pt(9), italic=True)

    row_y += Inches(1.7)

    _add_section_title(slide, left_x, row_y, left_w, "续约条件分布")
    row_y += Inches(0.4)
    tb_renewal = _add_textbox(slide, left_x, row_y, left_w, Inches(1.5))
    tf2 = tb_renewal.text_frame
    tf2.word_wrap = True
    p0 = tf2.paragraphs[0]
    _run(p0, "", size=Pt(2))
    _add_field(tf2, "无条件自动续约", f"{unconditional_count}/{total} 人", value_bold=True,
               value_size=Pt(14), value_color=GREEN_DARK)
    _add_field(tf2, "附条件/协商续约", f"{conditional_count}/{total} 人（良田、小鱼海棠）",
               value_size=Pt(9), italic=True)
    _add_field(tf2, "续约违约金范围", "合作总收益 50% ~ 150%", value_size=Pt(9),
               value_color=RED_DARK, value_bold=True)

    # Right panel: highlights
    right_x = Inches(6.6)
    right_w = Inches(6.4)

    _add_section_title(slide, right_x, Inches(0.8), right_w, "收益分成模式总结")
    tb_profit = _add_textbox(slide, right_x, Inches(1.2), right_w, Inches(5.0))
    tf3 = tb_profit.text_frame
    tf3.word_wrap = True
    p0 = tf3.paragraphs[0]
    _run(p0, "", size=Pt(2))

    findings = [
        ("🔴 核心模式", "全部采用「净利收益模式」：收入先扣显性成本，以净利润为基数分成"),
        ("📊 分成梯度", "普遍设置 2~3 级时间梯度，乙方（达人）比例逐年提升\n  · 典型路径：甲60:乙40 → 甲50:乙50 → 甲40:乙60"),
        ("⏱ 结算周期", "全部月度结算，每月8日前出结算单"),
        ("⚠ 极端案例", "小鱼海棠（2021年签）：广告/电商甲90%:乙10%，乙方比例最低"),
        ("🔒 账号所有权", "所有合同均约定：账号及内容知识产权完全归甲方（公司）所有"),
    ]
    for label, detail in findings:
        _add_field(tf3, label, detail, value_size=Pt(9))

    # Risk callout box
    risk_y = Inches(5.6)
    risk_box = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), risk_y, Inches(12.33), Inches(1.3)
    )
    risk_box.fill.solid()
    risk_box.fill.fore_color.rgb = RGBColor(0xFF, 0xF2, 0xCC)
    risk_box.line.color.rgb = RGBColor(0xE0, 0xC0, 0x40)
    risk_box.line.width = Pt(1)

    tb_risk = _add_textbox(slide, Inches(0.7), risk_y + Inches(0.1),
                           Inches(11.9), Inches(1.1))
    tf_risk = tb_risk.text_frame
    tf_risk.word_wrap = True
    p_risk = tf_risk.paragraphs[0]
    _run(p_risk, "⚡ 法律风险提示", size=Pt(13), bold=True, color=RED_DARK)
    risks = [
        "① 20+20年「无条件自动延续」条款可能构成《民法典》第497条格式条款无效情形（不合理地限制对方主要权利）",
        "② 账号及知识产权全部归甲方，乙方解约后「净身出户」——人力资本与财产性权益严重不对等",
        "③ 续约优先权+天文违约金（合作总收益100%-150%）形成事实上的「终身绑定」，可能违反《民法典》公平原则",
    ]
    for r in risks:
        _add_paragraph(tf_risk, r, size=Pt(8), color=RGBColor(0x80, 0x40, 0x00),
                       space_before=Pt(2))


def build_end(prs):
    """Slide 12: End / Thank you."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    bg = slide.background
    bg.fill.solid()
    bg.fill.fore_color.rgb = BLUE_DARK

    tb = _add_textbox(slide, Inches(2), Inches(2.5), Inches(9), Inches(1.5))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _run(p, "谢谢！", size=Pt(44), bold=True, color=WHITE)

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(20)
    _run(p2, "欢迎提问与讨论", size=Pt(20), color=RGBColor(0xDA, 0xEE, 0xF3))

    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(4), Inches(4.5), Inches(5), Pt(3)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = WHITE
    line.line.fill.background()


# ── 猴哥 contract data ─────────────────────────────────────────────────────
HOUGE_CONTRACT = {
    "contract_name": "猴哥说车-侯焜山-南京车节奏-2024年签",
    "contract_term": {
        "found": True,
        "source_text_raw": "本协议有效期及双方合作期限为二十年，自2024年2月1日起，至2044年1月31日止。协议到期后，无条件自动延续二十年至2064年1月31日。",
        "start_date": "2024年2月1日",
        "end_date": "2044年1月31日（无条件自动延续至2064年1月31日）",
        "term_years": 20,
        "term_type": "固定期限（实际40年：20+20无条件自动延续）",
        "early_termination": "甲方无故未付超3个月→乙方书面解约；连续3次评估不达标→甲方无责解约；身体/心理疾病或信誉降低→甲方书面解约；乙方私下承接超3次→甲方解约",
        "confidence": "high"
    },
    "renewal_conditions": {
        "found": True,
        "source_text_raw": "协议到期后，无条件自动延续二十年至2064年1月31日。乙方承诺合同期限届满后优先与甲方续约，否则按第八条第3款承担违约责任。",
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
        "source_text_raw": "附件1：2024.2.1-2025.2.28 甲50%:乙50%；2025.3.1起 甲40%:乙60%。净利收益=甲方实际到账收入-账号经营所有显性成本。月度结算，每月8日前出结算单。甲方纳税后按比例支付，乙方个税代扣代缴。",
        "base_split_ratio": "甲50%:乙50%（第1年）→ 甲40%:乙60%（第2年起）",
        "split_adjustments": [
            {"income_type": "全部商业合作（2024.2.1-2025.2.28）", "ratio": "甲50%:乙50%"},
            {"income_type": "全部商业合作（2025.3.1起）", "ratio": "甲40%:乙60%"}
        ],
        "cost_deduction_policy": "净利收益模式：甲方实际到账收入 - 账号经营所有显性成本 = 净利润基数",
        "settlement_cycle": "月度",
        "settlement_deadline_days": 8,
        "tax_handling": "甲方按法律法规纳税→税后按附件比例支付乙方；乙方个税自行承担，甲方代扣代缴",
        "minimum_guarantee": "无保底条款（甲方未收到客户款→不对乙方承担结算义务）",
        "appendix_reference": "附件1（含收益分配比例表、合作账号：抖音猴哥说车3679.2万粉、小红书36.2万粉）",
        "appendix_found_in_docx": True,
        "confidence": "high"
    }
}

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


def generate_ppt(existing_json: str, output_path: str):
    """Generate the 12-slide PPT."""
    # Load existing 7 contracts
    with open(existing_json, "r", encoding="utf-8") as f:
        existing_data = json.load(f)

    if isinstance(existing_data, list):
        existing_contracts = existing_data
    elif isinstance(existing_data, dict):
        existing_contracts = existing_data.get("contracts", list(existing_data.values()))
    else:
        raise ValueError(f"Unexpected data format: {type(existing_data)}")

    # Combine: 7 + 1 = 8 contracts
    all_contracts = list(existing_contracts) + [HOUGE_CONTRACT]

    # Match ordering: TALENT_META order
    # The JSON order: 大黄, 良田, 阿飞, 小鱼海棠, 小鱼, 刘冠宇, 顾猛 (+ 猴哥)
    if len(all_contracts) != len(TALENT_META):
        print(f"WARNING: {len(all_contracts)} contracts vs {len(TALENT_META)} metadata")

    # ── Create presentation (16:9 widescreen) ──
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Slide 1: Cover
    print("  Slide 1: 封面")
    build_cover(prs)

    # Slide 2: Overview table
    print("  Slide 2: 综合对比总览")
    build_overview_table(prs, all_contracts, TALENT_META)

    # Slides 3-10: Per talent detail
    for i, (contract, meta) in enumerate(zip(all_contracts, TALENT_META)):
        name = meta.get("talent_name", contract.get("contract_name", "?"))
        print(f"  Slide {i+3}: {name}")
        build_talent_detail(prs, contract, meta)

    # Slide 11: Summary
    print("  Slide 11: 关键发现与风险提示")
    build_summary(prs, all_contracts, TALENT_META)

    # Slide 12: End
    print("  Slide 12: 结束页")
    build_end(prs)

    # Save
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    print(f"\nPPT saved to: {output_path}")
    print(f"  Total slides: {len(prs.slides)}")


if __name__ == "__main__":
    import sys

    input_json = sys.argv[1] if len(sys.argv) > 1 else \
        os.path.join(config.BASE, "_brokerage_extract", "_combined_results.json")
    output_path = sys.argv[2] if len(sys.argv) > 2 else \
        os.path.join(config.BASE, "brokerage_excels", "经纪约关键条款分析_达人视图_20260710.pptx")

    generate_ppt(input_json, output_path)
