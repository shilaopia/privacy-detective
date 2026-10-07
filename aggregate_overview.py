# -*- coding: utf-8 -*-
"""
aggregate_overview.py
扫描 _tmp_gb_render/ 所有 JSON，跳过 0-命中标准，生成 4-sheet 横向对比 Excel：
  Sheet1 总览           — 一行一标准
  Sheet2 维度×标准 矩阵  — 12 行 × N 列
  Sheet3 重点条款全集    — 按维度分组排序
  Sheet4 阶段×维度 交叉  — 4 行 × 12 列
"""
from __future__ import annotations

import config
import json
import re
from pathlib import Path
from collections import defaultdict

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(config.BASE)
JSON_DIR = ROOT / "_tmp_gb_render"
OUT_PATH = ROOT / "National Standard" / "re-nationalstandard" / "国标汇总_横向对比.xlsx"

# ---------- styles ----------
HEADER = PatternFill("solid", fgColor="2F5496")
SUBHDR = PatternFill("solid", fgColor="BDD7EE")
GROUP = PatternFill("solid", fgColor="FFE699")
DIM_FILL = PatternFill("solid", fgColor="DEEAF1")
ZERO_FILL = PatternFill("solid", fgColor="F2F2F2")
HEAT1 = PatternFill("solid", fgColor="EAF3FB")
HEAT2 = PatternFill("solid", fgColor="BDD7EE")
HEAT3 = PatternFill("solid", fgColor="9CC2E5")
HEAT4 = PatternFill("solid", fgColor="5B9BD5")
HEAT5 = PatternFill("solid", fgColor="2E75B6")

thin = Side(style="thin", color="AAAAAA")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
GROUP_FONT = Font(bold=True, size=11)
WHITE_FONT = Font(color="FFFFFF", bold=True)


DIM_NAMES = {
    1: "收集处理与授权",
    2: "去标识化/匿名化",
    3: "退出训练权",
    4: "敏感个人信息",
    5: "存储/加密/删除/泄露",
    6: "共享与转移",
    7: "自动化决策/推荐",
    8: "主体权利响应",
    9: "委托处理(DPA)",
    10: "数据跨境",
    11: "合规审计/PIA",
    12: "告知/显著标识/隐私政策",
}
DIM_NAMES_EXT = {0: "周边标准（IT 服务/云/外包/未归 12 维度）", **DIM_NAMES}

STAGES = ["训练阶段", "部署阶段", "API场景", "通用"]


def heat(v: int):
    if v == 0: return ZERO_FILL
    if v <= 2: return HEAT1
    if v <= 5: return HEAT2
    if v <= 10: return HEAT3
    if v <= 20: return HEAT4
    return HEAT5


def sc(ws, r, c, v, fill=None, font=None, align=None):
    cell = ws.cell(row=r, column=c, value=v)
    if fill: cell.fill = fill
    if font: cell.font = font
    cell.alignment = align or center
    cell.border = border
    return cell


def short_label(stem: str, sid: str, title: str) -> str:
    """Build a short label like 'GBT 45574-2025' or '算法推荐规定'"""
    m = re.search(r"GB[T\s/]*(\d+(?:\.\d+)?[-－—]\d{4})", sid or "")
    if m:
        return f"GB/T {m.group(1)}"
    m = re.search(r"GB[T\s/]*(\d+(?:\.\d+)?[-－—]\d{4})", stem)
    if m:
        return f"GB/T {m.group(1)}"
    # non-GB regulations: use the stem cleaned
    s = stem
    # strip suffixes like "(1)", " (1)", "(2)" etc
    s = re.sub(r"\s*\(\d+\)\s*$", "", s)
    if len(s) > 22:
        s = s[:20] + "…"
    return s


def stage_norm(s: str) -> str:
    s = (s or "").strip()
    if not s: return "通用"
    # The schema uses these exact strings, but be defensive
    for k in STAGES:
        if k in s: return k
    if "训练" in s: return "训练阶段"
    if "部署" in s or "推理" in s: return "部署阶段"
    if "API" in s.upper() or "接口" in s: return "API场景"
    return "通用"


def dim_int(v) -> int:
    try:
        return int(v)
    except Exception:
        m = re.search(r"\d+", str(v))
        return int(m.group()) if m else 0


def to_text(v) -> str:
    if v is None: return ""
    if isinstance(v, str): return v
    if isinstance(v, dict):
        return "；".join(f"{k}：{to_text(vv)}" for k, vv in v.items())
    if isinstance(v, list):
        return "；".join(to_text(x) for x in v)
    return str(v)


def req_norm(s: str) -> str:
    s = (s or "")
    if "强制" in s: return "强制性"
    if "推荐" in s: return "推荐性"
    if "可选" in s: return "可选性"
    return "推荐性"


# ---------- load ----------
records = []
for jp in sorted(JSON_DIR.glob("*.json")):
    try:
        data = json.loads(jp.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"SKIP {jp.name}: parse error {e}")
        continue
    cls = data.get("clauses", []) or []
    if len(cls) == 0:
        continue
    sid = to_text(data.get("standard_id")).strip()
    title = to_text(data.get("title")).strip()
    stem = jp.stem
    label = short_label(stem, sid, title)
    records.append({
        "stem": stem,
        "sid": sid,
        "title": title,
        "label": label,
        "publish": to_text(data.get("publish_date")).strip(),
        "nature": to_text(data.get("nature")).strip(),
        "scan": to_text(data.get("scan_summary")).strip(),
        "clauses": cls,
        "baseline": data.get("baseline_summary", []),
    })

print(f"Aggregating {len(records)} non-zero-clause standards")
total_clauses = sum(len(r["clauses"]) for r in records)
print(f"Total clauses: {total_clauses}")


# ---------- workbook ----------
wb = openpyxl.Workbook()

# ============== Sheet 1 总览 ==============
ws = wb.active
ws.title = "总览"

title = "国家标准/规章 横向对比总览（生成式AI服务链条 PI 保护视角）"
ws.merge_cells("A1:N1")
c = ws["A1"]; c.value = title; c.fill = HEADER; c.font = Font(bold=True, color="FFFFFF", size=13); c.alignment = center; c.border = border

ws.merge_cells("A2:N2")
c = ws["A2"]
c.value = f"共 {len(records)} 项标准（已跳过 0-命中样本），合计 {total_clauses} 条 PI 相关合规要求；分析框架：GB/T 35273 12 维度 + 特殊优于一般"
c.fill = SUBHDR; c.font = Font(italic=True); c.alignment = center; c.border = border

headers = ["序号", "标准简称", "标准号", "标准全称", "性质", "发布/实施", "总条款", "强制", "推荐", "可选", "命中维度数", "主要维度Top3", "scan_summary 摘要"]
for col, h in enumerate(headers, 1):
    sc(ws, 3, col, h, fill=HEADER, font=HEADER_FONT)

row = 4
for idx, r in enumerate(records, 1):
    cls = r["clauses"]
    mand = sum(1 for c0 in cls if req_norm(c0.get("requirement_type","")) == "强制性")
    rec  = sum(1 for c0 in cls if req_norm(c0.get("requirement_type","")) == "推荐性")
    opt  = sum(1 for c0 in cls if req_norm(c0.get("requirement_type","")) == "可选性")
    dim_count = defaultdict(int)
    for c0 in cls:
        dim_count[dim_int(c0.get("dim", 0))] += 1
    hit_dims = sorted([(d, n) for d, n in dim_count.items() if d in DIM_NAMES], key=lambda x: -x[1])
    top3 = "、".join(f"D{d}({n})" for d, n in hit_dims[:3])
    scan_short = r["scan"]
    if len(scan_short) > 280:
        scan_short = scan_short[:275] + "…"

    sc(ws, row, 1, idx)
    sc(ws, row, 2, r["label"], align=left)
    sc(ws, row, 3, r["sid"], align=left)
    sc(ws, row, 4, r["title"], align=left)
    sc(ws, row, 5, r["nature"], align=left)
    sc(ws, row, 6, r["publish"], align=left)
    sc(ws, row, 7, len(cls))
    sc(ws, row, 8, mand)
    sc(ws, row, 9, rec)
    sc(ws, row, 10, opt)
    sc(ws, row, 11, len(hit_dims))
    sc(ws, row, 12, top3, align=left)
    sc(ws, row, 13, scan_short, align=left)
    row += 1

widths = [5, 22, 28, 36, 22, 22, 8, 6, 6, 6, 10, 18, 60]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w
ws.freeze_panes = "A4"


# ============== Sheet 2 维度×标准 矩阵 ==============
ws2 = wb.create_sheet("维度×标准矩阵")
ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(records)+2)
c = ws2.cell(row=1, column=1, value="12 维度 × 标准矩阵（每格：总条款数 / 含强制条款数；底色按数量梯度）")
c.fill = HEADER; c.font = Font(bold=True, color="FFFFFF", size=12); c.alignment = center; c.border = border

# headers
sc(ws2, 2, 1, "维度", fill=HEADER, font=HEADER_FONT)
sc(ws2, 2, 2, "维度名称", fill=HEADER, font=HEADER_FONT)
for j, r in enumerate(records, 3):
    sc(ws2, 2, j, r["label"], fill=HEADER, font=HEADER_FONT)

# data: D1-D12 + D0 row
dim_row_order = list(range(1, 13)) + [0]
for i, d in enumerate(dim_row_order, 3):
    label = f"D{d}" if d > 0 else "D0"
    sc(ws2, i, 1, label, fill=DIM_FILL, font=GROUP_FONT)
    sc(ws2, i, 2, DIM_NAMES_EXT[d], fill=DIM_FILL, align=left, font=GROUP_FONT)
    for j, r in enumerate(records, 3):
        n = 0; m = 0
        for c0 in r["clauses"]:
            dv = dim_int(c0.get("dim", 0))
            if (d == 0 and dv not in DIM_NAMES) or (d > 0 and dv == d):
                n += 1
                if req_norm(to_text(c0.get("requirement_type",""))) == "强制性":
                    m += 1
        val = "" if n == 0 else (f"{n}/{m}" if m else f"{n}")
        cell = sc(ws2, i, j, val, fill=heat(n))
        if n >= 11:
            cell.font = WHITE_FONT

# totals row (after D0 row at row 16)
total_row = 3 + len(dim_row_order)
sc(ws2, total_row, 1, "合计", fill=HEADER, font=HEADER_FONT)
sc(ws2, total_row, 2, "本标准 PI 条款总数", fill=HEADER, font=HEADER_FONT)
for j, r in enumerate(records, 3):
    sc(ws2, total_row, j, len(r["clauses"]), fill=SUBHDR, font=GROUP_FONT)

ws2.column_dimensions['A'].width = 6
ws2.column_dimensions['B'].width = 24
for j in range(3, len(records)+3):
    ws2.column_dimensions[get_column_letter(j)].width = 14
ws2.row_dimensions[2].height = 60
ws2.freeze_panes = "C3"


# ============== Sheet 3 重点条款全集（按维度分组） ==============
ws3 = wb.create_sheet("重点条款全集")
ws3.merge_cells("A1:L1")
c = ws3["A1"]
c.value = f"PI 合规条款全集（{total_clauses} 条；按 12 维度分组，组内按标准号 + 章节号排序）"
c.fill = HEADER; c.font = Font(bold=True, color="FFFFFF", size=13); c.alignment = center; c.border = border

cols3 = ["维度", "阶段", "标准简称", "标准号", "章节号", "章节标题", "类型", "引用原文", "核心点(≤50字)", "义务主体", "表达问题", "显著标识"]
for col, h in enumerate(cols3, 1):
    sc(ws3, 2, col, h, fill=HEADER, font=HEADER_FONT)

# collect all clauses grouped by dim
by_dim = defaultdict(list)
for r in records:
    for c0 in r["clauses"]:
        d = dim_int(c0.get("dim", 0))
        if d not in DIM_NAMES:
            d = 0
        by_dim[d].append((r, c0))

# sort: 1..12, then 0 at the end
ordered_keys = [k for k in range(1, 13) if k in by_dim] + ([0] if 0 in by_dim else [])
row = 3
for d in ordered_keys:
    items = by_dim[d]
    dn = DIM_NAMES_EXT[d]
    # group header
    ws3.merge_cells(start_row=row, start_column=1, end_row=row, end_column=12)
    cell = ws3.cell(row=row, column=1, value=f"D{d}  {dn}    ── 共 {len(items)} 条")
    cell.fill = GROUP; cell.font = GROUP_FONT; cell.alignment = left; cell.border = border
    row += 1
    items.sort(key=lambda x: (x[0]["label"], x[1].get("clause_no","")))
    for r0, c0 in items:
        stage = stage_norm(c0.get("stage",""))
        sc(ws3, row, 1, f"D{d}")
        sc(ws3, row, 2, stage)
        sc(ws3, row, 3, r0["label"], align=left)
        sc(ws3, row, 4, r0["sid"], align=left)
        sc(ws3, row, 5, to_text(c0.get("clause_no","")))
        sc(ws3, row, 6, to_text(c0.get("section_title","")), align=left)
        sc(ws3, row, 7, req_norm(to_text(c0.get("requirement_type",""))))
        sc(ws3, row, 8, to_text(c0.get("quote","")), align=left)
        sc(ws3, row, 9, to_text(c0.get("core_point","")), align=left)
        sc(ws3, row, 10, to_text(c0.get("subject","")), align=left)
        sc(ws3, row, 11, to_text(c0.get("expression_issue","")), align=left)
        sc(ws3, row, 12, to_text(c0.get("salient_marking","")), align=left)
        row += 1

widths3 = [6, 12, 22, 26, 10, 26, 10, 60, 32, 22, 22, 18]
for i, w in enumerate(widths3, 1):
    ws3.column_dimensions[get_column_letter(i)].width = w
ws3.freeze_panes = "A3"


# ============== Sheet 4 阶段×维度 交叉 ==============
ws4 = wb.create_sheet("阶段×维度交叉")
ws4.merge_cells("A1:N1")
c = ws4["A1"]
c.value = "阶段 × 维度 交叉条款数（衡量哪些阶段-维度组合在国标体系中受规范覆盖最强）"
c.fill = HEADER; c.font = Font(bold=True, color="FFFFFF", size=13); c.alignment = center; c.border = border

sc(ws4, 2, 1, "阶段＼维度", fill=HEADER, font=HEADER_FONT)
for d in range(1, 13):
    sc(ws4, 2, 1 + d, f"D{d}\n{DIM_NAMES[d]}", fill=HEADER, font=HEADER_FONT)
sc(ws4, 2, 14, "合计", fill=HEADER, font=HEADER_FONT)

cross = defaultdict(lambda: defaultdict(int))
for r in records:
    for c0 in r["clauses"]:
        s = stage_norm(c0.get("stage",""))
        d = dim_int(c0.get("dim", 0))
        if d in DIM_NAMES:
            cross[s][d] += 1

for i, stg in enumerate(STAGES, 3):
    sc(ws4, i, 1, stg, fill=DIM_FILL, font=GROUP_FONT)
    total = 0
    for d in range(1, 13):
        v = cross[stg][d]; total += v
        cell = sc(ws4, i, 1 + d, v if v else "", fill=heat(v))
        if v >= 11: cell.font = WHITE_FONT
    sc(ws4, i, 14, total, fill=SUBHDR, font=GROUP_FONT)

# 维度合计行
sc(ws4, 7, 1, "维度合计", fill=HEADER, font=HEADER_FONT)
for d in range(1, 13):
    v = sum(cross[s][d] for s in STAGES)
    sc(ws4, 7, 1 + d, v, fill=SUBHDR, font=GROUP_FONT)
sc(ws4, 7, 14, total_clauses, fill=HEADER, font=HEADER_FONT)

ws4.column_dimensions['A'].width = 14
for d in range(1, 13):
    ws4.column_dimensions[get_column_letter(1 + d)].width = 14
ws4.column_dimensions['N'].width = 10
ws4.row_dimensions[2].height = 42

# save
OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
wb.save(OUT_PATH)
print(f"WROTE: {OUT_PATH}")
