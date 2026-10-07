# -*- coding: utf-8 -*-
"""
gen_gb_batch.py — 国标条款基线批量提炼

输入目录的每个 PDF 用 gb-standard-analyzer SKILL 框架 + Claude API 提炼条款，
输出一份 Excel 到输出目录（一份标准一个 xlsx）。

用法：
    python gen_gb_batch.py                  # 全部 42 份
    python gen_gb_batch.py --limit 1        # 只跑第一个（冒烟测试）
    python gen_gb_batch.py --only GBT+34942 # 只跑文件名包含该子串的
    python gen_gb_batch.py --force          # 已存在的 xlsx 也覆盖
"""
from __future__ import annotations

import config
import argparse
import base64
import json
import os
import re
import sys
import time
import traceback
from pathlib import Path
from typing import Optional

import fitz  # PyMuPDF
import pdfplumber
from anthropic import Anthropic
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────
# 路径与常量
# ─────────────────────────────────────────────────────────────
PROJ_ROOT = Path(config.BASE)
IN_DIR    = PROJ_ROOT / "National Standard" / "re-nationalstandard" / "re_national_standard"
OUT_DIR   = PROJ_ROOT / "National Standard" / "re-nationalstandard" / "re_national_standard_excels"
SKILL_MD  = PROJ_ROOT / ".claude" / "skills" / "gb-standard-analyzer" / "SKILL.md"
TMP_DIR   = PROJ_ROOT / "_tmp_gb_render"

MODEL = "claude-sonnet-4-5"  # SDK 别名解析到当前 Sonnet
# 若 SDK 不识别，则回退使用具体 ID
MODEL_FALLBACK = "claude-sonnet-4-5-20250929"

MAX_PAGES_PER_CALL = 30      # 单次 API 调用最多附 30 页 PNG（防 token 爆）
RENDER_ZOOM        = 1.6     # PNG 渲染倍率（清晰度 vs. token 成本）
TEXT_MODE_MIN_CHARS = 2000   # 文本提取超过此字符数即用文本模式
MAX_OUTPUT_TOKENS  = 16000

# ─────────────────────────────────────────────────────────────
# Excel 样式（与 make_report.py / gen_gb_report.py 一致）
# ─────────────────────────────────────────────────────────────
GREEN  = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFEB9C")
RED    = PatternFill("solid", fgColor="FFC7CE")
GRAY   = PatternFill("solid", fgColor="D9D9D9")
HEADER = PatternFill("solid", fgColor="2F5496")
SUBHDR = PatternFill("solid", fgColor="BDD7EE")
BLUE1  = PatternFill("solid", fgColor="DEEAF1")   # 训练
BLUE2  = PatternFill("solid", fgColor="E2EFDA")   # 部署
PURPLE = PatternFill("solid", fgColor="EAE0F0")   # API/中转
ORANGE = PatternFill("solid", fgColor="FFE4B5")   # 通用
WARN   = PatternFill("solid", fgColor="FFF2CC")
COLHDR = PatternFill("solid", fgColor="1F4E79")

thin   = Side(style="thin", color="AAAAAA")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

STAGE_FILL = {
    "训练阶段": BLUE1,
    "部署阶段": BLUE2,
    "API场景":  PURPLE,
    "通用":     ORANGE,
}
TYPE_FILL = {
    "强制性": RED,
    "推荐性": YELLOW,
    "可选性": GRAY,
}

DIM_NAMES = {
    1:  "收集处理与授权",
    2:  "去标识化/匿名化",
    3:  "用户退出权利",
    4:  "收集处理与授权（敏感个人信息）",
    5:  "存储与安全",
    6:  "共享与转移",
    7:  "个性化决策",
    8:  "用户权利",
    9:  "API与委托处理",
    10: "跨境传输专项",
    11: "内设监督机构",
    12: "格式与显著标识",
}


def sc(ws, row, col, value, fill=None, font=None, align=None):
    if isinstance(value, (dict, list)):
        value = json.dumps(value, ensure_ascii=False)
    c = ws.cell(row=row, column=col, value=value)
    if fill:  c.fill = fill
    if font:  c.font = font
    if align: c.alignment = align
    c.border = border
    return c


# ─────────────────────────────────────────────────────────────
# PDF → 文本 / PNG
# ─────────────────────────────────────────────────────────────
def try_text_extraction(pdf_path: Path) -> str:
    """尝试用 pdfplumber 提文本。返回拼接后的字符串（可能为空）。"""
    try:
        out = []
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                t = page.extract_text() or ""
                if t.strip():
                    out.append(f"\n===== Page {i+1} =====\n{t}")
        return "\n".join(out)
    except Exception:
        return ""


def render_pdf_to_pngs(pdf_path: Path, dest_dir: Path, zoom: float = RENDER_ZOOM) -> list[Path]:
    """用 PyMuPDF 渲染所有页面为 PNG。返回 PNG 路径列表。"""
    dest_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    doc = fitz.open(pdf_path)
    mat = fitz.Matrix(zoom, zoom)
    for i in range(len(doc)):
        out = dest_dir / f"p{i+1:03d}.png"
        if not out.exists():
            pix = doc[i].get_pixmap(matrix=mat)
            pix.save(out)
        paths.append(out)
    doc.close()
    return paths


def png_to_b64(p: Path) -> str:
    with open(p, "rb") as f:
        return base64.standard_b64encode(f.read()).decode("ascii")


# ─────────────────────────────────────────────────────────────
# Prompt 构造
# ─────────────────────────────────────────────────────────────
def load_skill() -> str:
    return SKILL_MD.read_text(encoding="utf-8")


JSON_INSTRUCTION = """
你将按上方 SKILL 框架对一份国标/规范性文件做条款提炼。

**输出要求：** 仅输出一段 JSON（不要 markdown 代码块包围、不要任何其他文字），结构如下：

{
  "standard_id": "如 GB/T 35273-2020 或 部门规章名",
  "title": "标准/文件标题",
  "publish_date": "如 2020-03-06 或 -",
  "nature": "国家标准/行业标准/部门规章/规范性文件",
  "scan_summary": "一句话陈述本次提炼是否做了逐章遍历、覆盖了哪些主要章节",
  "clauses": [
    {
      "stage":            "训练阶段|部署阶段|API场景|通用",
      "dim":              1-12 的整数,
      "dim_name":         "维度N：xxx",
      "clause_no":        "X.X.X 或 第X条",
      "section_title":    "章节标题",
      "requirement_type": "强制性|推荐性|可选性",
      "quote":            "直接引用国标原文（保留标点）",
      "core_point":       "≤50字一句话提炼",
      "subject":          "个人信息控制者|处理者|双方|第三方|监管机构",
      "expression_issue": "如有：定义模糊/衔接矛盾/与上位法冲突；无则空字符串",
      "salient_marking":  "如有：标注该条款是否要求显著方式呈现，以及典型合规/违规表现；无则空字符串"
    }
  ],
  "baseline_summary": [
    {"dim": 1-12, "mandatory": 0, "recommended": 0, "optional": 0, "with_issue": 0, "note": "命中数为0时填写'全文扫描确认未规定'等"}
  ],
  "miss_check": {
    "chapters_scanned": ["章节1标题", "章节2标题"],
    "keyword_hits": "简述关键词回扫结果（哪些关键词命中、是否均已提炼）",
    "priority_reclass_notes": "特殊优于一般核查后做的归类调整说明（无则填'无'）"
  }
}

**关键要求：**
1. **穷尽提炼**：标准全文逐章扫描，附录与术语定义章也要看。不要因"明显条款已提炼"就停止。
2. **特殊优于一般归类**：opt-out→维度3、敏感个人信息→维度4、共享委托→维度6/9、跨境→维度10、监管→维度11、告知形式→维度12。不重复计入一般维度。
3. **空维度警示**：baseline_summary 必须包含 12 个维度（dim 1-12），命中数为 0 时在 note 字段说明。
4. **原文必须真实**：quote 字段必须是文件中能搜到的原文片段（允许省略号截断），不得编造。
5. 若该文件与个人信息保护明显无关或只是术语性文件，clauses 可为空数组，但仍需填写 standard_id/title/scan_summary。
"""


def build_messages(content_blocks: list) -> list:
    """组装 Claude API messages。"""
    user_content = content_blocks + [
        {"type": "text", "text": JSON_INSTRUCTION},
    ]
    return [{"role": "user", "content": user_content}]


# ─────────────────────────────────────────────────────────────
# Claude API 调用
# ─────────────────────────────────────────────────────────────
def call_claude(client: Anthropic, skill_md: str, content_blocks: list, file_label: str) -> dict:
    """调用 Claude，返回解析后的 JSON dict。"""
    system = [
        {
            "type": "text",
            "text": skill_md,
            "cache_control": {"type": "ephemeral"},  # 跨 42 份调用复用 SKILL prompt
        }
    ]
    messages = build_messages(content_blocks)

    last_err = None
    for attempt in range(3):
        try:
            resp = client.messages.create(
                model=MODEL,
                max_tokens=MAX_OUTPUT_TOKENS,
                system=system,
                messages=messages,
            )
            text = "".join(b.text for b in resp.content if hasattr(b, "text"))
            usage = resp.usage
            print(f"  [{file_label}] tokens: in={usage.input_tokens} "
                  f"cache_read={getattr(usage, 'cache_read_input_tokens', 0)} "
                  f"out={usage.output_tokens}")
            return parse_json_robust(text, file_label)
        except Exception as e:
            last_err = e
            wait = 2 ** attempt * 5
            print(f"  [{file_label}] attempt {attempt+1} failed: {e}; retry in {wait}s")
            time.sleep(wait)
    raise RuntimeError(f"Claude API failed after 3 attempts: {last_err}")


def parse_json_robust(text: str, file_label: str) -> dict:
    """尽量从模型回复中抽出 JSON。"""
    # 去掉可能的 markdown fence
    text = re.sub(r"^```(?:json)?\s*", "", text.strip())
    text = re.sub(r"\s*```$", "", text.strip())
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        # 兜底：找第一个 { 到最后一个 }
        s = text.find("{")
        e = text.rfind("}")
        if s >= 0 and e > s:
            try:
                return json.loads(text[s:e+1])
            except json.JSONDecodeError as ex:
                raise RuntimeError(f"[{file_label}] JSON 解析失败: {ex}\n首200字: {text[:200]}")
        raise RuntimeError(f"[{file_label}] 回复中找不到 JSON 对象\n首200字: {text[:200]}")


# ─────────────────────────────────────────────────────────────
# 文本模式 vs 视觉模式
# ─────────────────────────────────────────────────────────────
def build_content_blocks_text(pdf_text: str, filename: str) -> list:
    return [
        {
            "type": "text",
            "text": f"文件名：{filename}\n\n以下是该标准的全文（pdfplumber 提取，可能保留少量布局噪音）：\n\n{pdf_text}",
        }
    ]


def build_content_blocks_vision(png_paths: list[Path], filename: str) -> list:
    blocks = [{"type": "text", "text": f"文件名：{filename}\n\n以下是该标准的逐页图像（PyMuPDF 渲染）："}]
    for p in png_paths:
        blocks.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/png",
                "data": png_to_b64(p),
            },
        })
    return blocks


# ─────────────────────────────────────────────────────────────
# Excel 输出
# ─────────────────────────────────────────────────────────────
COLS = [
    ("适用阶段",      10),
    ("审查维度",      26),
    ("条款编号",      12),
    ("章节标题",      22),
    ("要求类型",       9),
    ("原文摘录",      52),
    ("核心要点",      36),
    ("适用主体",      14),
    ("表述问题",      30),
    ("显著标识要求",  28),
]


def _normalize_dim(item: dict) -> dict:
    """让 clause / baseline 条目兼容 dim 整数缺失（仅有 dim_name / dimension）的情况。"""
    if "dim" not in item or not isinstance(item.get("dim"), int):
        name = item.get("dim_name") or item.get("dimension") or ""
        m = re.search(r"(\d+)", str(name))
        if m:
            item["dim"] = int(m.group(1))
    return item


def write_excel(data: dict, out_path: Path, filename: str):
    # 兼容多种 JSON 输出格式
    for cl in data.get("clauses", []):
        _normalize_dim(cl)
    for b in data.get("baseline_summary", []):
        _normalize_dim(b)

    wb = Workbook()
    # ── Sheet1: 条款基线 ────────────────────────────────
    ws = wb.active
    ws.title = "条款基线"

    title = f"{data.get('standard_id', filename)} · {data.get('title', '')}"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(COLS))
    c = ws.cell(1, 1, title)
    c.fill = HEADER; c.font = Font(bold=True, color="FFFFFF", size=13)
    c.alignment = center; c.border = border

    sub = f"性质：{data.get('nature', '-')}  |  发布/施行：{data.get('publish_date', '-')}  |  文件：{filename}"
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(COLS))
    c = ws.cell(2, 1, sub)
    c.fill = SUBHDR; c.font = Font(size=9, italic=True, color="1F3864")
    c.alignment = center; c.border = border

    for i, (name, width) in enumerate(COLS, 1):
        sc(ws, 3, i, name,
           fill=COLHDR, font=Font(bold=True, color="FFFFFF", size=10), align=center)
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 18

    clauses = data.get("clauses", [])
    # 排序：维度 → 条款编号
    clauses.sort(key=lambda x: (x.get("dim", 99), str(x.get("clause_no", ""))))
    for r, cl in enumerate(clauses, start=4):
        stage = cl.get("stage", "通用")
        req_type = cl.get("requirement_type", "")
        row_fill = STAGE_FILL.get(stage)
        sc(ws, r, 1, stage, fill=row_fill, align=center)
        sc(ws, r, 2, cl.get("dim_name", DIM_NAMES.get(cl.get("dim", 0), "")), fill=row_fill, align=left)
        sc(ws, r, 3, cl.get("clause_no", ""), fill=row_fill, align=center)
        sc(ws, r, 4, cl.get("section_title", ""), fill=row_fill, align=left)
        sc(ws, r, 5, req_type, fill=TYPE_FILL.get(req_type, row_fill), align=center,
           font=Font(bold=(req_type == "强制性")))
        sc(ws, r, 6, cl.get("quote", ""), fill=row_fill, align=left)
        sc(ws, r, 7, cl.get("core_point", ""), fill=row_fill, align=left)
        sc(ws, r, 8, cl.get("subject", ""), fill=row_fill, align=center)
        sc(ws, r, 9, cl.get("expression_issue", ""), fill=WARN if cl.get("expression_issue") else row_fill, align=left)
        sc(ws, r, 10, cl.get("salient_marking", ""), fill=row_fill, align=left)
        ws.row_dimensions[r].height = 42

    ws.freeze_panes = "A4"

    # ── Sheet2: 12 维度要求基线总表 ─────────────────────
    ws2 = wb.create_sheet("基线总表")
    ws2.merge_cells("A1:G1")
    c = ws2.cell(1, 1, f"{data.get('standard_id', filename)} 12 维度命中数总表")
    c.fill = HEADER; c.font = Font(bold=True, color="FFFFFF", size=12); c.alignment = center; c.border = border

    hdr2 = ["维度", "维度名称", "强制性", "推荐性", "可选性", "含表述问题", "备注"]
    widths2 = [6, 30, 10, 10, 10, 12, 50]
    for i, (h, w) in enumerate(zip(hdr2, widths2), 1):
        sc(ws2, 2, i, h, fill=COLHDR, font=Font(bold=True, color="FFFFFF"), align=center)
        ws2.column_dimensions[get_column_letter(i)].width = w

    summary_by_dim = {}
    for b in data.get("baseline_summary", []):
        d = b.get("dim")
        if not isinstance(d, int):
            # 兼容 agent 用 dimension 字符串 / dim_name 的情况
            name = b.get("dimension") or b.get("dim_name") or ""
            m = re.search(r"(\d+)", str(name))
            if m:
                d = int(m.group(1))
        if isinstance(d, int) and 1 <= d <= 12:
            summary_by_dim[d] = b
    for d in range(1, 13):
        b = summary_by_dim.get(d, {})
        mand = b.get("mandatory", 0)
        reco = b.get("recommended", 0)
        opt  = b.get("optional", 0)
        wi   = b.get("with_issue", 0)
        note = b.get("note", "")
        empty = (mand + reco + opt == 0)
        row_fill = GRAY if empty else None
        sc(ws2, d+2, 1, d, fill=row_fill, align=center)
        sc(ws2, d+2, 2, DIM_NAMES[d], fill=row_fill, align=left)
        sc(ws2, d+2, 3, mand, fill=row_fill, align=center,
           font=Font(bold=True) if mand else None)
        sc(ws2, d+2, 4, reco, fill=row_fill, align=center)
        sc(ws2, d+2, 5, opt, fill=row_fill, align=center)
        sc(ws2, d+2, 6, wi, fill=row_fill, align=center)
        sc(ws2, d+2, 7, note, fill=row_fill, align=left)

    ws2.freeze_panes = "A3"

    # ── Sheet3: 漏检自查与元数据 ────────────────────────
    ws3 = wb.create_sheet("漏检自查")
    ws3.merge_cells("A1:B1")
    c = ws3.cell(1, 1, "漏检自查记录"); c.fill = HEADER
    c.font = Font(bold=True, color="FFFFFF", size=12); c.alignment = center; c.border = border

    mc = data.get("miss_check", {})
    ws3.column_dimensions["A"].width = 22
    ws3.column_dimensions["B"].width = 90

    rows = [
        ("standard_id",  data.get("standard_id", "")),
        ("title",        data.get("title", "")),
        ("发布/施行",     data.get("publish_date", "")),
        ("性质",         data.get("nature", "")),
        ("文件名",        filename),
        ("条款总数",      len(clauses)),
        ("scan_summary", data.get("scan_summary", "")),
        ("章节扫描清单",  " | ".join(mc.get("chapters_scanned", []) or [])),
        ("关键词回扫结果", mc.get("keyword_hits", "")),
        ("特殊优于一般核查", mc.get("priority_reclass_notes", "")),
    ]
    for i, (k, v) in enumerate(rows, start=2):
        sc(ws3, i, 1, k, fill=SUBHDR, font=Font(bold=True), align=left)
        sc(ws3, i, 2, v, align=left)
        ws3.row_dimensions[i].height = 32

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# ─────────────────────────────────────────────────────────────
# 单文件处理流程
# ─────────────────────────────────────────────────────────────
def process_one(client: Anthropic, skill_md: str, pdf_path: Path, out_dir: Path,
                force: bool = False) -> dict:
    """返回处理统计 dict。"""
    out_path = out_dir / (pdf_path.stem + ".xlsx")
    if out_path.exists() and not force:
        return {"file": pdf_path.name, "status": "skip", "reason": "already_exists",
                "out": str(out_path)}

    t0 = time.time()
    text = try_text_extraction(pdf_path)
    mode = "text" if len(text) >= TEXT_MODE_MIN_CHARS else "vision"

    if mode == "text":
        blocks = build_content_blocks_text(text, pdf_path.name)
        page_info = f"text_chars={len(text)}"
    else:
        render_dir = TMP_DIR / pdf_path.stem
        pngs = render_pdf_to_pngs(pdf_path, render_dir)
        if len(pngs) > MAX_PAGES_PER_CALL:
            pngs = pngs[:MAX_PAGES_PER_CALL]  # 防爆 token；多页 PDF 后续按需扩展
            page_info = f"vision_pages={len(pngs)} (truncated from {len(list(render_dir.glob('*.png')))})"
        else:
            page_info = f"vision_pages={len(pngs)}"
        blocks = build_content_blocks_vision(pngs, pdf_path.name)

    print(f"[{pdf_path.name}] mode={mode} {page_info}")
    data = call_claude(client, skill_md, blocks, pdf_path.name)
    write_excel(data, out_path, pdf_path.name)
    elapsed = time.time() - t0
    print(f"  ✓ {len(data.get('clauses', []))} clauses → {out_path.name}  ({elapsed:.1f}s)")
    return {
        "file": pdf_path.name,
        "status": "ok",
        "mode": mode,
        "clauses": len(data.get("clauses", [])),
        "elapsed_sec": round(elapsed, 1),
        "out": str(out_path),
    }


# ─────────────────────────────────────────────────────────────
# 入口
# ─────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0, help="只跑前 N 份（0 = 全部）")
    ap.add_argument("--only",  type=str, default="", help="文件名子串过滤")
    ap.add_argument("--force", action="store_true", help="覆盖已存在的 xlsx")
    ap.add_argument("--smallest", action="store_true", help="按文件大小升序（冒烟测试用）")
    args = ap.parse_args()

    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("ERROR: 请先 export ANTHROPIC_API_KEY=sk-...")
        sys.exit(2)

    base_url = os.environ.get("ANTHROPIC_BASE_URL") or os.environ.get("PACKY_BASE_URL")
    if base_url:
        print(f"Using base_url: {base_url}")
        client = Anthropic(base_url=base_url)
    else:
        client = Anthropic()
    skill_md = load_skill()
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    files = sorted([p for p in IN_DIR.glob("*.pdf")])
    if args.only:
        files = [f for f in files if args.only in f.name]
    if args.smallest:
        files.sort(key=lambda p: p.stat().st_size)
    if args.limit > 0:
        files = files[:args.limit]

    print(f"Total to process: {len(files)}")
    log = []
    for i, f in enumerate(files, 1):
        print(f"\n[{i}/{len(files)}] {f.name}  ({f.stat().st_size//1024} KB)")
        try:
            log.append(process_one(client, skill_md, f, OUT_DIR, force=args.force))
        except Exception as e:
            traceback.print_exc()
            log.append({"file": f.name, "status": "error", "error": str(e)[:200]})

    # 写汇总日志
    log_path = OUT_DIR / "_batch_log.json"
    log_path.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    ok = sum(1 for r in log if r.get("status") == "ok")
    print(f"\n=== Done: {ok}/{len(log)} OK; log saved → {log_path}")


if __name__ == "__main__":
    main()
