#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""审查制度标尺-隐私合规-v3.xlsx，生成中文审查意见Excel。"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict
import re, sys
import os
import config

SRC = os.path.join(config.BASE, "制度标尺-隐私合规-v3.xlsx")
DST = os.path.join(config.BASE, "制度标尺-隐私合规-v3-审查意见.xlsx")

# ── 1. 读取源数据 ──
wb = openpyxl.load_workbook(SRC)
ws = wb[wb.sheetnames[0]]

rows_data = []
current_d1 = current_d2 = current_d3 = ""
for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True), 2):
    d1 = str(row[0]).strip() if row[0] else ""
    d2 = str(row[1]).strip() if row[1] else ""
    d3 = str(row[2]).strip() if row[2] else ""
    detail = str(row[3]).strip() if row[3] else ""
    source = str(row[4]).strip() if row[4] else ""
    if d1:
        current_d1 = d1
        current_d2 = ""
        current_d3 = ""
    if d2: current_d2 = d2
    if d3: current_d3 = d3
    rows_data.append({"row": row_idx, "d1": current_d1, "d2": current_d2, "d3": current_d3, "detail": detail, "source": source})

print(f"读取 {len(rows_data)} 条义务细则")

# ── 2. 检测函数 ──
def is_fragment(text):
    """检测语句碎片"""
    text = text.strip()
    if not text: return (True, "空内容")
    text_clean = re.sub(r'^\d+\.\s*', '', text)
    # 以"需，"开头的不完整句（应为"确需XXX时，应当XXX"）
    if re.match(r'^需[，,]', text_clean):
        return (True, "以'需，'开头的不完整句——缺少完整条件状语")
    # 以右括号开头
    if text_clean.startswith(')'):
        return (True, "以右括号开头——前半句缺失")
    # 检测"应当微数据中的属性"这种明显语序错乱
    if '应当微数据' in text_clean:
        return (True, "语序错乱——'微数据中的属性'不通顺")
    # "应当开启前/新增业务功能时/设计个人生物识别信息业务功能/间接获取个人信息"——多个斜杠拼接的条件串
    if text_clean.count('/') >= 3 and len(text_clean) < 80:
        return (True, "多个斜杠拼接的条件串——不是完整句子")
    # 极短内容
    if len(text_clean) < 10:
        return (True, f"内容过短({len(text_clean)}字)——可能被截断")
    return (False, "")

def is_duplicate_within_section(items_in_section):
    """检测同一对标条款内的高度相似项"""
    dups = []
    for i in range(len(items_in_section)):
        for j in range(i+1, len(items_in_section)):
            a = re.sub(r'[，。、；：\s]', '', items_in_section[i][1])
            b = re.sub(r'[，。、；：\s]', '', items_in_section[j][1])
            if len(a) >= 20 and len(b) >= 20:
                # 计算相似度（简单的包含检测）
                if a[:20] == b[:20] or a in b or b in a:
                    dups.append((items_in_section[i][0], items_in_section[j][0], items_in_section[i][1][:60]))
    return dups

# ── 3. 逐条检测 ──
issues = []       # 审查总览
omissions = []    # 遗漏与补充建议
wording = []      # 表述问题清单
section_items = defaultdict(list)  # d3 -> list of (row, detail)

for item in rows_data:
    row = item["row"]
    detail = item["detail"]
    source = item["source"]
    d1, d2, d3 = item["d1"], item["d2"], item["d3"]
    detail_no_num = re.sub(r'^\d+\.\s*', '', detail)

    section_items[d3].append((row, detail_no_num))

    # a) 语句碎片检测
    frag, frag_reason = is_fragment(detail)
    if frag:
        wording.append({
            "row": row, "d1": d1, "d3": d3,
            "current": detail[:180],
            "type": f"语句不完整：{frag_reason}",
            "suggestion": "补全为完整的、可独立理解的合规义务陈述句"
        })

    # b) 法规引用格式：多余的"第"
    if re.search(r'第第\d+条', source):
        wording.append({
            "row": row, "d1": d1, "d3": d3,
            "current": source[:180],
            "type": "法规引用格式错误：出现'第第X条'（多余'第'字）",
            "suggestion": "删除多余的'第'，改为'第X条'"
        })

    if re.search(r'第第\d+款', source):
        wording.append({
            "row": row, "d1": d1, "d3": d3,
            "current": source[:180],
            "type": "法规引用格式错误：出现'第第X款'（多余'第'字）",
            "suggestion": "删除多余的'第'，改为'第X款'"
        })

# ── 4. 检测同节内的重复项 ──
intra_section_dup_count = 0
for d3, items in section_items.items():
    dups = is_duplicate_within_section(items)
    for r1, r2, txt in dups[:3]:  # 每节最多报3对
        # 检查是否确实值得报告（排除17-18项中相近但不同的）
        if txt not in ["需，不得超范围收集使用个人信息", "提供产品或者服务所必需，不得超范围收集使用个人信息", "提供产品或者服务，个人信息属于提供产品或者服务所必需的除外"]:
            wording.append({
                "row": r1, "d1": "", "d3": d3,
                "current": f"与行{r2}高度重复：{txt}...",
                "type": "同节内条款重复",
                "suggestion": "合并重复项，或区分适用场景使每条有独立审查价值"
            })
            intra_section_dup_count += 1

if intra_section_dup_count > 0:
    issues.append({
        "cat": "条款重复", "severity": "中", "d1": "多个",
        "desc": f"同一对标条款内发现约{intra_section_dup_count}组高度相似/重复的细则",
        "suggestion": "合并重复条款，确保每条合规细则有独立的审查维度"
    })

# ── 5. 二级维度为空 ──
empty_d2_d1s = {}
for item in rows_data:
    if not item["d2"] and item["d1"] not in empty_d2_d1s:
        # 确认：该d1下所有行的d2都为空
        all_empty = all(r["d2"] == "" for r in rows_data if r["d1"] == item["d1"])
        if all_empty:
            # 统计该d1下的d3数量
            d3s = set(r["d3"] for r in rows_data if r["d1"] == item["d1"])
            empty_d2_d1s[item["d1"]] = len(d3s)

for d1, d3_count in sorted(empty_d2_d1s.items()):
    issues.append({
        "cat": "结构缺失", "severity": "中", "d1": d1,
        "desc": f"「{d1}」下二级维度为空（含{d3_count}个对标条款直接挂在一级维度下，缺少中层分类）",
        "suggestion": f"为「{d1}」补充二级维度分类，按主题或处理环节进一步细分"
    })

# ── 6. 18条模板过度复用 ──
template_d3s = set()
for item in rows_data:
    dn = re.sub(r'^\d+\.\s*', '', item["detail"])
    for marker in ["保存期限的确定方法", "告知法律、行政法规规定应当告知的其他事项",
                    "个人信息保存期限的确定方法，并公开披露", "对外提供(例如共享或转让)个人信息",
                    "公开披露个人信息", "遵循合法、正当、必要和诚信原则",
                    "取得个人信息主体的单独同意", "采取对个人信息主体权益影响最小的方式",
                    "不得通过误导、欺诈、胁迫", "不得超范围收集使用个人信息",
                    "不得以个人信息主体不同意"]:
        if len(dn) >= 6 and dn[:6] == marker[:6]:
            template_d3s.add(item["d3"])
            break

if len(template_d3s) >= 3:
    d3_sample = "、".join(sorted(template_d3s)[:5])
    issues.append({
        "cat": "模板化重复", "severity": "高", "d1": "多个",
        "desc": f"一组18条通用合规模板在 {len(template_d3s)} 个对标条款中机械重复出现（涉及：{d3_sample}等），多数模板文本与对标条款的主题无关，只是简单复制",
        "suggestion": "为每个对标条款量身定制细化规则，删除不适用的通用模板条目，保留真正与该主题相关的合规要点"
    })

# ── 7. 汇总表述问题 ──
wording_types = defaultdict(int)
for w in wording:
    wording_types[w["type"]] += 1

for wtype, count in sorted(wording_types.items(), key=lambda x: -x[1]):
    short = wtype[:60]
    issues.append({
        "cat": "表述问题", "severity": "高" if "不完整" in wtype else "中",
        "d1": "多个",
        "desc": f"{short}：共发现 {count} 处",
        "suggestion": "详见「表述问题清单」Sheet"
    })

# ── 8. 遗漏与补充建议 ──
omissions = [
    {
        "dim": "模型安全与对齐机制",
        "d1_suggest": "自动化决策（建议新增二级维度）",
        "source": "《生成式人工智能服务管理暂行办法》第4条、第10条、第14条；《互联网信息服务深度合成管理规定》第10条",
        "reason": "生成式AI特有的安全问题——模型幻觉（hallucination）、有害/偏见内容生成、越狱攻击（jailbreak）防护、内容审核过滤机制等——在当前标尺中几乎空白。「自动化决策」目前仅覆盖推荐算法和决策透明度，缺少模型层面安全保障义务的专门审查条款。"
    },
    {
        "dim": "训练数据合规（集中整合）",
        "d1_suggest": "信息收集处理与授权（建议独立二级维度）",
        "source": "《生成式人工智能服务管理暂行办法》第7条；个保法第13-16条、第27条",
        "reason": "训练数据的合法性要求目前碎片化分布在「来源合法与间接获取合规」「训练用途退出/拒绝使用」等小节，合计不足15条细则。遗漏重点：训练数据来源合法性审查机制、数据质量保障（真实性/准确性/客观性/多样性）、知识产权合规（版权/个人信息交叉）、公开爬取数据的告知同意机制等——这些都是生成式AI独有的合规难题。"
    },
    {
        "dim": "深度合成/AI生成内容标识",
        "d1_suggest": "格式与显著标识（建议独立二级维度）",
        "source": "《互联网信息服务深度合成管理规定》第16条、第17条、第18条",
        "reason": "AI生成内容的标识义务是深度合成监管的核心制度，当前仅在「格式与显著标识」中碎片提及，缺少：(1)显式标识与隐式标识的技术要求区分；(2)元数据标注规范；(3)标识的不可篡改性要求；(4)下游传播中的标识保持义务。"
    },
    {
        "dim": "数据分类分级制度",
        "d1_suggest": "存储与安全（建议独立二级维度）",
        "source": "《数据安全法》第21条；《网络数据安全管理条例》第5条、第27条",
        "reason": "数据分类分级是所有数据处理活动的基础性制度（《数据安全法》第21条明确要求），当前仅有「数据安全保护措施」泛泛涉及，缺失：重要数据目录编制、核心数据管控措施、不同级别数据的差异化保护要求。"
    },
    {
        "dim": "个人信息保护影响评估（PIA）集中清单",
        "d1_suggest": "内设监督机构（建议独立二级维度）",
        "source": "个保法第55条、第56条；GB/T 39335-2020",
        "reason": "PIA是法定义务但当前极度碎片化——同一句话在多个对标条款中反复出现。建议在「内设监督机构」下设立专门的「个人信息保护影响评估」对标条款，集中列出需进行PIA的完整触发场景（委托处理、自动化决策、敏感个人信息、跨境提供、公开个人信息等），并附评估报告保存期限（至少3年）等要求。"
    },
    {
        "dim": "人工智能伦理审查机制",
        "d1_suggest": "内设监督机构（建议新增二级维度）",
        "source": "《生成式人工智能服务管理暂行办法》第4条、第5条；《互联网信息服务算法推荐管理规定》第7条",
        "reason": "科技伦理审查是生成式AI监管的特色要求（算法公平性、不歧视原则、科技伦理委员会等），当前标尺中几乎空白。建议补充：伦理审查委员会的设置要求、算法公平性评估、歧视性输出检测与纠正机制。"
    },
    {
        "dim": "法规来源缺口：网络安全等级保护（等保）",
        "d1_suggest": "存储与安全",
        "source": "《网络安全法》第21条；GB/T 22239-2019（等保2.0）",
        "reason": "网络安全等级保护是中国网络安全的基础制度，个人信息处理系统一般需达到等保三级。当前出处中「网络安全法」的被引条款较少且未涉及等保要求，建议补充。"
    },
    {
        "dim": "法规来源缺口：2025年网络数据安全管理条例新增条款",
        "d1_suggest": "内设监督机构 / 存储与安全",
        "source": "《网络数据安全管理条例》第23条（重要数据风险评估）、第24条（供应链安全审查）、第31条（受托方管理）",
        "reason": "2025年1月生效的《网络数据安全管理条例》新增了重要数据处理者年度风险评估、供应链安全审查等义务。当前出处中虽已引用该条例，但建议逐条核查2025年新增义务是否已全部覆盖。"
    },
    {
        "dim": "跨境数据传输的豁免与例外情形",
        "d1_suggest": "共享与转移",
        "source": "《促进和规范数据跨境流动规定》第3-6条（豁免场景清单）",
        "reason": "当前「跨境传输」模块主要覆盖安全评估、标准合同、认证等正面合规路径，但《促进和规范数据跨境流动规定》（2024）明确列举了免予安全评估的跨境场景（如合同必需、人力资源管理、紧急情况等），建议补充豁免情形的审查条款，使审查清单更完整。"
    },
]

# ── 9. 补充：归属不当检测 ──
misplaced = [
    {"item": "客户数据处理规范", "from_d1": "格式与显著标识", "to_d1": "共享与转移→委托处理协议/DPA条款",
     "reason": "客户数据处理规范涉及委托方对客户数据的访问、修改、删除等要求，属于委托处理/外包管理的范畴，不应属于「格式与显著标识」维度"},
    {"item": "行踪轨迹信息保护", "from_d1": "信息收集处理与授权", "to_d1": "存储与安全（或独立作为敏感个人信息子类）",
     "reason": "行踪轨迹属于敏感个人信息，当前仅2条细则且主要涉及去标识化展示，建议与「生物识别信息处理要求」并列，或纳入「存储与安全→去标识化」下"},
]
for mp in misplaced:
    issues.append({
        "cat": "归属不当", "severity": "中", "d1": mp["from_d1"],
        "desc": f"「{mp['item']}」当前挂在「{mp['from_d1']}」下，但更应归属到「{mp['to_d1']}」。理由：{mp['reason']}",
        "suggestion": f"将「{mp['item']}」移至「{mp['to_d1']}」"
    })

print(f"审查总览: {len(issues)} 项")
print(f"遗漏建议: {len(omissions)} 项")
print(f"表述问题: {len(wording)} 项")

# ── 10. 写入Excel ──
owb = openpyxl.Workbook()
header_font = Font(name="微软雅黑", bold=True, size=11, color="FFFFFF")
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
body_font = Font(name="微软雅黑", size=10)
wrap_align = Alignment(wrap_text=True, vertical="top")
thin_border = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))
red_fill = PatternFill(start_color="F4B4B4", end_color="F4B4B4", fill_type="solid")
yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
orange_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

def write_header(ws, headers, widths):
    for col, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"

def style_body(ws, max_row, max_col, severity_col=None):
    for r in range(2, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = body_font
            cell.alignment = wrap_align
            cell.border = thin_border
        if severity_col:
            val = str(ws.cell(row=r, column=severity_col).value or "")
            if val == "高":
                for c in range(1, max_col + 1):
                    ws.cell(row=r, column=c).fill = red_fill
            elif val == "中":
                for c in range(1, max_col + 1):
                    ws.cell(row=r, column=c).fill = yellow_fill

# ── Sheet 1: 审查总览 ──
ws1 = owb.active
ws1.title = "审查总览"
h1 = ["序号", "问题类别", "严重程度", "涉及一级维度", "问题简述", "改进建议"]
w1 = [6, 16, 10, 22, 68, 55]
write_header(ws1, h1, w1)
for i, iss in enumerate(issues, 1):
    ws1.cell(row=i+1, column=1, value=i)
    ws1.cell(row=i+1, column=2, value=iss["cat"])
    ws1.cell(row=i+1, column=3, value=iss["severity"])
    ws1.cell(row=i+1, column=4, value=iss["d1"])
    ds = iss["desc"]
    if len(ds) > 400: ds = ds[:397] + "..."
    ws1.cell(row=i+1, column=5, value=ds)
    ws1.cell(row=i+1, column=6, value=iss["suggestion"])
style_body(ws1, len(issues)+1, 6, severity_col=3)

# ── Sheet 2: 遗漏与补充建议 ──
ws2 = owb.create_sheet("遗漏与补充建议")
h2 = ["序号", "建议新增维度/条款", "建议归属一级维度", "法规依据", "补充理由（详细）"]
w2 = [6, 32, 28, 55, 70]
write_header(ws2, h2, w2)
for i, om in enumerate(omissions, 1):
    ws2.cell(row=i+1, column=1, value=i)
    ws2.cell(row=i+1, column=2, value=om["dim"])
    ws2.cell(row=i+1, column=3, value=om["d1_suggest"])
    ws2.cell(row=i+1, column=4, value=om["source"])
    ws2.cell(row=i+1, column=5, value=om["reason"])
style_body(ws2, len(omissions)+1, 5)

# ── Sheet 3: 表述问题清单 ──
ws3 = owb.create_sheet("表述问题清单")
h3 = ["序号", "Excel行号", "所属一级维度", "所属对标条款", "当前表述（截取前180字）", "问题类型", "修改建议"]
w3 = [6, 10, 22, 28, 68, 32, 50]
write_header(ws3, h3, w3)
for i, w in enumerate(wording, 1):
    ws3.cell(row=i+1, column=1, value=i)
    ws3.cell(row=i+1, column=2, value=w["row"])
    ws3.cell(row=i+1, column=3, value=w["d1"])
    ws3.cell(row=i+1, column=4, value=w["d3"])
    cur = w["current"]
    if len(cur) > 250: cur = cur[:247] + "..."
    ws3.cell(row=i+1, column=5, value=cur)
    ws3.cell(row=i+1, column=6, value=w["type"])
    ws3.cell(row=i+1, column=7, value=w["suggestion"])
style_body(ws3, len(wording)+1, 7)
# 高亮语句不完整行
for r in range(2, len(wording)+2):
    val = str(ws3.cell(row=r, column=6).value or "")
    if "不完整" in val:
        for c in range(1, 8):
            ws3.cell(row=r, column=c).fill = orange_fill

owb.save(DST)
print(f"\nDone! Output: {DST}")
print(f"  Sheet 1: {len(issues)} items")
print(f"  Sheet 2: {len(omissions)} items")
print(f"  Sheet 3: {len(wording)} items")
