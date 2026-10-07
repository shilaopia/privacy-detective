import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import config

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "DeepSeek隐私政策审查"

# 颜色
GREEN  = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFEB9C")
RED    = PatternFill("solid", fgColor="FFC7CE")
GRAY   = PatternFill("solid", fgColor="D9D9D9")
HEADER = PatternFill("solid", fgColor="2F5496")
SUBHDR = PatternFill("solid", fgColor="BDD7EE")
BLUE1  = PatternFill("solid", fgColor="DEEAF1")
BLUE2  = PatternFill("solid", fgColor="E2EFDA")

thin   = Side(style="thin", color="AAAAAA")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

def sc(ws, row, col, value, fill=None, font=None, align=None):
    c = ws.cell(row=row, column=col, value=value)
    if fill:  c.fill  = fill
    if font:  c.font  = font
    if align: c.alignment = align
    c.border = border
    return c

# ── 标题行 ──────────────────────────────────────────
ws.merge_cells("A1:H1")
c = ws["A1"]
c.value = "DeepSeek 隐私政策合规审查报告（基于生成式AI平台实证研究框架）"
c.fill = HEADER
c.font = Font(bold=True, color="FFFFFF", size=13)
c.alignment = center
c.border = border

ws.merge_cells("A2:H2")
c = ws["A2"]
c.value = "平台：DeepSeek（杭州深度求索）  |  政策版本：2026-02-10  |  审查框架：《个保法》《生成式AI服务管理暂行办法》"
c.fill = SUBHDR
c.font = Font(size=9, italic=True, color="1F3864")
c.alignment = center
c.border = border

# ── 列标题 ──────────────────────────────────────────
headers = ["阶段", "审查维度", "具体审查项", "原文摘录（关键句）", "出处", "合规判断", "说明", "风险等级"]
for i, h in enumerate(headers, 1):
    sc(ws, 3, i, h,
       fill=PatternFill("solid", fgColor="1F4E79"),
       font=Font(bold=True, color="FFFFFF", size=10),
       align=center)

# ── 数据 ──────────────────────────────────────────
rows = [
    # 训练阶段
    ["模型训练阶段", "收集处理与授权", "是否基于必需目的",
     "在全安全加密技术处理和去标识化前提下，我们可能会将服务所收集的输入及对应输出，用于DeepSeek模型训练和服务的优化。",
     "第一章·第2节（智能对话）", "⚠️ 存疑",
     "训练优化不属于提供对话服务的必需目的；默认开启，用户须主动关闭，构成默认捆绑授权", "中"],

    ["模型训练阶段", "收集处理与授权", "是否捆绑授权",
     "（默认开启训练数据收集，用户需主动关闭[数据用于优化体验]）",
     "第一章·第2节", "⚠️ 存疑",
     "默认opt-in而非opt-out，与最小必要原则存在张力", "中"],

    ["模型训练阶段", "去标识化/匿名化", "是否承诺去标识化",
     "在全安全加密技术处理和去标识化前提下，我们可能会将服务所收集的输入及对应输出……",
     "第一章·第2节", "⚠️ 存疑",
     "承诺去标识化✅，但\"可能会\"表述模糊，去标识化是前提条件还是尽力义务不清晰", "低"],

    ["模型训练阶段", "用户opt-out权利", "是否提供opt-out选项",
     "如您拒绝将您的数据用于模型训练，可以在产品内通过关闭\"数据用于优化体验\"来选择退出。",
     "第一章·第2节；第六章·第2.1节", "⚠️ 存疑",
     "有opt-out路径且操作明确✅，但未承诺删除opt-out前已收集并用于训练的数据❌", "高"],

    ["模型训练阶段", "用户opt-out权利", "opt-out后是否删除已有数据",
     "关闭后您的输入和输出将不会再用于我们的模型训练。（未提及已有数据处理）",
     "第一章·第2节", "❌ 不合规",
     "仅说明关闭后不再使用，对历史数据去向完全未作说明", "高"],

    # 部署阶段
    ["模型部署阶段", "收集处理与授权", "是否基于必需目的",
     "除非是为实现业务功能或根据法律法规要求所必需的必要信息，您均可以拒绝提供且不影响其他功能或服务。",
     "第一章·引言", "✅ 合规",
     "声明非必要信息可拒绝，不影响其他功能", "低"],

    ["模型部署阶段", "收集处理与授权", "是否单独同意（敏感信息）",
     "敏感权限不会默认开启，只有经过您的明示授权才会在为实现特定功能或服务时使用，您也可以撤回授权。",
     "第一章·第2节", "✅ 合规",
     "相机/相册/麦克风需单独明示授权，拒绝后不影响其他功能", "低"],

    ["模型部署阶段", "收集处理与授权", "剪贴板信息收集合法性",
     "当您选择从系统剪贴板粘贴内容至智能对话输入框时，我们会采集您的剪贴板信息。",
     "第一章·第2节", "⚠️ 存疑",
     "收集剪贴板信息的合法性基础未单独说明，敏感性较高", "中"],

    ["模型部署阶段", "存储与安全", "存储地点",
     "将在境内运营过程中收集和产生的您的个人信息存储于中华人民共和国境内。目前，我们不会将上述信息传输至境外。",
     "第五章·第1节", "✅ 合规",
     "明确境内存储，暂不跨境", "低"],

    ["模型部署阶段", "存储与安全", "存储期限",
     "我们仅在为提供DeepSeek之目的所必需的期限内保留……手机号码：持续保留；对话记录：保留以展示对话历史。",
     "第五章·第2节", "⚠️ 存疑",
     "手机号\"持续保留\"无明确上限；对话记录未设终止期限；整体期限表述笼统", "中"],

    ["模型部署阶段", "存储与安全", "加密措施",
     "我们会对相关信息采用专业加密存储与传输方式，保障用户个人信息的安全。",
     "第四章·第1节", "✅ 合规",
     "承诺加密存储与传输", "低"],

    ["模型部署阶段", "存储与安全", "安全事件通知",
     "及时向您告知：安全事件发生的原因……以电子邮件、短信、电话、发送通知等方式告知您。",
     "第四章·第2节", "✅ 合规",
     "安全事件通知机制完善，多渠道告知", "低"],

    ["模型部署阶段", "共享与转移", "第三方SDK是否逐一披露",
     "我们会对合作方获取信息的SDK、API进行严格的安全监控……（另附第三方信息共享清单）",
     "第三章·第1节", "⚠️ 存疑",
     "正文未逐一披露SDK名称及数据用途，仅以外链清单替代，削弱告知同意效果", "中"],

    ["模型部署阶段", "共享与转移", "API数据流向说明",
     "（全文未单独说明API调用场景下用户数据的具体流向及责任归属）",
     "第三章", "❌ 不合规",
     "API中转场景下数据去向、责任主体（通知/删除/跨境由谁负责）完全缺失", "高"],

    ["模型部署阶段", "共享与转移", "跨境传输",
     "如果我们向境外传输，会严格遵守中国的相关法律、监管政策，并会遵循相关国家规定或者征求您的同意。",
     "第五章·第1节", "✅ 合规",
     "原则性合规，实际暂无跨境传输", "低"],

    ["模型部署阶段", "个性化决策", "是否说明算法推荐逻辑",
     "（全文未提及个性化推荐算法的存在及运作逻辑）",
     "—", "❌ 不合规",
     "全文未披露是否存在个性化推荐，违反《个保法》第24条自动化决策告知义务", "高"],

    ["模型部署阶段", "个性化决策", "是否提供关闭个性化推荐选项",
     "（全文未提供关闭个性化推荐的入口）",
     "—", "❌ 不合规",
     "无关闭选项，无删除画像标签功能", "高"],

    ["模型部署阶段", "用户权利", "opt-out（拒绝训练）",
     "可以在产品内通过关闭\"数据用于优化体验\"来选择退出。",
     "第一章·第2节", "⚠️ 存疑",
     "有路径但缺历史数据删除承诺", "高"],

    ["模型部署阶段", "用户权利", "撤回同意",
     "您可以在设备本身的操作系统中，关闭相机、相册、麦克风权限，改变同意范围或撤回您的授权。",
     "第六章·第2.1节", "✅ 合规",
     "可通过系统设置撤回敏感权限", "低"],

    ["模型部署阶段", "用户权利", "查阅/复制/更正/删除",
     "您可以通过查阅/更正/补充账号信息、查阅/删除历史对话等方式行使权利。",
     "第六章·第2.2节", "✅ 合规",
     "提供操作路径，覆盖主要权利", "低"],

    ["模型部署阶段", "用户权利", "注销账户",
     "您有权根据《DeepSeek注销须知》注销您的账号……并在十五个工作日内回复您的请求。",
     "第六章·第4节", "✅ 合规",
     "15个工作日处理，注销后删除或匿名化", "低"],
]

fill_map = {
    "✅ 合规":  GREEN,
    "⚠️ 存疑":  YELLOW,
    "❌ 不合规": RED,
    "— 未提及": GRAY,
}
risk_fill = {"高": RED, "中": YELLOW, "低": GREEN}

for i, row in enumerate(rows, 4):
    stage   = row[0]
    verdict = row[5]
    risk    = row[7]
    for j, val in enumerate(row, 1):
        f = None
        if j == 6: f = fill_map.get(verdict)
        elif j == 8: f = risk_fill.get(risk)
        elif j == 1: f = BLUE1 if "训练" in stage else BLUE2
        sc(ws, i, j, val,
           fill=f,
           font=Font(size=9),
           align=left if j > 2 else center)

# ── Sheet2：风险汇总 ──────────────────────────────
ws2 = wb.create_sheet("风险汇总")
ws2.merge_cells("A1:D1")
c = ws2["A1"]
c.value = "主要合规风险汇总（按风险程度排序）"
c.fill = HEADER
c.font = Font(bold=True, color="FFFFFF", size=12)
c.alignment = center
c.border = border

for i, h in enumerate(["优先级", "风险描述", "涉及条款位置", "整改建议"], 1):
    sc(ws2, 2, i, h,
       fill=PatternFill("solid", fgColor="1F4E79"),
       font=Font(bold=True, color="FFFFFF", size=10),
       align=center)

risks = [
    ["风险1（高）",
     "个性化决策完全缺失披露：未说明是否存在算法推荐，无关闭入口，违反《个保法》第24条",
     "第一章（全文未提及）",
     "应在隐私政策中明确说明算法推荐逻辑，并提供关闭入口及删除画像功能"],
    ["风险2（高）",
     "opt-out后训练历史数据去向不明：用户关闭训练使用后，已被用于训练的历史数据无任何处置说明",
     "第一章·第2节",
     "应补充说明opt-out前数据的删除或匿名化承诺，明确时限"],
    ["风险3（高）",
     "API数据链条责任主体缺失：通过API流转的用户数据由谁处理、谁负责通知/删除/跨境完全未说明",
     "第三章",
     "应单独披露API调用场景下的数据处理规则及责任分配"],
    ["风险4（中）",
     "第三方SDK未在正文逐一披露：仅以外链清单替代，实质削弱告知同意效果",
     "第三章·第1节",
     "应在政策正文中列明主要SDK名称、收集数据类型及用途"],
    ["风险5（中）",
     "对话记录及手机号存储期限无上限：\"持续保留\"与最小必要原则冲突",
     "第五章·第2节",
     "应设定明确的存储期限上限，或说明超期自动删除机制"],
]

for i, row in enumerate(risks, 3):
    r_fill = RED if "高" in row[0] else YELLOW
    for j, val in enumerate(row, 1):
        sc(ws2, i, j, val,
           fill=r_fill if j == 1 else None,
           font=Font(size=9, bold=(j == 1)),
           align=left)

# ── 列宽 / 行高 ──────────────────────────────────
for i, w in enumerate([14, 16, 22, 44, 20, 12, 38, 10], 1):
    ws.column_dimensions[get_column_letter(i)].width = w
for i, w in enumerate([14, 46, 22, 46], 1):
    ws2.column_dimensions[get_column_letter(i)].width = w

for row in ws.iter_rows():
    ws.row_dimensions[row[0].row].height = 40
ws.row_dimensions[1].height = 28
ws.row_dimensions[2].height = 18
ws.row_dimensions[3].height = 22

for row in ws2.iter_rows():
    ws2.row_dimensions[row[0].row].height = 52
ws2.row_dimensions[1].height = 24
ws2.row_dimensions[2].height = 20

ws.freeze_panes  = "A4"
ws2.freeze_panes = "A3"

out = os.path.join(config.BASE, "DeepSeek隐私政策审查报告.xlsx")
wb.save(out)
print("saved:", out)
