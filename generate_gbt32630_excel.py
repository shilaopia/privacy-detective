"""
GB/T 32630-2016《非结构化数据管理系统技术要求》条款提炼脚本
按12维度框架映射个人信息保护相关条款，输出10列Excel
"""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import os
import config

# ============================================================
# 数据：每条 [A,B,C,D,E,F,G,H,I,J]
# A=条款编号 B=章节标题 C=审查维度 D=适用阶段 E=要求类型
# F=原文摘录 G=核心要点(≤50字) H=适用主体 I=表述问题 J=显著标识要求
# ============================================================
data = [
    # ===== 第4章 术语和定义 =====
    ["4.1", "术语和定义",
     "维度1：收集处理与授权",
     "通用",
     "可选性",
     "非结构化数据 unstructured data：没有明确结构约束的数据，如文本、图像、音频、视频等。",
     "界定非结构化数据外延，含文本/图像/音频/视频，为个人信息载体提供概念基础",
     "个人信息控制者/处理者",
     "定义未区分「非结构化数据」与「个人信息」的关系，需结合GB/T 35273界定",
     ""],

    ["4.2", "术语和定义",
     "维度5：存储与安全",
     "通用",
     "可选性",
     "非结构化数据管理系统 unstructured data management system：对非结构化数据进行管理、操作的大型基础软件，提供非结构化数据存储、特征抽取、索引、查询等管理功能。",
     "定义UDMS为管理非结构化数据的基础软件，存储与索引为核心功能",
     "系统提供者",
     "",
     ""],

    # ===== 第6章 功能性要求 - 访问接口 =====
    ["6.7.1 a)", "访问接口（基本要求）",
     "维度9：API与委托处理",
     "部署阶段",
     "强制性",
     "应依从GB/T 32908-2016中第4章规定的查询语言访问接口要求。",
     "API查询语言接口须依从GB/T 32908-2016第4章标准，实现标准化访问控制",
     "个人信息处理者",
     "",
     ""],

    ["6.7.1 b)", "访问接口（基本要求）",
     "维度9：API与委托处理",
     "部署阶段",
     "强制性",
     "应依从GB/T 32908-2016中第5章规定的应用程序访问接口要求。",
     "应用程度API接口须依从GB/T 32908-2016第5章，确保接口安全合规",
     "个人信息处理者",
     "",
     ""],

    ["6.7.2", "访问接口（扩展要求）",
     "维度9：API与委托处理",
     "部署阶段",
     "强制性",
     "应依从GB/T 32908-2016中第6章规定的Web服务访问接口要求。",
     "Web服务API接口须依从GB/T 32908-2016第6章，适用于开放API场景",
     "个人信息处理者",
     "",
     ""],

    # ===== 7.1 信息安全性 - 用户管理 =====
    ["7.1.1 a)", "信息安全性（基本要求）",
     "维度4：收集处理与授权",
     "部署阶段",
     "强制性",
     "应支持创建、删除用户。",
     "系统须具备用户账号生命周期管理功能（创建与删除）",
     "个人信息控制者",
     "2016年标准未区分「业务功能所必需」与「可选功能」的同意层次",
     ""],

    ["7.1.1 b)", "信息安全性（基本要求）",
     "维度5：存储与安全",
     "部署阶段",
     "强制性",
     "应支持用户设置密码。",
     "系统须支持用户设置密码，提供基础身份认证安全机制",
     "个人信息控制者",
     "",
     ""],

    ["7.1.1 c)", "信息安全性（基本要求）",
     "维度4：收集处理与授权",
     "部署阶段",
     "强制性",
     "应支持创建、删除角色。",
     "系统须支持基于角色的访问控制(RBAC)，实现角色生命周期管理",
     "个人信息控制者",
     "",
     ""],

    ["7.1.1 d)", "信息安全性（基本要求）",
     "维度4：收集处理与授权",
     "部署阶段",
     "强制性",
     "应支持用户角色的授予、收回。",
     "系统须支持角色分配与撤销，实现最小权限原则的基础设施",
     "个人信息控制者",
     "",
     ""],

    ["7.1.1 e)", "信息安全性（基本要求）",
     "维度4：收集处理与授权",
     "部署阶段",
     "强制性",
     "应支持用户和角色权限的授予、查看、收回。",
     "须建立完整权限体系，支持用户级与角色级权限的授予/查看/收回",
     "个人信息控制者",
     "",
     ""],

    # ===== 7.1 信息安全性 - 数据安全 =====
    ["7.1.1 f)", "信息安全性（基本要求）",
     "维度5：存储与安全",
     "部署阶段",
     "强制性",
     "应支持数据加密。",
     "非结构化数据须具备加密保护能力，覆盖存储与传输场景",
     "个人信息控制者/处理者",
     "未明确加密算法标准或密钥管理要求，与GB/T 35273-2020第6.3条衔接需补充",
     ""],

    # ===== 7.1 信息安全性 - 审计 =====
    ["7.1.1 g)", "信息安全性（基本要求）",
     "维度11：内设监督机构",
     "部署阶段",
     "强制性",
     "应支持用户审计。",
     "系统须支持用户操作行为审计，实现操作可追溯性",
     "个人信息控制者",
     "审计要求为基本级系统功能，非独立的个人信息保护审计机制；与《个保法》第54条合规审计效力存在层级差异",
     ""],

    # ===== 7.1 信息安全性 - 日志 =====
    ["7.1.1 h)", "信息安全性（基本要求）",
     "维度5：存储与安全",
     "部署阶段",
     "强制性",
     "应支持日志机制。",
     "系统须具备日志记录功能，用于运行状态、操作行为及异常的记录追溯",
     "个人信息控制者/处理者",
     "",
     ""],

    # ===== 7.1 信息安全性 - 扩展要求 =====
    ["7.1.2 a)", "信息安全性（扩展要求）",
     "维度5：存储与安全",
     "部署阶段",
     "强制性",
     "应支持数据多副本。",
     "高安全等级场景须支持非结构化数据多副本存储，提升数据可靠性与灾备能力",
     "个人信息处理者",
     "",
     ""],

    ["7.1.2 b)", "信息安全性（扩展要求）",
     "维度5：存储与安全",
     "部署阶段",
     "强制性",
     "应支持存储实例的备份与恢复。",
     "扩展级系统须支持存储实例备份与恢复，满足数据可恢复性要求",
     "个人信息处理者",
     "",
     ""],

    # ===== 7.2 易用性 =====
    ["7.2 基本", "易用性（基本要求）",
     "维度12：格式与显著标识",
     "部署阶段",
     "强制性",
     "应提供完整的用户手册、联机帮助、图形化管理界面、模型定义和数据操作的交互工具。",
     "系统须提供完整用户手册与交互工具，保障用户操作知情与可控",
     "个人信息控制者",
     "未明确要求用户手册中须以显著方式呈现安全/隐私条款",
     ""],

    # ===== 7.3 维护性 =====
    ["7.3 基本", "维护性（基本要求）",
     "维度5：存储与安全",
     "部署阶段",
     "强制性",
     "应支持日志机制。",
     "维护性维度要求系统具备日志机制，补充信息安全维度的日志要求",
     "个人信息处理者",
     "7.1.1 h)与7.3均含日志机制要求，表述重复但侧重不同（安全vs运维）",
     ""],

    ["7.3 基本", "维护性（基本要求）",
     "维度5：存储与安全",
     "部署阶段",
     "强制性",
     "应支持存储实例的备份与恢复。",
     "维护性维度要求备份恢复能力，与信息安全扩展要求形成双重覆盖",
     "个人信息处理者",
     "",
     ""],

    ["7.3 基本", "维护性（基本要求）",
     "维度5：存储与安全",
     "部署阶段",
     "强制性",
     "应提供故障恢复工具。",
     "系统须提供故障恢复工具，保障数据可用性与服务连续性",
     "个人信息处理者",
     "",
     ""],

    ["7.3 基本", "维护性（基本要求）",
     "维度5：存储与安全",
     "部署阶段",
     "强制性",
     "应支持系统模块的热插拔。",
     "系统模块须支持热插拔，确保安全更新不中断服务",
     "个人信息处理者",
     "",
     ""],

    # ===== 7.4 兼容性 =====
    ["7.4 基本", "兼容性（基本要求）",
     "维度6：共享与转移",
     "部署阶段",
     "强制性",
     "应支持GB 18030-2005的强制部分。",
     "非结构化数据字符编码须符合GB 18030-2005强制性条款，保障数据跨系统兼容",
     "个人信息处理者",
     "",
     ""],

    ["7.4 基本", "兼容性（基本要求）",
     "维度6：共享与转移",
     "部署阶段",
     "强制性",
     "应支持多种操作系统运行环境；应支持C++或Java主流编程语言。",
     "系统须多平台兼容且支持主流编程语言，保障数据可移植性与跨系统互操作",
     "个人信息处理者",
     "",
     ""],

    # ===== 第2章 符合性（框架性条款） =====
    ["2", "符合性",
     "维度5：存储与安全",
     "通用",
     "可选性",
     "非结构化数据管理系统若满足本标准基本要求中的所有要求，则称其满足本标准的基本要求。",
     "确立基本/扩展双层符合性评价体系，基本要求为最低安全基线",
     "个人信息处理者",
     "",
     ""],

    # ===== 第3章 规范性引用文件 =====
    ["3", "规范性引用文件",
     "维度9：API与委托处理",
     "通用",
     "可选性",
     "下列文件对于本文件的应用是必不可少的：GB 18030-2005、GB/T 32908-2016。",
     "引用的GB/T 32908-2016为API访问接口安全规范的上位标准",
     "个人信息处理者",
     "",
     ""],
]

# ============================================================
# 需求类型说明行（独立sheet）
# ============================================================
notes = [
    ["说明项", "说明内容"],
    ["标准基本信息", "GB/T 32630-2016《非结构化数据管理系统技术要求》，2016-04-25发布，2016-11-01实施，推荐性国家标准"],
    ["标准定位", "本标准的「应」条款为系统技术功能层面的强制性要求，非个人信息保护法层面的法律义务"],
    ["与个保法关系", "本标准2016年发布，早于《个人信息保护法》(2021)，缺乏「同意」「单独同意」「个人信息主体权利」等个保法概念"],
    ["维度覆盖说明", "12维度中，维度2(去标识化/匿名化)、维度3(用户opt-out训练阶段)、维度7(个性化决策)、维度8(用户权利)、维度10(跨境传输)在本标准中无直接对应条款"],
    ["维度11特别说明", "7.1.1 g)用户审计为系统级审计功能要求，非《个保法》第52条DPO设置义务或第54条合规审计要求，效力层级为「应支持」技术功能(GB/T推荐性标准内部强制)vs《个保法》法定义务"],
    ["维度12特别说明", "本标准无隐私政策格式/显著标识的直接要求；7.2用户手册条款为唯一与用户信息呈现相关的条款，但未设显著标识标准"],
    ["扩展/基本层级", "本标准将要求分为「基本要求」(最低满足)和「扩展要求」(高安全/全功能)，扩展要求以基本要求满足为前提"],
    ["规范性语言", "全标准使用「应」(强制性技术规范，系统必须满足)、「宜」(推荐性技术建议，如6.1b)，「可/允许/能够」未在安全相关条款中出现"],
]

# ============================================================
# 写入 Excel
# ============================================================
wb = openpyxl.Workbook()

# --- Sheet 1: 条款明细 ---
ws = wb.active
ws.title = "条款明细"

thin = Side(style="thin", color="AAAAAA")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left_a = Alignment(horizontal="left", vertical="center", wrap_text=True)
HF = PatternFill("solid", fgColor="1F4E79")
HFont = Font(bold=True, color="FFFFFF", size=10)

headers = ["条款编号", "章节标题", "审查维度", "适用阶段", "要求类型",
           "原文摘录", "核心要点", "适用主体", "表述问题", "显著标识要求"]

for i, h in enumerate(headers, 1):
    c = ws.cell(row=1, column=i, value=h)
    c.fill = HF
    c.font = HFont
    c.alignment = center
    c.border = border

# 交替行颜色
light_fill = PatternFill("solid", fgColor="EBF5FB")

for r, row in enumerate(data, 2):
    for ci, val in enumerate(row, 1):
        cell = ws.cell(row=r, column=ci, value=val)
        cell.border = border
        cell.alignment = left_a if ci >= 6 else center
        cell.font = Font(size=9)
        if r % 2 == 0:
            cell.fill = light_fill

# 列宽
for i, w in enumerate([16, 22, 24, 14, 12, 52, 30, 18, 30, 30], 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

# 高亮"强制性"行
mandatory_fill = PatternFill("solid", fgColor="FFF2CC")
for r in range(2, len(data) + 2):
    if ws.cell(row=r, column=5).value == "强制性":
        for ci in range(1, 11):
            ws.cell(row=r, column=ci).fill = mandatory_fill

# --- Sheet 2: 汇总统计 ---
ws2 = wb.create_sheet("汇总统计")

# 标题
ws2.merge_cells("A1:J1")
title_cell = ws2.cell(row=1, column=1,
                      value="GB/T 32630-2016《非结构化数据管理系统技术要求》个人信息保护相关条款提炼汇总")
title_cell.font = Font(bold=True, size=13, color="1F4E79")
title_cell.alignment = Alignment(horizontal="center", vertical="center")

# 基本信息
info_rows = [
    ["标准号", "GB/T 32630-2016"],
    ["标准名称", "非结构化数据管理系统技术要求"],
    ["英文名称", "Technical requirements for unstructured data management system"],
    ["发布日期", "2016-04-25"],
    ["实施日期", "2016-11-01"],
    ["标准性质", "推荐性国家标准（GB/T）"],
    ["归口单位", "全国信息技术标准化技术委员会（SAC/TC 28）"],
    ["起草单位", "浙江大学、中国电子技术标准化研究院、清华大学、中国人民大学、北京航空航天大学"],
    ["提炼日期", "2026-05-19"],
    ["总提取条款数", str(len(data))],
    ["其中强制性条款", str(sum(1 for d in data if "强制性" in d[4]))],
    ["其中推荐性条款", str(sum(1 for d in data if "推荐性" in d[4]))],
    ["含表述问题条款", str(sum(1 for d in data if d[8]))],
]
for ri, (k, v) in enumerate(info_rows, 3):
    ws2.cell(row=ri, column=1, value=k).font = Font(bold=True, size=10)
    ws2.cell(row=ri, column=2, value=v).font = Font(size=10)
    for ci in range(1, 3):
        ws2.cell(row=ri, column=ci).border = border

# 维度分布表
ws2.cell(row=len(info_rows) + 4, column=1, value="维度分布统计").font = Font(bold=True, size=11)
dim_headers = ["审查维度", "条款数", "强制性", "推荐性", "含表述问题"]
for ci, h in enumerate(dim_headers, 1):
    c = ws2.cell(row=len(info_rows) + 5, column=ci, value=h)
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = HF
    c.alignment = center
    c.border = border

from collections import Counter
dim_counter = Counter(d[2] for d in data)
dim_mandatory = Counter(d[2] for d in data if "强制性" in d[4])
dim_issue = Counter(d[2] for d in data if d[8])

for di, dim_name in enumerate(sorted(dim_counter.keys()), len(info_rows) + 6):
    ws2.cell(row=di, column=1, value=dim_name).font = Font(size=9)
    ws2.cell(row=di, column=2, value=dim_counter[dim_name]).font = Font(size=9)
    ws2.cell(row=di, column=3, value=dim_mandatory[dim_name]).font = Font(size=9)
    ws2.cell(row=di, column=4, value=dim_counter[dim_name] - dim_mandatory[dim_name]).font = Font(size=9)
    ws2.cell(row=di, column=5, value=dim_issue[dim_name]).font = Font(size=9)
    for ci in range(1, 6):
        ws2.cell(row=di, column=ci).border = border
        ws2.cell(row=di, column=ci).alignment = center

# 列宽
for ci, w in enumerate([30, 12, 12, 12, 12], 1):
    ws2.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = w

# --- Sheet 3: 说明 =====
ws3 = wb.create_sheet("说明")
ws3.merge_cells("A1:B1")
ws3.cell(row=1, column=1, value="提炼说明与注意事项").font = Font(bold=True, size=12, color="1F4E79")
for ri, (k, v) in enumerate(notes, 3):
    ws3.cell(row=ri, column=1, value=k).font = Font(bold=True, size=10)
    ws3.cell(row=ri, column=2, value=v).font = Font(size=10)
    ws3.cell(row=ri, column=2).alignment = Alignment(wrap_text=True)
    for ci in range(1, 3):
        ws3.cell(row=ri, column=ci).border = border
ws3.column_dimensions["A"].width = 22
ws3.column_dimensions["B"].width = 80

# ============================================================
# 保存
# ============================================================
out = os.path.join(config.BASE, "National Standard", "归档_excels", "GBT_32630-2016_非结构化数据管理.xlsx")
wb.save(out)
print(f"Saved: {out}")
print(f"Total rows (条款明细): {len(data)}")
print(f"强制性: {sum(1 for d in data if '强制性' in d[4])}")
print(f"推荐性: {sum(1 for d in data if '推荐性' in d[4])}")
print(f"含表述问题: {sum(1 for d in data if d[8])}")
print("Done!")
