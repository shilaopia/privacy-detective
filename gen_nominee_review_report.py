# -*- coding: utf-8 -*-
"""Generate legal review Excel for nominee holding agreement.
Data embedded directly. All em-dashes replaced with -- for Python 3.14 compat.
"""
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
import config

# -- Styles --
GREEN = PatternFill("solid", fgColor="C6EFCE")
YELLOW = PatternFill("solid", fgColor="FFEB9C")
RED = PatternFill("solid", fgColor="FFC7CE")
GRAY = PatternFill("solid", fgColor="D9D9D9")
BLUE_HINT = PatternFill("solid", fgColor="DAEEF3")
HEADER_FILL = PatternFill("solid", fgColor="2F5496")
COLHDR_FILL = PatternFill("solid", fgColor="1F4E79")
SECTION_FILL = PatternFill("solid", fgColor="2F5496")
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")
thin = Side(style="thin", color="AAAAAA")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
HDR_F = Font(name="Microsoft YaHei", size=14, bold=True, color="FFFFFF")
SHDR_F = Font(name="Microsoft YaHei", size=10, color="2F5496")
CHDR_F = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
DF = Font(name="Microsoft YaHei", size=10)
LF = Font(name="Microsoft YaHei", size=10, bold=True, color="1F4E79")
TF = Font(name="Microsoft YaHei", size=16, bold=True, color="2F5496")
SF = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
BF = Font(name="Microsoft YaHei", size=10, bold=True)
CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
LFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LFTT = Alignment(horizontal="left", vertical="top", wrap_text=True)

def sc(ws, r, c, v, fill=None, font=None, align=None):
    cell = ws.cell(row=r, column=c, value=v)
    if fill: cell.fill = fill
    if font: cell.font = font
    if align: cell.alignment = align
    cell.border = border
    return cell

def rfill(lvl):
    m = {"HIGH": RED, "MID": YELLOW, "LOW": GREEN}; return m.get(lvl, None)
def pfill(p):
    m = {"HIGH": RED, "MID": YELLOW, "LOW": GREEN}; return m.get(p, None)
def ifill(i):
    m = {"SEVERE": RED, "MAJOR": YELLOW, "MINOR": GREEN}; return m.get(i, None)

# -- Data --
R = [
{"id":"R01","name":"全权接受、无任何异议构成过度权利放弃 -- 显失公平风险","type":"争议风险+效力风险","layer":"条款层","clause":"Art I.3, Art II.1","desc":"甲方两次作出全权接受、无任何异议声明，覆盖单位9增资扩股及第一轮融资导致的6%至5%稀释。若认定为权利放弃，放弃范围远超合理商业安排，可能因显失公平被撤销（民法典151条）；若认定为格式免责条款，可能依民法典497条被认定无效。最严重后果：甲方在被持续稀释后无法以稀释不公平为由寻求司法救济。","law":"民法典第151条（显失公平）、第497条（格式条款无效）、第6条（公平原则）、第7条（诚信原则）","prob":"HIGH","impact":"SEVERE","level":"HIGH","advice":"【预防】删除全权接受、无任何异议表述，改为知悉并同意上述特定增资安排；增加本同意不构成对任何未来稀释的预先豁免限制性表述。\n【控制】甲方应在每次融资发生时单独出具书面同意。\n【补救】如已签署且发生严重不公平稀释，可依据民法典第151条在知道或应当知道撤销事由之日起1年内主张撤销。","confidence":"HIGH"},
{"id":"R02","name":"第一轮融资与第二轮融资区分标准完全缺失 -- 核心结构缺陷","type":"争议风险+履约风险","layer":"条款层+交易层","clause":"Art II.1","desc":"两个融资轮次的门槛均为融资额不小于数字金额1 -- 完全相同。无时间、投资人、估值或其他标准区分两者。丙方可以：(1)将天使轮、天使+轮、Pre-A轮作为第一轮融资之前的前置轮次处理，使甲方在被宣告第一轮融资前已被多次稀释；(2)将实质上的多轮合并命名为第一轮融资；(3)始终不宣告完成第二轮融资，使反稀释保护永不生效。","law":"民法典第142条（意思表示解释）、第510条（补充规则）、第498条（格式条款歧义时不利于提供方的解释）","prob":"HIGH","impact":"SEVERE","level":"HIGH","advice":"【预防】明确第一轮融资定义，增加至少一项独立区分标准（时间标准、投资人标准）；增加禁止规避条款 -- 丙方不得通过拆分融资、关联方代持、名义轮次等方式规避本条反稀释义务。\n【控制】甲方应要求丙方在每次拟融资时书面通知并说明是否构成第一轮融资。\n【补救】争议中主张依民法典第498条作出不利于提供方（丙方）的解释。","confidence":"HIGH"},
{"id":"R03","name":"反稀释保护严重滞后 -- 仅第二轮起生效，且定义缺失","type":"履约风险","layer":"交易层","clause":"Art II.1","desc":"甲方的反稀释保护从第二轮融资才开始生效。首轮6%至5%稀释已不可逆（降幅16.7%）。更严重的是，不做任何稀释的定义完全缺失 -- 是保持股权比例不变（完全棘轮）还是保持持股数量不变？若仅保持数量不变，持股比例仍会因总股本增加而下降，反稀释保护完全失效。","law":"民法典第510条（约定不明确的补充规则）；商业惯例（标准VC条款反稀释保护从首轮即生效）","prob":"HIGH","impact":"SEVERE","level":"HIGH","advice":"【预防】明确不做任何稀释定义 -- 甲方股权比例（持股数除以公司总股本）应保持不变，如因后续融资导致比例下降，丙方应无偿向甲方转让相应股份；将反稀释保护提前至首轮融资。\n【控制】每次融资后立即核算持股比例变化。\n【补救】主张依民法典第510条进行补充解释。","confidence":"HIGH"},
{"id":"R04","name":"融资信息披露义务完全缺失 -- 甲方知情同意基础不成立","type":"争议风险","layer":"条款层+交易层","clause":"Art I.3, Art II.1","desc":"甲方签署时可能对以下关键信息完全不知情：单位9的具体身份和增资规模；第一轮融资的投资人、估值、条款；历次增资对甲方股权经济价值的实际影响；未来融资计划。在信息严重不对称下作出的全权接受声明，意思表示的真实性和完整性存在重大疑问。民法典第147条（重大误解）和第151条（显失公平）为甲方提供潜在救济路径。","law":"民法典第147条（重大误解）、第500条（缔约过失责任）、公司法（2024修订）第57条（股东知情权）","prob":"HIGH","impact":"SEVERE","level":"HIGH","advice":"【预防】增加融资信息披露条款 -- 丙方或标的公司在拟进行任何股权融资前，应至少提前30日书面通知甲方融资条款；要求丙方在附件中完整披露截至签署日的全部融资历史，并对真实性和完整性作出陈述与保证。\n【控制】甲方应积极行使Art III.1财务检查权，持续跟踪公司资本变动。\n【补救】如发现丙方隐瞒重大融资信息，可依民法典第500条主张缔约过失责任。","confidence":"HIGH"},
{"id":"R05","name":"丙方可单方操控融资节奏 -- 甲方无参与权或否决权","type":"履约风险+争议风险","layer":"交易层","clause":"Art II.1","desc":"丙方作为61.1%大股东和法定代表人，可单方决定：融资时机（可在甲方激励条件未成就时融资最大化稀释效果）；融资轮次命名（可规避反稀释）；投资人选择（可以低价向关联方增发）；估值（丙方无义务为甲方争取有利估值）。甲方（5-6%）无否决权，增资决议需2/3以上表决权通过 -- 丙方61.1%低于66.67%，需联合至少一个其他股东，给甲方留下有限协商空间。","law":"公司法（2024修订）第66条（股东会决议 -- 增资需2/3以上表决权）","prob":"HIGH","impact":"SEVERE","level":"HIGH","advice":"【预防】增加甲方保护性否决权 -- 以下事项须经甲方事先书面同意：(a)以低于公允估值基准的价格增发新股；(b)对甲方持股产生稀释影响的关联方交易；(c)任何可能导致反稀释条款被规避的融资安排。\n【控制】甲方应密切关注标的公司工商登记变更。\n【补救】如丙方以明显低价向关联方增发，可主张丙方违反诚信义务，依据Art XIV.3或Art XIV.4要求回购。","confidence":"HIGH"},
{"id":"R06","name":"两次全权接受声明的效力重叠与体系矛盾","type":"争议风险","layer":"条款层","clause":"Art I.3, Art II.1","desc":"Art I.3针对单位9增资扩股作出全权接受声明；Art II.1针对第一轮融资再次作出全权接受声明。若为同一事件：甲方对同一稀释做出两次声明，Art II.1完全冗余。若为不同事件：甲方先被单位9稀释，再被第一轮融资稀释 -- 实际被稀释了两次且幅度远超6%至5%，协议文本存在误导。需要明确单位9增资与第一轮融资的关系。","law":"民法典第142条（意思表示解释）、第498条（格式条款歧义时不利于提供方的解释）","prob":"MID","impact":"SEVERE","level":"HIGH","advice":"【预防】在协议中明确：(1)单位9增资是否即为第一轮融资；(2)若非，要求丙方在附件中披露完整的稀释路径图。\n【控制】核查标的公司工商变更记录。\n【补救】如发现丙方隐瞒额外稀释事件，可主张欺诈（民法典第148条）或缔约过失。","confidence":"MID"},
{"id":"R07","name":"代持结构加剧稀释风险 -- 乙方无力保护甲方","type":"履约风险","layer":"交易层","clause":"Art VI, Art VIII, Art IX + Art II.1","desc":"乙方为无偿代理（Art IX.1），对代持股份无经济利益，无动力为甲方主张权利。当丙方通过稀释侵害甲方利益时，乙方既不关心也不主动反对。乙方在工商登记中为股东，但实际投票行为可能不按甲方指示执行 -- 如果乙方实际上更听命于丙方，甲方将完全丧失投票权。丙方可能利用甲方不愿显名的心理进行施压。","law":"民法典第925条（间接代理）、公司法司法解释三第24条（实际出资人与名义股东）","prob":"MID","impact":"MAJOR","level":"MID","advice":"【预防】增加甲方在股东会上以自己名义直接投票的不可撤销授权条款；要求丙方书面确认乙方投票按甲方指示执行；增加甲方有权随时更换代持人条款。\n【控制】甲方应保留所有向乙方发出书面投票指示的记录。\n【补救】如乙方未按甲方指示投票，可追究乙方违反Art VIII.6（诚实信用义务）的违约责任。","confidence":"MID"},
{"id":"R08","name":"股权激励的对冲效果不确定 -- 稀释豁免的对价缺失","type":"履约风险","layer":"交易层","clause":"Art I.3","desc":"协议结构上，甲方放弃稀释异议权的对价是两阶段各3%的股权激励。但稀释是确定的（单位9增资和第一轮融资必然发生），激励是不确定的（条件模糊、控制权在丙方）。如果激励条件永远不成就：甲方净损失1%（确定）而获得0%补偿（不确定）。这是确定损失交换不确定收益的结构性不对等。","law":"民法典第158条（附条件的民事法律行为 -- 当事人为自己的利益不正当地阻止条件成就的，视为条件已成就）","prob":"MID","impact":"MAJOR","level":"MID","advice":"【预防】增加股权激励条件的客观化标准（如种子客户需定义付费金额、合作期限等）；增加若因丙方原因导致激励条件无法在合理期限内成就，甲方有权要求丙方以现金或股权补偿。\n【控制】甲方应留存与种子客户验证、订单获取相关的全部沟通记录。\n【补救】如丙方恶意阻止激励条件成就，可依民法典第158条第2句主张条件视为已成就。","confidence":"MID"},
{"id":"R09","name":"格式条款风险 -- 稀释豁免条款可能被认定无效","type":"效力风险","layer":"环境层","clause":"Art I.3, Art II.1","desc":"若稀释豁免条款由丙方或其律师单方拟定、甲方无实质协商修改机会，则可能被认定为格式条款（民法典第496条）。虽条款有请注意的前置标记，但不足以构成采取合理的方式提示对方注意（民法典第496条第2款）。该条款不合理地免除或者减轻其责任、加重对方责任、限制对方主要权利 -- 可能构成民法典第497条无效情形。","law":"民法典第496条（格式条款定义和订入控制）、第497条（格式条款无效情形）、第498条（歧义解释规则）；最高法合同编通则解释第10条","prob":"MID","impact":"SEVERE","level":"MID","advice":"【预防（对甲方有利的不作为）】甲方可保留该条款的格式条款特征证据（原始版本、无修订痕迹、多份同类协议使用相同措辞），为未来争议中的格式条款无效主张奠定基础。\n【补救】争议中积极主张该条款因违反民法典第497条而无效。","confidence":"MID"},
{"id":"R10","name":"甲方个人原因隐名可能削弱法庭立场 -- 不洁之手风险","type":"争议风险","layer":"环境层","clause":"鉴于条款","desc":"若甲方未来主张稀释豁免条款无效或可撤销，丙方可能抗辩：甲方本身选择隐名代持（可能涉及规避竞业限制、保密义务、公务员经商限制等），甲方自身存在不洁行为。若甲方系为规避对第三方的义务而隐名，法院在评估显失公平时可能考虑甲方自身过错。在代持协议无效的情形下（如代持目的违法），甲方可能无法依协议主张权利。","law":"民法典第7条（诚信原则）、第153条（违反强制性规定或公序良俗的民事法律行为无效）、公司法司法解释三第24条（代持协议效力）","prob":"LOW","impact":"MAJOR","level":"MID","advice":"【预防】如甲方的个人原因仅为个人隐私考虑而非法律规避，应在鉴于条款中予以明确；如涉及法律规避风险，建议签约前寻求法律意见。\n【补救】在争议中，甲方应主动说明并证明个人原因不涉及违法事项。","confidence":"LOW"},
{"id":"R11","name":"Art XIV.3丙方回购对价严重不足 -- 无法弥补稀释损失","type":"履约风险","layer":"条款层","clause":"Art XIV.3","desc":"丙方侵占公司经济利益导致亏损时，甲方回购对价仅为实际出资额加银行两倍LPR利息 -- 仅保护本金加利息，无法分享公司成长价值。与Art XIV.1不对等：对乙方的违约回购对价为历史最高估值（有利甲方），但对丙方（实控人）的侵占行为，回购对价反而仅为出资额加2倍LPR（不利甲方）。与稀释豁免叠加：甲方已被稀释至5%，再按原始出资额被回购 -- 丧失股权增值全部上行空间。","law":"民法典第584条（损害赔偿范围 -- 包括可得利益损失）、第585条（违约金低于实际损失可请求增加）","prob":"MID","impact":"MAJOR","level":"MID","advice":"【预防】修改Art XIV.3回购对价 -- 取以下两项之较高者：(a)实际出资额加2倍LPR利息；(b)以最近一轮融资估值为基础计算的甲方持股公允价值。\n【控制】甲方应定期审查公司财务状况。\n【补救】如回购对价明显偏低，可主张按民法典第584条实际损失赔偿。","confidence":"HIGH"},
{"id":"R12","name":"协议生效条款技术性遗漏 -- 丙方非签署方","type":"效力风险+争议风险","layer":"条款层","clause":"Art XIII.1, Art XV, Art XVI","desc":"协议明确将丙方列为当事人且丙方承担多项实质性义务。但Art XIII.1仅约定自甲、乙双方签订之日起生效 -- 未提及丙方签署。Art XVI约定签署双方各执一份，签署双方表述与三方结构矛盾。如果丙方未签署或主张其仅是确认方而非签署方，丙方在本协议项下的全部义务（反稀释、回购等）的合同基础可能被质疑。","law":"民法典第490条（合同成立 -- 当事人均签名时成立）、第502条（合同生效）","prob":"LOW","impact":"SEVERE","level":"MID","advice":"【预防】修改Art XIII.1为本协议自甲、乙、丙三方签订之日起生效；确保三方均在协议每一页签字盖章。\n【控制】核查实际签署版本的丙方签署状态。","confidence":"MID"},
{"id":"R13","name":"乙方死亡或失踪或丧失行为能力 -- 无继承条款","type":"履约风险","layer":"交易层","clause":"合同未约定","desc":"协议未约定若乙方死亡、丧失民事行为能力或下落不明时代持股份的处理方式。根据民法典第六编（继承），乙方名下的代持股份可能被其继承人继承 -- 继承人是否知悉或承认代持安排完全不确定。最坏情形：继承人拒绝承认代持，要求作为乙方自有财产继承。","law":"民法典第六编（继承） -- 继承人继承被继承人的全部财产权利","prob":"LOW","impact":"SEVERE","level":"MID","advice":"【预防】增加条款 -- 如乙方死亡、被宣告失踪或丧失民事行为能力，甲方有权单方决定将代持股份转移至甲方或甲方指定的新代持人名下，乙方的继承人、财产管理人或监护人应无条件配合。\n【控制】甲方应确保乙方亲属知悉代持安排。","confidence":"HIGH"},
{"id":"R14","name":"通知条款对甲方不利 -- 2日视为送达过短","type":"争议风险","layer":"条款层","clause":"Art XI","desc":"2日视为送达对于邮寄通知而言过短；即时通讯的当日视为送达 -- 若乙方不在线或未读，当日即视为送达可能不具合理性。","law":"民法典第137条（意思表示的生效）","prob":"LOW","impact":"MINOR","level":"LOW","advice":"延长至快递签收后第3个工作日视为送达；即时通讯增加发送方收到对方确认回执时视为送达。","confidence":"HIGH"},
{"id":"R15","name":"争议管辖条款对甲方可能不便","type":"争议风险","layer":"条款层","clause":"Art XV","desc":"标的公司注册地可能与甲方居住地不一致，增加甲方维权成本。","law":"民事诉讼法第23条（合同纠纷管辖 -- 被告住所地或合同履行地）","prob":"LOW","impact":"MINOR","level":"LOW","advice":"协商修改为原告所在地人民法院或甲方住所地人民法院。","confidence":"MID"},
{"id":"R16","name":"保密条款仅约束协议双方 -- 丙方约束不明","type":"履约风险","layer":"条款层","clause":"Art XII","desc":"保密义务主体为协议双方，但协议为三方结构。丙方是否受保密条款约束不明确。","law":"民法典第501条（缔约过程中的保密义务）","prob":"LOW","impact":"MINOR","level":"LOW","advice":"修改为甲、乙、丙三方。","confidence":"HIGH"},
{"id":"R17","name":"协议终止条件单一 -- 仅以显名化完成为终止","type":"履约风险","layer":"条款层","clause":"Art XIII.2","desc":"未约定双方协商一致解除、一方根本违约导致解除等情形下的终止效力。特别是若丙方严重违约（如Art XIV.4），甲方在获得回购之前，代持协议是否终止不明确。","law":"民法典第563条（法定解除权）、第566条（解除后的处理）","prob":"LOW","impact":"MINOR","level":"LOW","advice":"增加解除条款 -- 甲方在丙方发生Art XIV.3或Art XIV.4违约情形时，有权单方解除本协议并要求立即回购。","confidence":"HIGH"},
{"id":"R18","name":"薪资条款构成甲方额外负担 -- 盈利确认由丙方控制","type":"履约风险","layer":"交易层","clause":"Art V","desc":"甲方作为联合创始人可能全职投入但薪酬为零（直至盈利）。盈利状态由丙方控制的财务报表决定，丙方可通过会计手段延迟确认盈利。甲方可能长期无薪酬劳动。","law":"无直接法律依据；属商业条款不公问题","prob":"MID","impact":"MINOR","level":"LOW","advice":"约定最低基本薪酬；约定盈利确认的独立审计机制。","confidence":"HIGH"},
]

A = [
{"id":"A01","term":"全权接受、无任何异议","clause":"Art I.3, Art II.1","type":"语义模糊+范围模糊","interp_bad":"甲方无条件接受稀释事实和法律后果，放弃任何形式的异议权","interp_good":"甲方仅确认知悉特定增资安排，不构成对未来稀释的预先豁免","court":"若为格式条款，歧义时作不利于提供方（丙方）的解释（民法典第498条）；结合体系解释（Art VII.6甲方有权将股权转移至自身名下 -- 与稀释豁免构成体系矛盾，暗示甲方保留核心股东权利）","fix":"删除全权接受、无任何异议，改为甲方知悉并同意上述特定增资安排","level":"HIGH"},
{"id":"A02","term":"第一轮融资","clause":"Art II.1","type":"定义缺失","interp_bad":"标的公司接受的第一次外部股权融资（不论金额）","interp_good":"标的公司接受的第一次融资额不小于数字金额1的外部股权融资（天使轮、天使+轮等若金额不足门槛则不计入）","court":"文义解释：协议约定融资额须在数字金额1及数字金额1以上，故金额门槛为定义要素；但第一轮的第一无其他独立标准，实质上由丙方单方认定 -- 依民法典第142条结合诚信原则，应解释为客观上第一笔不小于数字金额1的融资","fix":"增加至少一项独立区分标准（时间/投资人/序位）；增加本协议签署后标的公司接受的首次来自现有股东以外投资人的股权融资","level":"HIGH"},
{"id":"A03","term":"不做任何稀释","clause":"Art II.1","type":"定义缺失+标准缺失","interp_bad":"保持甲方持股数量不变（稀释仅指数量减少）","interp_good":"保持甲方持股比例不变（即完全棘轮反稀释 -- 比例因总股本增加而下降即需补偿）","court":"体系解释：Art I.3和Art II.1整个条款讨论的是股权比例的变化（6%至5%），而非持股数量，因此不做任何稀释应指股权比例不变；目的解释：条款目的是保护甲方利益，仅保护数量不保护比例将使保护完全失效，不符合条款目的","fix":"明确为甲方持股比例（即甲方持股数除以标的公司总股本）应保持不变；如因后续融资使该比例下降，丙方应无偿向甲方转让相应股份","level":"HIGH"},
{"id":"A04","term":"第一轮与第二轮的区分","clause":"Art II.1","type":"标准缺失","interp_bad":"第一轮融资为单位9增资扩股（即Art I.3所述事件），第二轮融资为后续专业投资机构主导的首次融资","interp_good":"第一轮融资是单位9增资扩股与后续融资的合称，第二轮融资是该合称之后的下一次融资","court":"体系解释：Art I.3与Art II.1分别使用了不同的稀释触发主体（单位9 vs 第一轮融资），说明单位9增资独立于融资轮次计数体系。依民法典第498条歧义不利于提供方规则，应解释为单位9增资等于第一轮融资（对甲方最有利的解释 -- 甲方仅被稀释一次而非两次）","fix":"明确单位9增资是否即为第一轮融资；在附件中列出全部融资轮次的时间线","level":"HIGH"},
{"id":"A05","term":"种子客户验证","clause":"Art I.3（股权激励）","type":"定义缺失+标准缺失","interp_bad":"任何客户对公司产品进行测试并给予正面反馈","interp_good":"付费客户且合作持续一定期限、满足特定客观指标","court":"行业习惯（B2B销售中种子客户通常指付费意愿强、合作稳定的早期客户）+ 诚信原则（不能以丙方单方满意的模糊标准限制甲方的激励权利）","fix":"客观化定义：种子客户指与标的公司签订正式采购合同，且合同金额不低于X万元、合作期限不少于Y个月的客户","level":"MID"},
{"id":"A06","term":"亲属关系由未使用亲属关系的一方进行解释","clause":"Art IV","type":"权利配置失衡","interp_bad":"甲方（作为未使用方）有权单方解释哪些人属于亲属关系","interp_good":"丙方（作为未使用方）也有权单方解释；双方解释冲突时无解决机制","court":"字面含义：确实赋予未使用亲属关系方解释权 -- 但若双方均未使用，则双方均有解释权，产生冲突。此条款的不对等性在于：丙方（大股东）比甲方（小股东）更可能需要使用亲属，甲方的解释权实质上成为对丙方用人权的制衡。","fix":"修改为：亲属关系包括配偶、直系血亲、三代以内旁系血亲及近姻亲关系；如有争议，以本定义为准","level":"MID"},
]

H = [
{"id":"H1","name":"善意假设：甲方系充分知情的联合创始人，自愿接受稀释换取公司发展","scenario":"甲方作为联合创始人，在签署协议时已经充分了解了单位9增资和第一轮融资的商业安排，自愿以接受稀释为代价参与公司创业。","evidence":"甲方为引入的联合创始人，应当具备一定的商业判断能力；协议中请注意标记提供了提示；甲方未提出修改意见即签署。","consequence":"该假设成立的情形下，稀释豁免条款属于有效商业安排，甲方应自行承担稀释后果。","prevention":"无需特别防范（因为风险不存在）","likelihood":"低（甲方选择隐名代持配合稀释条款的不对等程度超常规商业安排，更可能是信息不对称而非完全知情同意）"},
{"id":"H2","name":"信息不对称假设：甲方在被引入时未被告知完整的增资计划和稀释影响","scenario":"丙方在引入甲方时，仅告知了6%股权和联合创始人身份，但未完整披露单位9增资扩股和后续融资轮次的具体安排（估值、幅度、时间），甲方在信息严重不对称的情况下签署了包含稀释豁免的协议。","evidence":"协议未附任何融资计划的说明或披露文件；甲方不具备获取融资信息的独立渠道（隐名）；全权接受、无任何异议的绝对化表述超出正常商业条款范围。","consequence":"该假设成立的情形下，甲方可依据民法典第147条（重大误解）或第151条（显失公平）主张撤销稀释豁免条款。举证难点：需证明甲方签约时确实不知情。","prevention":"甲方应保留签约前的全部沟通记录（微信、邮件等）作为信息不对称的证据；要求丙方在协议附件中补充披露融资历史。","likelihood":"中高（协议缺少信息披露条款本身即为客观证据；甲方隐名身份和个人原因进一步削弱了甲方获取信息的能力）"},
{"id":"H3","name":"结构压迫假设：丙方利用实控人地位和甲方的隐名弱势，通过稀释条款系统性压缩甲方权益","scenario":"丙方利用其作为61.1%大股东和实控人的地位，以及甲方因个人原因必须隐名的弱势，在协议中设置了极端不对等的稀释豁免条款，实质上将甲方的股权投资转化为类似次级债的低保护投资。","evidence":"稀释豁免+反稀释保护滞后+融资信息披露缺失+丙方单方控制+股权激励条件模糊 -- 五重机制的叠加已超出正常商业安排；甲方除稀释豁免外还有：零薪资直至盈利（Art V）+ 出资违约责任（Art VI.1）+ 竞业限制（Art VII.7） -- 甲方承担的义务远多于权利。","consequence":"该假设成立的情形下，稀释豁免条款可能因违反公序良俗原则（民法典第153条第2款）或构成显失公平（第151条）而无效或可撤销。另可考虑适用民法典第154条（恶意串通损害他人合法权益 -- 若乙方与丙方存在串通）。","prevention":"甲方应在签约前尝试协商修改上述不对等条款，保留协商过程的书面记录 -- 若丙方拒绝任何实质性修改，此事实本身可佐证结构压迫；考虑整个投资安排中是否有其他文件（如投资协议、股东协议）提供额外保护。","likelihood":"中（五重机制的叠加是有力的客观证据，但结构压迫的司法认定门槛较高，需结合具体事实）"},
{"id":"H4","name":"格式条款假设：稀释豁免条款为丙方或其律师单方拟定，甲方无实质协商机会","scenario":"该协议由丙方或其法律顾问起草，甲方作为个人投资者或联合创始人无法律顾问参与，稀释豁免条款未经任何实质性协商修改即纳入协议，构成民法典第496条意义上的格式条款。丙方未以加粗、单独签署等方式履行合理提示说明义务，甲方可主张该条款不成为合同内容（第496条第2款）；该条款因不合理地免除丙方责任、限制甲方主要权利而无效（第497条）。","evidence":"协议文本高度专业化、结构完整，具有单方起草的特征；全权接受、无任何异议的表述在商业合同范本中不常见 -- 若甲方能证明该协议来源于丙方且甲方无修改机会，格式条款的认定可能性较大。","consequence":"若该假设成立且被法院采纳，稀释豁免条款将自始不产生效力。格式条款无效不影响协议其他条款效力（第156条）。","prevention":"甲方应保留丙方提供协议原始版本的全部证据（包括文件创建日期、发送记录、修改痕迹等）；获取丙方与其他投资人签署的同类协议以证明格式条款的重复使用。","likelihood":"中（格式条款的认定取决于未与对方协商和重复使用两个要素的证明 -- 需具体证据支撑）"},
]

C = [
{"id":"S1","scenario":"稀释豁免条款被认定无效 -- 甲方主张撤销或确认无效成功","law":"民法典第151条或第497条","trigger":"Art I.3, Art II.1稀释豁免条款被撤销或认定无效","low":"甲方恢复对稀释提出异议的完整权利","mid":"法院可能判令丙方对不当稀释造成的损失进行赔偿；赔偿金额需评估甲方股权在被稀释前后的价值差异 -- 涉及公司估值这一复杂事实问题","high":"甲方持股比例恢复至稀释前水平（如6%）或获得等值经济补偿","dispute":"需承担举证责任（证明显失公平或格式条款或重大误解） -- 举证成本高（可能需要司法鉴定或专家证人证明估值）","risk":"高（因举证难度大，实际恢复或获赔的不确定性高）","confidence":"MID"},
{"id":"S2","scenario":"丙方以第一轮融资名义完成两笔融资 -- 仅第二笔触发反稀释","law":"民法典第142条（意思表示解释）","trigger":"Art II.1反稀释保护仅从第二轮起生效","low":"甲方在第一笔融资（名义上的第一轮）中被稀释一次，在第二笔（名义上的第二轮）中不再被稀释","mid":"但由于第一轮融资界定模糊，丙方完全可以将天使+轮和Pre-A轮包装为第一轮融资的一部分，在第二轮正式宣布前完成多次稀释","high":"假设6%至天使+轮后4.5%至Pre-A轮后约3.4%至第二轮后维持在3.4% -- 甲方持股从6%降至3.4%，损失近半","dispute":"甲方主张第一轮融资应解释为客观上首次不小于数字金额1的融资的难度 -- 需结合签订时的具体沟通记录","risk":"高（以投资额X万元估算，3.4%与6%的差额可能对应相当可观的经济损失）","confidence":"MID"},
{"id":"S3","scenario":"甲方主张显失公平撤销权的时效风险","law":"民法典第152条第1项","trigger":"Art I.3, Art II.1稀释豁免条款","low":"甲方自知道或应当知道撤销事由之日起1年内可行使撤销权；自民事法律行为发生之日起5年内未行使的，撤销权消灭","mid":"撤销权行使的法律后果：稀释豁免条款自始无效（民法典第155条），甲方恢复对稀释提出异议的权利","high":"时效过期后甲方永久丧失撤销权 -- 只能依赖其他法律路径（如格式条款无效主张，不受除斥期间限制）","dispute":"知道或应当知道撤销事由的起算点争议极大 -- 是签署协议之日？还是每次稀释实际发生之日？还是甲方发现丙方隐瞒信息之日？","risk":"中高（若签署日已过1年，撤销权消灭风险极大；但可保留格式条款无效主张作为替代路径）","confidence":"HIGH"},
{"id":"S4","scenario":"乙方死亡 -- 代持股份被继承人主张继承","law":"民法典第1121-1122条（继承财产范围）","trigger":"协议未约定乙方死亡情形","low":"乙方继承人主张代持股份为被继承人名下财产，属于遗产范围","mid":"甲方需证明代持关系 -- 乙方的声明（Art VI.2-3）是关键证据；但继承人可主张：(1)代持协议的真实性无法确认；(2)即使代持成立，继承人无义务继续代持","high":"甲方可能被迫对继承人提起确权之诉 -- 诉讼周期长（一审6个月至1年加二审），期间代持股份处于冻结或争议状态，甲方无法行使任何股东权利","dispute":"若甲方无法证明代持真实性和资金来源，可能永久丧失投资对应的股权","risk":"低（概率）但极高（影响） -- 属于小概率-大影响的长尾风险","confidence":"MID"},
]

AM = [
{"priority":"T1-1","clause":"Art I.3, Art II.1","problem":"删除全权接受、无任何异议","text":"甲方知悉并同意上述特定增资安排。本同意仅适用于本协议明确列明的增资事项，不构成甲方对任何未来增资、融资或其他可能导致股权稀释的事件的预先豁免或放弃异议权。","reason":"消除甲方权利放弃范围过宽的风险（R01），将绝对化的权利放弃降级为特定事项的知情同意","strategy":"核心条款，不可妥协。如丙方坚持保留，至少增加本同意不适用于本协议未明确列明的增资事项的限制性表述。"},
{"priority":"T1-2","clause":"Art II.1","problem":"明确定义第一轮融资","text":"第一轮融资指标的公司在签署本协议后，接受的首次来自现有股东及其关联方以外的投资人（专业投资机构）的、融资额不低于数字金额1的股权融资。丙方不得通过拆分融资交易、引入关联方代持、名义轮次命名或其他方式规避本条款的适用。","reason":"消除融资轮次区分标准缺失的风险（R02），封堵丙方通过命名套利规避反稀释义务的路径","strategy":"核心条款。如丙方拒绝增加投资人标准，至少保留禁止规避条款。"},
{"priority":"T1-3","clause":"Art II.1","problem":"反稀释保护提前至首轮融资","text":"未来，标的公司在完成任何后续融资（包括第一轮融资及之后的所有轮次）的过程中，甲方的持股比例应保持不变。如因标的公司增资扩股导致甲方持股比例下降，丙方应于增资扩股完成后30日内，无偿向甲方（由乙方代持）转让相应股份以恢复甲方持股比例。","reason":"解决反稀释保护滞后问题（R03），确保甲方从首轮即受保护；明确定义为完全棘轮反稀释（比例不变）","strategy":"如丙方坚持从第二轮开始反稀释，则要求以加权平均方式对首轮稀释进行事后补偿（即在反稀释生效时将甲方持股比例一次性恢复至首轮稀释前的水平）。"},
{"priority":"T1-4","clause":"新增条款","problem":"增加融资信息披露义务","text":"丙方或标的公司在拟进行任何增资、融资或发行新股前，应至少提前30日向甲方发出书面通知，通知内容应包括但不限于：(a)投资人身份及背景；(b)投前和投后估值；(c)融资金额及每股价格；(d)主要投资条款摘要；(e)对甲方持股比例和股权价值的预计影响。甲方有权在前述30日内向丙方提出书面意见。","reason":"解决融资信息披露缺失问题（R04），确保甲方在充分知情的基础上行使权利","strategy":"甲方核心保护条款。如30日通知期无法满足，可缩短至15日。"},
{"priority":"T1-5","clause":"新增条款","problem":"增加甲方保护性条款（否决权）","text":"以下事项须经甲方事先书面同意方可实施：(a)以低于届时标的公司最近一轮融资每股价格80%的价格增发新股；(b)与丙方或其关联方进行的任何可能导致甲方股权被稀释的交易；(c)对融资轮次的命名或分类方式如可能影响本协议反稀释条款的适用。","reason":"解决甲方无参与权和否决权问题（R05），为甲方提供对极端稀释行为的否决能力","strategy":"保护性条款是VC投资的行业标准。如丙方拒绝，至少争取(a)和(c)项。"},
{"priority":"T2-1","clause":"Art XIV.3","problem":"修改丙方回购对价计算方式","text":"丙方作为标的公司实控人，在经营过程中，侵占标的公司经济利益，导致标的公司非正常状态下亏损的，甲方有权要求丙方对甲方的股权进行回购，回购对价取以下两项之较高者：(a)甲方的实际出资额，加上资金被占用期间的银行两倍贷款市场报价利率计算的利息；(b)以回购时标的公司最近一轮融资估值为基础计算的甲方持股对应的公允价值。","reason":"解决回购对价不足问题（R11），确保甲方在公司成功时能分享成长价值","strategy":"重要条款。如丙方拒绝公允价值选项，至少将利息计算基数从出资额改为出资额加累计应分配利润。"},
{"priority":"T2-2","clause":"Art XIII.1","problem":"修正签署和生效条款","text":"本协议自甲、乙、丙三方签订之日起生效。本协议一式五份，甲、乙、丙三方各执一份，标的公司留存一份，公证留存一份，每份均具有同等法律效力。","reason":"解决丙方签署效力不明确问题（R12），确保三方均为明确签署方","strategy":"技术性修补，丙方大概率不会拒绝。"},
{"priority":"T2-3","clause":"新增条款","problem":"增加乙方死亡或失能继承条款","text":"如乙方死亡、被宣告失踪或丧失民事行为能力，甲方有权单方决定：(a)将代持股份直接转移至甲方名下完成显名化；(b)或指定新的代持人继续代持。乙方的继承人、财产管理人或监护人应在收到甲方书面通知后15日内无条件配合签署相关文件及办理工商变更手续。乙方的继承人、财产管理人或监护人拒绝或迟延配合的，应向甲方赔偿因此产生的全部损失。","reason":"解决乙方死亡或失能无继承条款问题（R13），为甲方提供单方处置权","strategy":"重要条款。乙方应予以配合，因条款处理的是小概率但对甲方致命的事件。"},
{"priority":"T3-1","clause":"Art XI","problem":"修改通知送达条款","text":"甲方向乙方发出书面通知的，以快递签收后第3个工作日视为送达；以即时通讯方式发送的，以发送方收到对方确认回执时视为送达；以电子邮件方式发送的，以邮件进入对方指定邮箱系统时视为送达。","reason":"解决通知期限过短问题（R14）","strategy":"技术性优化，可直接提议。"},
{"priority":"T3-2","clause":"Art I.3（股权激励）","problem":"客观化股权激励条件","text":"种子客户指与标的公司签订正式采购合同，且合同金额不低于[具体金额]万元、合作期限不少于[具体月数]个月的客户。喷头订单(完整)指标的公司与客户签订的、覆盖完整产品型号和数量的正式采购订单。标的公司或其子公司当年的主营业务销售额以经审计的年度财务报告所载数据为准。","reason":"解决激励条件模糊问题（R08）","strategy":"如有谈判空间则增加，否则作为履约中重点关注的监控事项。"},
{"priority":"T3-3","clause":"Art XII","problem":"修正保密条款主体","text":"甲、乙、丙三方对本协议履行过程中所接触或获知的其他方的任何商业信息均有保密义务，除非有明显的证据证明该等信息属于公知信息或者事先得到对方的书面授权。该保密义务在本协议终止后仍然继续有效。任何一方因违反保密义务而给对方造成损失的，均应当赔偿对方的相应损失。","reason":"解决丙方保密义务不确定问题（R16）","strategy":"技术性修补。"},
{"priority":"T3-4","clause":"Art XIII.2","problem":"增加合同解除条款","text":"本协议在下列情形下终止：(a)乙方将代持股权全部移交甲方并完成工商变更登记；(b)甲、乙、丙三方协商一致书面同意终止；(c)因一方的根本违约行为导致本协议目的不能实现的，守约方有权书面通知违约方解除本协议，解除的后果按照本协议第十四条的约定执行。","reason":"解决终止条件单一问题（R17）","strategy":"补充性条款。"},
{"priority":"T3-5","clause":"Art V","problem":"增加最低薪酬保障","text":"在单位2不能实现盈利之前，甲方每月领取基本生活费[具体金额]元；如标的公司实现盈利，标的公司应一次性补齐甲方在标的公司所需要领取的历年薪资总额（计算标准为每月[具体金额]元），金额由股东会讨论决定。盈利状态的认定以经甲方认可的独立审计机构出具的审计报告为准。","reason":"减轻甲方零薪资负担（R18）","strategy":"视谈判情况而定。如丙方坚决反对，至少保留盈利审计的独立确认权。"},
]

Q = [
{"id":"Q1","q":"全权接受、无任何异议在法律上是否构成有效的权利放弃？放弃范围如何界定？","a":"是有效的权利放弃声明，但其效力受以下限制：(1)若构成格式条款，因违反民法典第497条（不合理限制对方主要权利）可能被认定无效；(2)若构成显失公平，可在1年内依民法典第151条主张撤销；(3)即使有效，放弃范围应限缩解释为仅针对协议明确列明的两个特定增资事件（单位9增资和第一轮融资），不构成对未明确列明的其他增资事件的预先豁免 -- 这一限缩解释有民法典第142条（诚信解释）和第498条（歧义不利于提供方）的支撑。"},
{"id":"Q2","q":"若甲方入股前公司已完成天使轮或A轮等增资，Art II.1的第一轮融资是否包含这些历史轮次？","a":"大概率不包含。Art II.1的第一轮融资须结合Art I.3的体系解释 -- 协议的核心时间线是：甲方入股（6%）然后单位9增资（6%至5%）然后第一轮融资（5%进一步稀释？）然后第二轮融资及后续（不做稀释）。第一轮融资是甲方入股后的未来事件，不追溯涵盖甲方入股前的历史轮次。但需注意：如果甲方入股前的历史轮次对甲方本次入股价格或估值的影响未被充分披露，甲方可能就整体投资安排主张重大误解或欺诈。"},
{"id":"Q3","q":"反稀释保护仅在第二轮及之后生效是否构成对少数股东的不公平对待？","a":"是。标准VC投资实践中，反稀释保护通常从投资人的首轮参与即生效。将反稀释保护推迟至第二轮，与VC行业惯例严重偏离。结合甲方5-6%的小股东身份、缺乏董事会席位、无融资否决权等因素，该安排整体上构成对少数股东的系统性不利。在法律上，能否构成民法典第151条显失公平，取决于具体事实 -- 特别是甲方在签约时是否被充分告知了这一安排，以及是否有机会进行协商修改。"},
{"id":"Q4","q":"甲方因个人原因选择代持（隐名）-- 这一背景是否削弱了甲方对稀释条款提出异议的能力？","a":"有一定削弱但不致命。若甲方的个人原因涉及规避法律或合同义务，丙方可能以此作为不洁之手抗辩，影响法庭对甲方诚信的评价。但：(1)稀释豁免条款的公平性审查是独立的合同法问题，甲方是否隐名不直接改变该条款的效力判断；(2)隐名代持本身在不涉及规避强制性规定的情形下，是受公司法司法解释三第24条保护的有效安排；(3)丙方作为代持安排的参与者和获益方（通过稀释扩大控制权），不应从该安排中获取额外的不当利益。建议甲方在签署前明确个人原因不涉及任何违法事项，并在鉴于条款中予以适当表述。"},
{"id":"Q5","q":"如何重构稀释条款以平衡公司融资灵活性与少数股东保护？","a":"推荐的重构方案（从强到弱三个版本）：【强版本】完全棘轮反稀释从首轮即生效 + 甲方对低价融资有否决权 + 融资信息披露义务 + 甲方优先认购权。【中版本】加权平均反稀释从首轮即生效（按广义加权平均公式计算）+ 融资信息披露义务（15日通知期）+ 甲方在低于公允估值一定比例（如30%）的融资中有否决权。【弱版本】仅信息权 + 首轮稀释后以补发股份方式补偿（基于首轮估值与后续估值的差额）+ 禁止规避条款。建议以中版本为谈判目标，弱版本为底线。如果连弱版本都无法达成，甲方应重新评估该投资的整体风险收益比。"},
]

# -- Build workbook --
wb = openpyxl.Workbook()

# === Sheet 1: Overview ===
ws1 = wb.active
ws1.title = "ReviewOverview"
ws1.merge_cells("A1:H1")
sc(ws1, 1, 1, "Investment Nominee Holding Agreement - Legal Review Report", fill=WHITE_FILL, font=TF, align=CTR)
ws1.merge_cells("A2:H2")
sc(ws1, 2, 1, "Review Date: 2026-07-10 | Position: Party A (Beneficial Owner/Co-founder) | Focus: Dilution Waiver Clauses (Art I.3 & Art II.1)", fill=WHITE_FILL, font=SHDR_F, align=CTR)
ws1.merge_cells("A3:H3")
sc(ws1, 3, 1, "Document: 202401 Nominee Agreement (Desensitized) | Ref: [001]", fill=WHITE_FILL, font=SHDR_F, align=CTR)
row = 5
ws1.merge_cells(f"A{row}:H{row}")
sc(ws1, row, 1, "RISK SUMMARY", fill=SECTION_FILL, font=SF, align=LFT); row += 1
high_n = sum(1 for r in R if r["level"]=="HIGH")
mid_n = sum(1 for r in R if r["level"]=="MID")
low_n = sum(1 for r in R if r["level"]=="LOW")
for label, cnt, note in [
    ("Total Risks", f"{len(R)} items", "Covering clause/transaction/environment layers"),
    ("HIGH Risk", f"{high_n} items ({high_n*100//len(R)}%)", "R01-R06: Structural defects centered on dilution waiver - must fix before signing"),
    ("MID Risk", f"{mid_n} items ({mid_n*100//len(R)}%)", "R07-R13: Nominee structure, buyback pricing, inheritance gap - negotiate"),
    ("LOW Risk", f"{low_n} items ({low_n*100//len(R)}%)", "R14-R18: Technical clause optimization - monitor"),
]:
    sc(ws1, row, 1, label, fill=WHITE_FILL, font=LF, align=LFT)
    sc(ws1, row, 2, cnt, fill=WHITE_FILL, font=BF, align=CTR)
    ws1.merge_cells(start_row=row, start_column=3, end_row=row, end_column=8)
    sc(ws1, row, 3, note, fill=WHITE_FILL, font=DF, align=LFT); row += 1
row += 1
ws1.merge_cells(f"A{row}:H{row}")
sc(ws1, row, 1, "Tier 1 - Must Fix Before Signing (Non-Negotiable)", fill=RED, font=SF, align=LFT); row += 1
for a in AM:
    if a["priority"].startswith("T1"):
        sc(ws1, row, 1, f"{a['priority']}: {a['problem']} -> {a['clause']}", fill=WHITE_FILL, font=BF, align=LFT)
        ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8); row += 1
row += 1
ws1.merge_cells(f"A{row}:H{row}")
sc(ws1, row, 1, "Tier 2 - Strongly Recommended (Negotiate)", fill=YELLOW, font=SF, align=LFT); row += 1
for a in AM:
    if a["priority"].startswith("T2"):
        sc(ws1, row, 1, f"{a['priority']}: {a['problem']} -> {a['clause']}", fill=WHITE_FILL, font=DF, align=LFT)
        ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8); row += 1
row += 1
ws1.merge_cells(f"A{row}:H{row}")
sc(ws1, row, 1, "Dilution Waiver Risk Matrix (R01-R11)", fill=SECTION_FILL, font=SF, align=LFT); row += 1
for ci, h in enumerate(["ID","Risk Name","Layer","Prob","Impact","Level"], 1):
    sc(ws1, row, ci, h, fill=COLHDR_FILL, font=CHDR_F, align=CTR); row += 1
for r in R:
    if int(r["id"][1:]) > 11: continue
    sc(ws1, row, 1, r["id"], fill=WHITE_FILL, font=BF, align=CTR)
    sc(ws1, row, 2, r["name"][:80], fill=WHITE_FILL, font=DF, align=LFT)
    sc(ws1, row, 3, r["layer"], fill=WHITE_FILL, font=DF, align=CTR)
    sc(ws1, row, 4, r["prob"], fill=pfill(r["prob"]), font=BF, align=CTR)
    sc(ws1, row, 5, r["impact"], fill=ifill(r["impact"]), font=BF, align=CTR)
    sc(ws1, row, 6, r["level"], fill=rfill(r["level"]), font=BF, align=CTR); row += 1
row += 1
sc(ws1, row, 1, "[Limitations] 1.Desensitized doc, amounts/dates redacted. 2.Party A personal reasons unknown (R10). 3.Unit9 vs Round1 relationship unconfirmed (R06). 4.No counterparty background info. 5.This report is legal risk analysis reference only, not formal legal opinion.", fill=WHITE_FILL, font=LF, align=LFT)
ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
for ci, w in enumerate([8,35,20,8,8,8,55,8], 1):
    ws1.column_dimensions[get_column_letter(ci)].width = w
ws1.freeze_panes = "A6"

# === Sheet 2: Risk Register ===
ws2 = wb.create_sheet("RiskRegister")
h2 = ["ID","Risk Name","Type","Layer","Clause","Description","Legal Basis","Prob","Impact","Level","Advice","Confidence"]
ws2.merge_cells("A1:L1"); sc(ws2, 1, 1, "Full Risk Register (18 items)", fill=SECTION_FILL, font=SF, align=LFT)
for ci, h in enumerate(h2, 1): sc(ws2, 2, ci, h, fill=COLHDR_FILL, font=CHDR_F, align=CTR)
for ri, r in enumerate(R):
    row = ri+3
    sc(ws2, row, 1, r["id"], fill=WHITE_FILL, font=BF, align=CTR)
    sc(ws2, row, 2, r["name"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws2, row, 3, r["type"], fill=WHITE_FILL, font=DF, align=CTR)
    sc(ws2, row, 4, r["layer"], fill=WHITE_FILL, font=DF, align=CTR)
    sc(ws2, row, 5, r["clause"], fill=WHITE_FILL, font=DF, align=CTR)
    sc(ws2, row, 6, r["desc"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws2, row, 7, r["law"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws2, row, 8, r["prob"], fill=pfill(r["prob"]), font=BF, align=CTR)
    sc(ws2, row, 9, r["impact"], fill=ifill(r["impact"]), font=BF, align=CTR)
    sc(ws2, row, 10, r["level"], fill=rfill(r["level"]), font=BF, align=CTR)
    sc(ws2, row, 11, r["advice"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws2, row, 12, r["confidence"], fill=WHITE_FILL, font=DF, align=CTR)
for ci, w in enumerate([8,30,18,15,15,50,40,8,8,8,50,8], 1): ws2.column_dimensions[get_column_letter(ci)].width = w
ws2.freeze_panes = "A3"; ws2.auto_filter.ref = f"A2:L{len(R)+2}"

# === Sheet 3: Ambiguity Analysis ===
ws3 = wb.create_sheet("AmbiguityAnalysis")
h3 = ["ID","Ambiguous Term","Clause","Type","Interp 1 (vs A)","Interp 2 (pro A)","Court Likely Interp","Fix","Level"]
ws3.merge_cells("A1:I1"); sc(ws3, 1, 1, "Ambiguity Analysis", fill=SECTION_FILL, font=SF, align=LFT)
for ci, h in enumerate(h3, 1): sc(ws3, 2, ci, h, fill=COLHDR_FILL, font=CHDR_F, align=CTR)
for ri, a in enumerate(A):
    row = ri+3
    sc(ws3, row, 1, a["id"], fill=WHITE_FILL, font=BF, align=CTR)
    sc(ws3, row, 2, a["term"], fill=WHITE_FILL, font=BF, align=LFTT)
    sc(ws3, row, 3, a["clause"], fill=WHITE_FILL, font=DF, align=CTR)
    sc(ws3, row, 4, a["type"], fill=WHITE_FILL, font=DF, align=CTR)
    sc(ws3, row, 5, a["interp_bad"], fill=RED, font=DF, align=LFTT)
    sc(ws3, row, 6, a["interp_good"], fill=GREEN, font=DF, align=LFTT)
    sc(ws3, row, 7, a["court"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws3, row, 8, a["fix"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws3, row, 9, a["level"], fill=rfill(a["level"]), font=BF, align=CTR)
for ci, w in enumerate([8,20,15,15,35,35,45,40,8], 1): ws3.column_dimensions[get_column_letter(ci)].width = w
ws3.freeze_panes = "A3"

# === Sheet 4: Abductive Hypotheses ===
ws4 = wb.create_sheet("AbductiveHypotheses")
h4 = ["ID","Hypothesis","Scenario","Evidence","Consequence","Prevention","Likelihood"]
ws4.merge_cells("A1:G1"); sc(ws4, 1, 1, "Abductive Risk Hypotheses", fill=SECTION_FILL, font=SF, align=LFT)
for ci, h in enumerate(h4, 1): sc(ws4, 2, ci, h, fill=COLHDR_FILL, font=CHDR_F, align=CTR)
for ri, h in enumerate(H):
    row = ri+3
    sc(ws4, row, 1, h["id"], fill=WHITE_FILL, font=BF, align=CTR)
    sc(ws4, row, 2, h["name"], fill=WHITE_FILL, font=BF, align=LFTT)
    sc(ws4, row, 3, h["scenario"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws4, row, 4, h["evidence"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws4, row, 5, h["consequence"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws4, row, 6, h["prevention"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws4, row, 7, h["likelihood"], fill=YELLOW, font=DF, align=CTR)
for ci, w in enumerate([8,20,40,40,40,35,40], 1): ws4.column_dimensions[get_column_letter(ci)].width = w
ws4.freeze_panes = "A3"

# === Sheet 5: Legal Consequences ===
ws5 = wb.create_sheet("LegalConsequences")
h5 = ["ID","Scenario","Legal Basis","Trigger","Low Estimate","Mid Estimate","High Estimate","Dispute Points","Risk","Confidence"]
ws5.merge_cells("A1:J1"); sc(ws5, 1, 1, "Quantified Legal Consequences", fill=SECTION_FILL, font=SF, align=LFT)
for ci, h in enumerate(h5, 1): sc(ws5, 2, ci, h, fill=COLHDR_FILL, font=CHDR_F, align=CTR)
for ri, c in enumerate(C):
    row = ri+3
    sc(ws5, row, 1, c["id"], fill=WHITE_FILL, font=BF, align=CTR)
    sc(ws5, row, 2, c["scenario"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws5, row, 3, c["law"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws5, row, 4, c["trigger"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws5, row, 5, c["low"], fill=GREEN, font=DF, align=LFTT)
    sc(ws5, row, 6, c["mid"], fill=YELLOW, font=DF, align=LFTT)
    sc(ws5, row, 7, c["high"], fill=RED, font=DF, align=LFTT)
    sc(ws5, row, 8, c["dispute"], fill=WHITE_FILL, font=DF, align=LFTT)
    # Risk level extraction
    risk_text = c["risk"]
    risk_lvl = "HIGH" if "HIGH" in risk_text else "MID" if "MID" in risk_text else "LOW"
    sc(ws5, row, 9, c["risk"], fill=rfill(risk_lvl), font=BF, align=CTR)
    sc(ws5, row, 10, c["confidence"], fill=WHITE_FILL, font=DF, align=CTR)
for ci, w in enumerate([8,22,18,22,30,30,30,35,8,8], 1): ws5.column_dimensions[get_column_letter(ci)].width = w
ws5.freeze_panes = "A3"

# === Sheet 6: Amendment Priorities ===
ws6 = wb.create_sheet("AmendmentPriorities")
h6 = ["Priority","Clause","Issue","Proposed Text (Full Replacement)","Rationale","Negotiation Strategy"]
ws6.merge_cells("A1:F1"); sc(ws6, 1, 1, "Amendment Priorities - Ready for Redlining", fill=SECTION_FILL, font=SF, align=LFT)
for ci, h in enumerate(h6, 1): sc(ws6, 2, ci, h, fill=COLHDR_FILL, font=CHDR_F, align=CTR)
for ri, a in enumerate(AM):
    row = ri+3
    tier = a["priority"][:2]
    tf = {"T1":RED,"T2":YELLOW,"T3":GREEN}.get(tier, WHITE_FILL)
    sc(ws6, row, 1, a["priority"], fill=tf, font=BF, align=CTR)
    sc(ws6, row, 2, a["clause"], fill=WHITE_FILL, font=DF, align=CTR)
    sc(ws6, row, 3, a["problem"], fill=WHITE_FILL, font=BF, align=LFTT)
    sc(ws6, row, 4, a["text"], fill=BLUE_HINT, font=DF, align=LFTT)
    sc(ws6, row, 5, a["reason"], fill=WHITE_FILL, font=DF, align=LFTT)
    sc(ws6, row, 6, a["strategy"], fill=WHITE_FILL, font=DF, align=LFTT)
for ci, w in enumerate([8,15,22,50,28,35], 1): ws6.column_dimensions[get_column_letter(ci)].width = w
ws6.freeze_panes = "A3"

# === Sheet 7: Key Control Questions ===
ws7 = wb.create_sheet("ControlQuestions")
ws7.merge_cells("A1:C1"); sc(ws7, 1, 1, "Key Control Questions - 5 Questions", fill=SECTION_FILL, font=SF, align=LFT)
for ci, h in enumerate(["#","Question","Answer"], 1): sc(ws7, 2, ci, h, fill=COLHDR_FILL, font=CHDR_F, align=CTR)
for ri, q in enumerate(Q):
    row = ri+3
    sc(ws7, row, 1, q["id"], fill=WHITE_FILL, font=BF, align=CTR)
    sc(ws7, row, 2, q["q"], fill=WHITE_FILL, font=LF, align=LFTT)
    sc(ws7, row, 3, q["a"], fill=WHITE_FILL, font=DF, align=LFTT)
for ci, w in enumerate([8,35,80], 1): ws7.column_dimensions[get_column_letter(ci)].width = w
ws7.freeze_panes = "A3"

# -- Save --
OUT = os.path.join(config.BASE, "202401 代持协议_曹(OCR)(1)脱敏版_法律审查报告.xlsx")
wb.save(OUT)
print(f"Done: {OUT}")
print(f"Sheets ({len(wb.sheetnames)}): {wb.sheetnames}")
