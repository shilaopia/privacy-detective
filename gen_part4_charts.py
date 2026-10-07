# -*- coding: utf-8 -*-
"""结项书第四部分配套图修订版：重绘 3 张瑕疵图。
图4-3 三层判定结构堆叠图（修Y轴叠字）
图4-7 七维度合规率均值图（修均值标注遮挡）
图4-9 跨境4-3系列分层图（修X轴标签乱码）
数据源：review_json/*.json（与终表一致）
输出：BASE/pictures/大创合规图/修订版/
"""
import json, os, re, collections
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "pictures", "大创合规图", "修订版")
os.makedirs(OUT, exist_ok=True)

# 调色板（对齐原图）
C_LAYER = {"应用层": "#2176ae", "模型层": "#5aa9e6", "分发层": "#c77dab"}
C_VERDICT = {"合规": "#16a085", "存疑": "#f39c12", "不合规": "#d35400", "未提及": "#7f7f7f"}

# ---- 读 LAYERS 并加载全部判定 ----
src = open(os.path.join(BASE, "gen_three_layer_review.py"), encoding="utf-8").read()
LAYERS = eval(re.search(r"LAYERS\s*=\s*(\[.*?\n\])", src, re.S).group(1))
verdicts = {}  # key -> {clause: 评估}
layer_of = {}
for layer, subs in LAYERS:
    for sub, items in subs:
        for key, name, end in items:
            d = json.load(open(os.path.join(BASE, "review_json", key + ".json"), encoding="utf-8"))
            verdicts[key] = {k: v["评估"] for k, v in d.items()}
            layer_of[key] = layer
baseline = json.load(open(os.path.join(BASE, "baseline_v7.json"), encoding="utf-8"))
clauses = [c["clause"] for c in baseline]

LAYER_ORDER = ["应用层", "模型层", "分发层"]
layer_keys = {l: [k for k in verdicts if layer_of[k] == l] for l in LAYER_ORDER}

# ---- 图4-3：三层判定结构堆叠图 ----
fig, ax = plt.subplots(figsize=(12, 6.3), dpi=150)
layers_rev = ["分发层", "模型层", "应用层"]  # 原图自上而下顺序
ypos = np.arange(len(layers_rev))[::-1]
left = np.zeros(len(layers_rev))
for v in ["合规", "存疑", "不合规", "未提及"]:
    vals = []
    for l in layers_rev:
        ks = layer_keys[l]
        cnt = sum(1 for k in ks for c in clauses if verdicts[k][c] == v)
        vals.append(100 * cnt / (len(ks) * 58))
    ax.barh(ypos, vals, left=left, color=C_VERDICT[v], height=0.62, label=v)
    for i, (y, val, lft) in enumerate(zip(ypos, vals, left)):
        if val > 2.5:
            ax.text(lft + val / 2, y, f"{val:.1f}%", ha="center", va="center",
                    color="white", fontsize=13, fontweight="bold")
    left += np.array(vals)
ax.set_yticks(ypos)
ax.set_yticklabels(layers_rev, fontsize=15)
ax.set_xlim(0, 100)
ax.set_xlabel("占比（%）", fontsize=13)
ax.set_title("三层平台合规判定结构", fontsize=18, pad=14)
ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.10), ncol=4, frameon=False, fontsize=13)
ax.spines[["top", "right"]].set_visible(False)
fig.tight_layout()
fig.savefig(os.path.join(OUT, "图4-3_三层判定结构堆叠图.png"), bbox_inches="tight")
plt.close(fig)

# ---- 图4-7：七维度合规率均值图 ----
DIM_SUBS = {  # 与原图一致的分组小字
    "1": ["告知同意", "敏感个人信息"],
    "2": ["一般权利", "退出权利"],
    "3": ["安全措施", "去标识化与匿名化"],
    "4": ["境内共享与转移", "API与委托处理", "跨境传输"],
}
DIM_NAMES = ["信息收集与授权", "用户权利", "存储与安全", "共享与转移",
             "自动化决策", "内设监督机构", "格式与显著标识"]
dim_names, dim_rates, dim_ns = [], [], []
overall_num = overall_den = 0
for dnum in range(1, 8):
    dcl = [c for c in clauses if c.startswith(f"{dnum}-")]
    cnt = sum(1 for k in verdicts for c in dcl if verdicts[k][c] == "合规")
    den = len(verdicts) * len(dcl)
    dim_names.append(DIM_NAMES[dnum - 1])
    dim_rates.append(100 * cnt / den)
    dim_ns.append(len(dcl))
    overall_num += cnt
    overall_den += den
overall = 100 * overall_num / overall_den

fig, ax = plt.subplots(figsize=(16.5, 7.8), dpi=150)
x = np.arange(7)
bars = ax.bar(x, dim_rates, width=0.58, color="#1f6fb4")
for xi, r, n in zip(x, dim_rates, dim_ns):
    ax.text(xi, r + 1.2, f"{r:.1f}%（n={n}）", ha="center", fontsize=13, fontweight="bold")
ax.axhline(overall, ls="--", color="#888888", lw=1.6)
ax.set_xlim(-0.6, 7.5)
ax.text(7.42, overall + 1.0, f"总体均值 {overall:.1f}%", fontsize=13, color="#555555",
        ha="right", va="bottom")
ax.set_xticks(x)
labels = []
for i, name in enumerate(dim_names):
    subs = "\n".join(DIM_SUBS.get(str(i + 1), []))
    labels.append(name + ("\n" + subs if subs else ""))
ax.set_xticklabels(labels, fontsize=12.5)
ax.set_ylabel("合规率（%）", fontsize=13)
ax.set_ylim(0, max(dim_rates) * 1.18)
ax.spines[["top", "right"]].set_visible(False)
fig.tight_layout()
fig.savefig(os.path.join(OUT, "图4-7_七维度合规率均值图.png"), bbox_inches="tight")
plt.close(fig)

# ---- 图4-9：跨境4-3系列分层合规率 ----
cb_clauses = [c for c in clauses if c.startswith("4-3-")]
short = {
    "4-3-1-个人信息保护认证": "4-3-1\n保护认证",
    "4-3-2-个人信息出境标准合同": "4-3-2\n出境标准合同",
    "4-3-3-境外接收方责任和资质": "4-3-3\n境外接收方责任",
    "4-3-4-数据出境安全评估": "4-3-4\n出境安全评估",
    "4-3-6-跨境传输一般合规要求": "4-3-6\n跨境一般合规",
    "4-3-7-跨境提供个人信息的单独同意": "4-3-7\n跨境单独同意",
}
fig, ax = plt.subplots(figsize=(11.5, 6.2), dpi=150)
ngroups = len(cb_clauses)
x = np.arange(ngroups)
w = 0.26
for i, l in enumerate(LAYER_ORDER):
    ks = layer_keys[l]
    vals = [100 * sum(1 for k in ks if verdicts[k][c] == "合规") / len(ks) for c in cb_clauses]
    ax.bar(x + (i - 1) * w, vals, width=w, color=C_LAYER[l], label=l)
ax.set_xticks(x)
ax.set_xticklabels([short[c] for c in cb_clauses], fontsize=11.5)
ax.set_ylabel("合规率（%）", fontsize=13)
ax.set_title("跨境传输条款合规率（4-3 系列）", fontsize=16, pad=12)
ax.legend(fontsize=12, frameon=False)
ax.spines[["top", "right"]].set_visible(False)
fig.tight_layout()
fig.savefig(os.path.join(OUT, "图4-9_跨境4-3系列分层合规率.png"), bbox_inches="tight")
plt.close(fig)

# ---- 打印核对数据 ----
print("总体合规率 %.1f%%" % overall)
for l in LAYER_ORDER:
    ks = layer_keys[l]
    print(l, len(ks), "家")
print("七维度:", [(n, round(r, 1)) for n, r in zip(dim_names, dim_rates)])
print("跨境4-3:", [short[c].replace(chr(10), " ") for c in cb_clauses])
for l in LAYER_ORDER:
    ks = layer_keys[l]
    print(" ", l, [round(100 * sum(1 for k in ks if verdicts[k][c] == "合规") / len(ks), 1) for c in cb_clauses])
print("输出目录:", OUT)
