# -*- coding: utf-8 -*-
"""结项书第四部分写作用数据全量核算：从 review_json 重算所有正文引用数字。
输出：层/子类统计、58条合规率排名、7维度×3层交叉、模型层ToC/ToB对照。
"""
import json, os, re, collections

BASE = os.path.dirname(os.path.abspath(__file__))
src = open(os.path.join(BASE, "gen_three_layer_review.py"), encoding="utf-8").read()
LAYERS = eval(re.search(r"LAYERS\s*=\s*(\[.*?\n\])", src, re.S).group(1))

verdicts, layer_of, sub_of, name_of = {}, {}, {}, {}
for layer, subs in LAYERS:
    for sub, items in subs:
        for key, name, end in items:
            d = json.load(open(os.path.join(BASE, "review_json", key + ".json"), encoding="utf-8"))
            verdicts[key] = {k: v["评估"] for k, v in d.items()}
            layer_of[key], sub_of[key], name_of[key] = layer, sub, name

baseline = json.load(open(os.path.join(BASE, "baseline_v7.json"), encoding="utf-8"))
clauses = [c["clause"] for c in baseline]
V = ["合规", "存疑", "不合规", "未提及"]

def stats(keys):
    cnt = collections.Counter()
    for k in keys:
        for c in clauses:
            cnt[verdicts[k][c]] += 1
    tot = len(keys) * 58
    return cnt, tot

print("== 层统计 ==")
for l in ["模型层", "分发层", "应用层"]:
    ks = [k for k in verdicts if layer_of[k] == l]
    cnt, tot = stats(ks)
    print(l, len(ks), "家", {v: cnt[v] for v in V}, "合规率 %.1f%%" % (100 * cnt["合规"] / tot))

print("\n== 子类统计 ==")
for layer, subs in LAYERS:
    for sub, items in subs:
        ks = [k for k, _, _ in items]
        cnt, tot = stats(ks)
        print(f"{layer}|{sub}|{len(ks)}家|合规{cnt['合规']}|存疑{cnt['存疑']}|不合规{cnt['不合规']}|未提及{cnt['未提及']}|合规率{100*cnt['合规']/tot:.1f}%")

print("\n== 58条合规率（升序，全样本64家） ==")
rates = []
for c in clauses:
    r = 100 * sum(1 for k in verdicts if verdicts[k][c] == "合规") / 64
    nm = sum(1 for k in verdicts if verdicts[k][c] == "未提及")
    ng = sum(1 for k in verdicts if verdicts[k][c] == "不合规")
    rates.append((r, c, nm, ng))
for r, c, nm, ng in sorted(rates):
    print(f"{r:5.1f}%  {c}  (未提及{nm} 不合规{ng})")

print("\n== 7维度 × 3层 合规率矩阵 ==")
for dnum in range(1, 8):
    dcl = [c for c in clauses if c.startswith(f"{dnum}-")]
    row = []
    for l in ["模型层", "分发层", "应用层"]:
        ks = [k for k in verdicts if layer_of[k] == l]
        r = 100 * sum(1 for k in ks for c in dcl if verdicts[k][c] == "合规") / (len(ks) * len(dcl))
        row.append(round(r, 1))
    print(f"维度{dnum}({len(dcl)}条): 模型{row[0]} 分发{row[1]} 应用{row[2]}")

print("\n== 模型层 ToC/ToB 对照 ==")
ml = [(k, name_of[k]) for k in verdicts if layer_of[k] == "模型层"]
vendors = collections.defaultdict(dict)
for k, nm in ml:
    v = nm.replace("_ToC", "").replace("_ToB", "")
    end = "ToC" if k.endswith("ToC") else "ToB"
    r = 100 * sum(1 for c in clauses if verdicts[k][c] == "合规") / 58
    vendors[v][end] = round(r, 1)
for v, d in vendors.items():
    print(v, d)
