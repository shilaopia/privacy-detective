# -*- coding: utf-8 -*-
"""导出分片任务单: python export_shards.py
把 baseline_v7.json 按 7 个一级维度拆成 review_json/shards/p{1..7}_clauses.json，
供审查子 Agent 直接读取（避免每个 Agent 全量解析 87KB 基准）。
"""
import json, os

BASE = os.path.dirname(os.path.abspath(__file__))
baseline = json.load(open(os.path.join(BASE, "baseline_v7.json"), encoding="utf-8"))
outdir = os.path.join(BASE, "review_json", "shards")
os.makedirs(outdir, exist_ok=True)
dims = {}
for c in baseline:
    n = int(c["dim1"].split("-")[0])
    dims.setdefault(n, []).append(c)
for n, clauses in sorted(dims.items()):
    slim = [{"clause": c["clause"], "dim1": c["dim1"], "dim2": c["dim2"],
             "details": c["details"]} for c in clauses]
    p = os.path.join(outdir, f"p{n}_clauses.json")
    json.dump(slim, open(p, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"p{n}: {len(clauses)} 条 -> {p}")
