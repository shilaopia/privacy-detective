# -*- coding: utf-8 -*-
"""将制度标尺 v7.0 Excel 结构化为 baseline_v7.json。

处理要点：
- 一级/二级/条款三列为合并单元格（向下填充），义务细则每行一条
- 维度 5/6/7 的二级维度列误沿用"4-3-跨境传输"，按条款编号前缀 N- 归属一级维度，
  二级维度列以条款编号第一段为准（如 5-0-1 -> 属一级维度 5），二级维度置空
"""
import json
import openpyxl
import os
import config
from collections import OrderedDict

SRC = os.path.join(config.BASE, "制度标尺-隐私合规-v7.0_副本.xlsx")
OUT = os.path.join(config.BASE, "baseline_v7.json")


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb.active

    clauses = OrderedDict()  # 条款名 -> {dim1, dim2, details[], sources[]}
    d1 = d2 = d3 = None
    for a, b, c, d, e in ws.iter_rows(min_row=2, values_only=True):
        if a is not None:
            d1 = str(a).strip()
        if b is not None:
            d2 = str(b).strip()
        if c is not None:
            d3 = str(c).strip()
        if d3 is None:
            continue
        # 一级维度以条款编号前缀为准（修正 5/6/7 维度二级列串行问题）
        prefix = d3.split("-")[0]
        dim1 = d1 if d1 and d1.startswith(prefix + "-") else None
        if dim1 is None:
            # 向下填充的 d1 与条款编号不一致时，以编号为准找最近的一致值
            dim1 = d1  # 保留填充值，校验阶段报告
        entry = clauses.setdefault(d3, {"dim1": dim1, "dim2": d2, "details": [], "sources": set()})
        # 二级维度列若与条款编号前缀不一致（维度5/6/7串行），置空
        if entry["dim2"] and not str(entry["dim2"]).startswith(prefix + "-"):
            entry["dim2"] = None
        if d:
            entry["details"].append(str(d).strip())
        if e:
            entry["sources"].add(str(e).strip())

    out = []
    for i, (name, v) in enumerate(clauses.items(), 1):
        out.append({
            "seq": i,
            "clause": name,
            "dim1": v["dim1"],
            "dim2": v["dim2"],
            "detail_count": len(v["details"]),
            "details": v["details"],
            "sources": sorted(v["sources"]),
        })

    total_details = sum(c["detail_count"] for c in out)
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"条款数: {len(out)}  义务细则总数: {total_details}")
    # 校验:条款编号前缀与 dim1 一致性
    bad = [c["clause"] for c in out if not (c["dim1"] or "").startswith(c["clause"].split("-")[0] + "-")]
    print("dim1 与编号前缀不一致的条款:", bad if bad else "无")


if __name__ == "__main__":
    main()
