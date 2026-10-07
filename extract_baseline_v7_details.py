# -*- coding: utf-8 -*-
"""将制度标尺 v7.0 Excel 按义务细则粒度结构化 → baseline_v7_details.json。

每细则: {id, clause, dim1, dim2, text, source, gai}
- id: 细则序号(1..449, 按原表行序)
- gai: 是否 GAI 特有(出处含《生成式人工智能服务管理暂行办法》或文本含 GAI 关键词)
合并单元格向下填充; 二级维度列与条款编号前缀不一致时置空(同 extract_baseline_v7.py)。
"""
import json
import openpyxl
import os
import config

SRC = os.path.join(config.BASE, "制度标尺-隐私合规-v7.0_副本.xlsx")
OUT = os.path.join(config.BASE, "baseline_v7_details.json")

GAI_SRC_KEY = "生成式人工智能服务管理暂行办法"
GAI_TEXT_KEYS = ("训练数据", "生成式", "合成内容", "深度合成", "模型训练", "生成内容")


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb.active

    details = []
    d1 = d2 = d3 = None
    for a, b, c, d, e in ws.iter_rows(min_row=2, values_only=True):
        if a is not None:
            d1 = str(a).strip()
        if b is not None:
            d2 = str(b).strip()
        if c is not None:
            d3 = str(c).strip()
        if d3 is None or d is None:
            continue
        prefix = d3.split("-")[0]
        dim2 = d2 if d2 and str(d2).startswith(prefix + "-") else None
        text = str(d).strip()
        source = str(e).strip() if e else ""
        gai = (GAI_SRC_KEY in source) or any(k in text for k in GAI_TEXT_KEYS)
        details.append({
            "id": len(details) + 1,
            "clause": d3,
            "dim1": d1,
            "dim2": dim2,
            "text": text,
            "source": source,
            "gai": gai,
        })

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(details, f, ensure_ascii=False, indent=1)
    clauses = {d["clause"] for d in details}
    gai_n = sum(1 for d in details if d["gai"])
    nosrc = [d["id"] for d in details if not d["source"]]
    print(f"细则数: {len(details)}  条款数: {len(clauses)}  GAI特有: {gai_n}  非GAI: {len(details)-gai_n}")
    print("出处为空的细则:", nosrc if nosrc else "无")


if __name__ == "__main__":
    main()
