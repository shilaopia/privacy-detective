# -*- coding: utf-8 -*-
"""合并分片审查结果: python merge_parts.py <key1> [key2 ...]
读 review_json/parts/{key}__p{1..7}.json -> review_json/{key}.json
校验: 7 分片齐全、58 条款无缺漏/多余键、评估值合法。
"""
import json, os, sys

BASE = os.path.dirname(os.path.abspath(__file__))
VALID = {"合规", "存疑", "不合规", "未提及"}

def merge(key):
    parts_dir = os.path.join(BASE, "review_json", "parts")
    baseline = json.load(open(os.path.join(BASE, "baseline_v7.json"), encoding="utf-8"))
    missing = [n for n in range(1, 8)
               if not os.path.exists(os.path.join(parts_dir, f"{key}__p{n}.json"))]
    if missing:
        return f"{key}: 缺分片 p{missing}"
    merged = {}
    bad_val = []
    for n in range(1, 8):
        part = json.load(open(os.path.join(parts_dir, f"{key}__p{n}.json"), encoding="utf-8"))
        for k, v in part.items():
            if not isinstance(v, dict) or v.get("评估") not in VALID:
                bad_val.append((n, k))
            merged[k] = v
    if bad_val:
        return f"{key}: 评估值非法 {bad_val}"
    lack = [c["clause"] for c in baseline if c["clause"] not in merged]
    extra = [k for k in merged if k not in {c["clause"] for c in baseline}]
    if lack or extra:
        return f"{key}: 缺条款{lack} 多余键{extra}"
    out = os.path.join(BASE, "review_json", key + ".json")
    json.dump(merged, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return f"OK {key}: 58条 -> review_json/{key}.json"

if __name__ == "__main__":
    fails = 0
    for key in sys.argv[1:]:
        r = merge(key)
        print(r)
        if not r.startswith("OK"):
            fails += 1
    sys.exit(1 if fails else 0)
