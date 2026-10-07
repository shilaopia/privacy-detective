# -*- coding: utf-8 -*-
"""新旧版本审查结果对比报告生成器。
用法: python -X utf8 gen_diff_report.py <key> <旧版标签> <新版标签> [新key后缀]
例: python -X utf8 gen_diff_report.py 扣子Coze_ToC 2024-05-10版 2026-08-10版
读 review_json/{key}.json(旧) vs review_json/{key}_2026.json(新)
输出 review_json/diff_{key}_旧版vs新版.md
"""
import json, os, sys

BASE = os.path.dirname(os.path.abspath(__file__))
RANK = {"合规": 0, "存疑": 1, "不合规": 2, "未提及": 3}


def main():
    key, old_label, new_label = sys.argv[1], sys.argv[2], sys.argv[3]
    new_key = sys.argv[4] if len(sys.argv) > 4 else key + "_2026"
    old = json.load(open(os.path.join(BASE, "review_json", key + ".json"), encoding="utf-8"))
    new = json.load(open(os.path.join(BASE, "review_json", new_key + ".json"), encoding="utf-8"))
    baseline = json.load(open(os.path.join(BASE, "baseline_v7.json"), encoding="utf-8"))
    clauses = [c["clause"] for c in baseline]

    changed, same = [], 0
    cnt_old = {v: 0 for v in RANK}
    cnt_new = {v: 0 for v in RANK}
    for c in clauses:
        vo, vn = old[c]["评估"], new[c]["评估"]
        cnt_old[vo] += 1
        cnt_new[vn] += 1
        if vo != vn:
            changed.append((c, vo, vn, old[c].get("合规说明", ""), new[c].get("合规说明", "")))
        else:
            same += 1

    lines = [
        f"# {key} 隐私政策新旧版本审查对比",
        "",
        f"- 旧版：{old_label}（审查文件 review_json/{key}.json）",
        f"- 新版：{new_label}（审查文件 review_json/{new_key}.json）",
        f"- 评估变化条款：**{len(changed)} / 58**；评估不变：{same}",
        "",
        "## 四项计数对比",
        "",
        "| 版本 | 合规 | 存疑 | 不合规 | 未提及 |",
        "|---|---|---|---|---|",
        f"| 旧版 | {cnt_old['合规']} | {cnt_old['存疑']} | {cnt_old['不合规']} | {cnt_old['未提及']} |",
        f"| 新版 | {cnt_new['合规']} | {cnt_new['存疑']} | {cnt_new['不合规']} | {cnt_new['未提及']} |",
        "",
        "## 评估变化条款明细",
        "",
    ]
    if not changed:
        lines.append("（无评估变化条款）")
    for c, vo, vn, so, sn in changed:
        lines += [
            f"### {c}：{vo} → {vn}",
            "",
            f"- 旧版说明：{so}",
            f"- 新版说明：{sn}",
            "",
        ]
    lines += [
        "## 58条全量对照表",
        "",
        "| 条款 | 旧版评估 | 新版评估 |",
        "|---|---|---|",
    ]
    for c in clauses:
        vo, vn = old[c]["评估"], new[c]["评估"]
        lines.append(f"| {c} | {vo} | {vn}{' **变化**' if vo != vn else ''} |")

    out = os.path.join(BASE, "review_json", f"diff_{key}_旧版vs新版.md")
    open(out, "w", encoding="utf-8").write("\n".join(lines))
    print(f"OK {out}: 变化 {len(changed)}/58")


if __name__ == "__main__":
    main()
