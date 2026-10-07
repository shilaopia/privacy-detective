# -*- coding: utf-8 -*-
"""
json_to_excel.py — 将 SKILL 框架提炼出的 JSON 写为标准化 Excel。

用法：
    python json_to_excel.py <input.json> <pdf_filename> <output.xlsx>

输入 JSON 格式参见 gen_gb_batch.py 的 JSON_INSTRUCTION。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from gen_gb_batch import write_excel


def main():
    if len(sys.argv) != 4:
        print("Usage: python json_to_excel.py <input.json> <pdf_filename> <output.xlsx>")
        sys.exit(2)

    json_path = Path(sys.argv[1])
    pdf_filename = sys.argv[2]
    out_path = Path(sys.argv[3])

    data = json.loads(json_path.read_text(encoding="utf-8"))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    write_excel(data, out_path, pdf_filename)
    print(f"OK: wrote {out_path} ({len(data.get('clauses', []))} clauses)")


if __name__ == "__main__":
    main()
