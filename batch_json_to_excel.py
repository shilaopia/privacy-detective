# -*- coding: utf-8 -*-
"""
batch_json_to_excel.py — 扫描 _tmp_gb_render/ 下所有 JSON，对应找不到 xlsx 的就生成。

用法：
    python batch_json_to_excel.py            # 增量补转
    python batch_json_to_excel.py --force    # 全部重转
"""
from __future__ import annotations

import config
import argparse
import json
import sys
import traceback
from pathlib import Path

from gen_gb_batch import write_excel

PROJ_ROOT = Path(config.BASE)
JSON_DIR  = PROJ_ROOT / "_tmp_gb_render"
OUT_DIR   = PROJ_ROOT / "National Standard" / "re-nationalstandard" / "re_national_standard_excels"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--force", action="store_true")
    args = ap.parse_args()

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    jsons = sorted(JSON_DIR.glob("*.json"))
    print(f"Found {len(jsons)} JSON files in {JSON_DIR}")

    ok, skipped, failed = 0, 0, 0
    for jp in jsons:
        stem = jp.stem
        xlsx = OUT_DIR / f"{stem}.xlsx"
        if xlsx.exists() and not args.force:
            print(f"  SKIP {stem} (xlsx exists)")
            skipped += 1
            continue
        try:
            data = json.loads(jp.read_text(encoding="utf-8"))
            write_excel(data, xlsx, f"{stem}.pdf")
            n = len(data.get("clauses", []))
            print(f"  OK   {stem}.xlsx ({n} clauses)")
            ok += 1
        except Exception as e:
            print(f"  FAIL {stem}: {e}")
            traceback.print_exc()
            failed += 1

    print(f"\n=== ok={ok} skip={skipped} fail={failed} ===")


if __name__ == "__main__":
    main()
