#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正式整合腳本 — 一次登入,依序跑完全部報表代號
================================================

把report_idlr17.py / report_idlr16.py / report_idlr68.py / report_mocr10.py
/ report_invr18.py(6組) / report_idlr69.py(2組)串在一起:只登入一次,
依序呼叫每支報表模組的run(main_win, out_lines)。

**設計原則(照使用者要求)**:每支報表各自獨立成一個檔案,這裡不重寫任何一支
的邏輯,只是照順序呼叫。哪一支報表出問題,直接單獨修改、單獨測試那支檔案
(python report_xxx.py --dry-run)就好,不用動這支整合腳本或其他報表。

執行順序刻意把最穩、最簡單的排前面(IDLR17),風險較高/較新的排後面
(IDLR69),萬一半路出錯,已經成功送出的報表不受影響——單一報表失敗不會
中斷整個流程,會記錄錯誤、繼續跑下一支,最後印出總結。

用法
----
    python run_all_reports.py                 # 全部真的送出
    python run_all_reports.py --dry-run        # 全部只填欄位、按取消,不送出
    python run_all_reports.py --skip IDLR16    # 跳過指定代號(可重複指定多次)
    python run_all_reports.py --only IDLR68 MOCR10   # 只跑指定的代號
"""

import argparse
import sys

from part2_inspect import log, login

import report_idlr16
import report_idlr17
import report_idlr68
import report_idlr69
import report_invr18
import report_mocr10

# 執行順序:已驗證程度由高到低、由簡單到複雜排列
REPORT_MODULES = [
    ("IDLR17", report_idlr17),
    ("IDLR16", report_idlr16),
    ("IDLR68", report_idlr68),
    ("MOCR10", report_mocr10),
    ("INVR18", report_invr18),
    ("IDLR69", report_idlr69),
]


def main():
    parser = argparse.ArgumentParser(description="正式整合腳本 — 依序跑完全部報表代號")
    parser.add_argument("--dry-run", action="store_true",
                         help="全部報表都只填欄位、按取消,不會真的送出查詢")
    parser.add_argument("--skip", action="append", default=[],
                         help="跳過指定的程式代號(可重複指定,例如--skip IDLR16 --skip IDLR69)")
    parser.add_argument("--only", nargs="+", default=None,
                         help="只跑指定的程式代號清單,例如--only IDLR68 MOCR10")
    args = parser.parse_args()

    skip = set(args.skip)
    only = set(args.only) if args.only else None

    out_lines = []
    summary = []
    try:
        app, main_win = login(out_lines)

        for code, module in REPORT_MODULES:
            if code in skip:
                log(f"\n[跳過] {code}(--skip指定)")
                continue
            if only is not None and code not in only:
                log(f"\n[跳過] {code}(不在--only清單裡)")
                continue

            log(f"\n{'=' * 70}\n開始跑: {code}\n{'=' * 70}")
            try:
                module.run(main_win, out_lines, dry_run=args.dry_run)
                summary.append((code, "成功"))
                log(f"[完成] {code}")
            except Exception as e:
                summary.append((code, f"失敗: {e!r}"))
                out_lines.append(f"[錯誤] {code} 失敗: {e!r}")
                log(f"[錯誤] {code} 失敗: {e!r}(繼續跑下一支報表)")

    except Exception as e:
        out_lines.append("")
        out_lines.append("!" * 70)
        out_lines.append(f"登入或啟動流程本身發生錯誤(還沒開始跑任何報表): {e!r}")
        out_lines.append("!" * 70)
        log(f"[致命錯誤] {e!r}")
        raise
    finally:
        out_lines.append("\n" + "=" * 70)
        out_lines.append("總結:")
        for code, result in summary:
            out_lines.append(f"  {code}: {result}")
        log("\n總結:")
        for code, result in summary:
            log(f"  {code}: {result}")

        with open("run_all_reports_output.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(out_lines))
        log("\n記錄已寫入: run_all_reports_output.txt")


if __name__ == "__main__":
    sys.exit(main())
