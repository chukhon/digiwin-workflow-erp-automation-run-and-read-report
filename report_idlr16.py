#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正式抓資料腳本 — IDLR16(Backlog)
==================================

**注意:這支跟IDLR17不一樣,還沒有真的用pywinauto單獨實測過。**
`erp-report-selection-methods-2026-09-03.md`裡的結論(全部欄位留白/預設,
不需要任何選擇)是根據錄影分析+跟IDLR17/68同一段錄影前後文推測的,不是
實測結果(錄影裡這支查詢畫面載入太快,沒有精準截到)。

這支腳本先照這個假設寫(跟report_idlr17.py同一套模板:什麼都不填,直接
送出)。**強烈建議第一次先加`--dry-run`跑一次**,實際看一次IDLR16查詢畫面
的截圖,確認真的沒有必填欄位、沒有意外的預設值,再拿掉`--dry-run`正式送出。

用法
----
    python report_idlr16.py --dry-run    # 建議第一次先這樣跑,只截圖不送出
    python report_idlr16.py              # 確認沒問題後,真的送出
"""

import argparse
import sys

from part2_inspect import (
    dump_control_tree,
    find_button,
    jump_to_code,
    log,
    login,
    save_screenshot,
    submit_query,
)

CODE = "IDLR16"


def run(main_win, out_lines, dry_run=False):
    query_win = jump_to_code(main_win, CODE, out_lines)
    if query_win is None:
        raise RuntimeError(f"沒有拿到{CODE}查詢畫面。")

    log(f"\n===== {CODE}: 假設全部欄位留白/預設(未實測,見腳本開頭說明) =====")
    save_screenshot(query_win, f"{CODE.lower()}_before_submit.png", out_lines)
    dump_control_tree(query_win, out_lines, f"{CODE}查詢畫面(送出之前)")

    if dry_run:
        log(f"  --dry-run: 不送出,按取消關閉{CODE}查詢畫面。")
        cancel_btn = find_button(query_win, "取消")
        if cancel_btn is not None:
            cancel_btn.click()
        return None

    return submit_query(query_win, CODE, out_lines)


def main():
    parser = argparse.ArgumentParser(description=f"正式抓資料腳本 — {CODE}(未實測,建議先--dry-run)")
    parser.add_argument("--dry-run", action="store_true",
                         help="只跑到填完欄位、截圖+dump控制項、按取消關閉,不會真的送出查詢")
    args = parser.parse_args()

    out_lines = []
    try:
        app, main_win = login(out_lines)
        run(main_win, out_lines, dry_run=args.dry_run)
    except Exception as e:
        out_lines.append("")
        out_lines.append("!" * 70)
        out_lines.append(f"執行中發生錯誤: {e!r}")
        out_lines.append("!" * 70)
        log(f"[錯誤] {e!r}")
        raise
    finally:
        with open(f"{CODE.lower()}_run_output.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(out_lines))
        log(f"\n記錄已寫入: {CODE.lower()}_run_output.txt")
        log("這支還沒實測過,請務必檢查截圖+輸出檔案,確認跟IDLR17一樣安全。")


if __name__ == "__main__":
    sys.exit(main())
