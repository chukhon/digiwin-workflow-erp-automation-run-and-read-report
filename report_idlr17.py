#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正式抓資料腳本 — IDLR17(訂單預計出貨明細表)
==============================================

狀態:已完整驗證,全部欄位留白/預設,直接送出。這是所有程式代號裡最簡單的
一支,也是最早驗證「登入→切換檢視類別→跳轉代號→按確認→關閉訊息視窗」整條
主流程的那一支。

用法
----
    python report_idlr17.py              # 登入+送出(真的按確認,會產生真實工作)
    python report_idlr17.py --dry-run    # 填完(其實沒有欄位要填)之後按取消,不送出

也可以被run_all_reports.py當模組import,呼叫run(main_win, out_lines)。
"""

import argparse
import sys

from part2_inspect import (
    find_button,
    jump_to_code,
    log,
    login,
    save_screenshot,
    submit_query,
)

CODE = "IDLR17"


def run(main_win, out_lines, dry_run=False):
    query_win = jump_to_code(main_win, CODE, out_lines)
    if query_win is None:
        raise RuntimeError(f"沒有拿到{CODE}查詢畫面。")

    log(f"\n===== {CODE}: 全部欄位留白/預設,不需要任何處理 =====")
    save_screenshot(query_win, f"{CODE.lower()}_before_submit.png", out_lines)

    if dry_run:
        log(f"  --dry-run: 不送出,按取消關閉{CODE}查詢畫面。")
        cancel_btn = find_button(query_win, "取消")
        if cancel_btn is not None:
            cancel_btn.click()
        return None

    return submit_query(query_win, CODE, out_lines)


def main():
    parser = argparse.ArgumentParser(description=f"正式抓資料腳本 — {CODE}")
    parser.add_argument("--dry-run", action="store_true",
                         help="只跑到填完欄位、按取消關閉,不會真的送出查詢")
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


if __name__ == "__main__":
    sys.exit(main())
