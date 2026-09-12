#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正式抓資料腳本 — MOCR10(製令明細表)
=======================================

狀態:「選擇製令狀態」5個checkbox的目標狀態已完整驗證成功(2026-09-03,
使用者確認MOCR10整體已完成)。這支腳本在驗證過的5個checkbox之外,多加了
1個先前沒有任何腳本處理過的獨立checkbox「是否列印材料明細」(目標:不勾)
——這一個是新加的,用的是同一套已驗證的截圖比對顏色方法(is_checkbox_checked
/set_checkbox_checked對同一種class TDSUCBitCheckBox.UnicodeClass都適用),
風險低,但**第一次執行請務必核對截圖**,確認這個checkbox真的被正確設定。

「選擇製令性質」「選擇確認狀態」這兩個下拉,設定表要求的值(全部/已確認)
剛好就是ERP畫面預設值,這支腳本**刻意不去動它們**——目前沒有驗證過的方式
可以在不猜座標/猜順序的情況下,可靠辨識出這兩個下拉分別是哪一個(它們沒有
獨立文字標籤可比對),與其冒險猜錯,不如維持預設(反正預設值本來就對)。
如果之後想要腳本主動明確選一次,需要先跑一次偵察截圖確認這兩個ComboBox的
正確辨識方式,再回來加這段邏輯。

「選擇製令編號」「選擇母製令編號」「選擇確認日期」這3個SelectPanel維持
留白/不動。

用法
----
    python report_mocr10.py              # 真的送出
    python report_mocr10.py --dry-run    # 填完欄位、按取消,不送出
"""

import argparse
import sys

from part2_inspect import (
    find_button,
    find_checkbox_by_text,
    jump_to_code,
    log,
    login,
    save_screenshot,
    set_checkbox_checked,
    submit_query,
)

CODE = "MOCR10"

# 2026-09-03 使用者核對實際截圖後確認的目標狀態(已驗證成功)
TARGET_STATES = {
    "未開工": True,
    "已發料": True,
    "生產中": True,
    "已完工": False,
    "指定完工": False,
    # 新加的獨立checkbox(不屬於「選擇製令狀態」GroupBox),同一種class,
    # 用同一套find_checkbox_by_text/set_checkbox_checked處理即可——但這一項
    # 本身還沒有實際跑過,第一次請核對截圖。
    "是否列印材料明細": False,
}


def run(main_win, out_lines, dry_run=False):
    query_win = jump_to_code(main_win, CODE, out_lines)
    if query_win is None:
        raise RuntimeError(f"沒有拿到{CODE}查詢畫面。")

    log(f"\n===== {CODE}: 設定選擇製令狀態(5個)+ 是否列印材料明細 =====")
    checkboxes = {}
    for label in TARGET_STATES:
        cb = find_checkbox_by_text(query_win, label)
        if cb is None:
            out_lines.append(f"  [警告] 找不到text=='{label}'的checkbox。")
            log(f"  [警告] 找不到text=='{label}'的checkbox。")
            continue
        checkboxes[label] = cb

    for label, want in TARGET_STATES.items():
        if label not in checkboxes:
            continue
        out_lines.append(f"--- {label}(目標: {'勾選' if want else '不勾選'}) ---")
        result = set_checkbox_checked(checkboxes[label], want, out_lines, label)
        if result != want:
            log(f"  [{label}] [注意] 最終狀態({result})跟目標({want})不符,請核對截圖。")

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
                         help="填完欄位後按取消,不會真的送出查詢")
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
