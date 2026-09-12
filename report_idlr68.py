#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正式抓資料腳本 — IDLR68(銷貨單明細表)
=========================================

狀態:全部欄位邏輯已完整驗證成功(含Ctrl+Tab切分頁),這支是把
part2_idlr68_test.py的驗證過邏輯,改成呼叫part2_inspect.py共用函式
+ 真的按確認送出。

欄位邏輯
--------
1. 「選擇銷貨日期」起=REPORT_WINDOW_START_DATE(見下方常數)、
   迄=今天(執行當下動態算,民國年格式)。
2. Ctrl+Tab切到「進階選項」分頁,「選擇結帳狀態」下拉改選「全部」。

關於「查詢起始日」怎麼挑——一個容易踩到的坑
--------------------------------------------------------------------------
這支報表原本的直覺寫法是「起始日=當月1號」,只抓當月至今的資料,因為多數
情況下只需要「這個月目前為止的出貨」。但如果下游有任何流程會把這份報表的
結果**累積成一份長期歷史清單**(例如每次都跟已存的舊資料比對、只把沒看過的
新記錄補進去),「當月1號」這個設計會有一個不容易發現的破洞:假設某次執行
之後,中間隔了超過一個月才又執行下一次,下一次抓到的「當月」範圍,會完全
跳過上一次執行之後、到這次「當月1號」之前的那一整段期間——這段資料在ERP裡
從此沒有任何一次查詢範圍能再涵蓋到,不是重跑就能補回來的資料遺失。

修法:起始日期改成一個**固定的、夠早的錨點日期**(`REPORT_WINDOW_START_DATE`),
讓這份報表每次都撈「錨點至今」的資料,不再只抓當月——不管兩次執行之間隔了
多久,只要下游用唯一key(這裡是銷貨單號+品號+訂單單號+批號)去重合併,就
不會有任何一段期間永遠補不回來的情況。

代價是查詢/匯出的資料量會隨錨點與今天的距離變大。這裡示範一個簡單、寫死
日期的版本;如果你的下游系統會持續記錄「目前已經確認安全處理過的最新一筆
資料日期是哪天」,可以把這個常數改成動態算出來的「該日期往前留一段安全緩衝
天數」,錨點本身只當作「安全下限」(下游系統太久沒更新、不可信任時的退路)
——概念上是同一件事:找一個「保證涵蓋所有可能被漏掉的期間」的起始日,只是
把它做成可以隨時間自動往前推進、不必永遠寫死。

用法
----
    python report_idlr68.py              # 真的送出
    python report_idlr68.py --dry-run    # 填完欄位、按取消,不送出
"""

import argparse
import datetime
import sys
import time

from part2_inspect import (
    find_by_class,
    find_button,
    find_select_panel,
    find_tab_sheet,
    jump_to_code,
    log,
    login,
    roc_date_str,
    save_screenshot,
    submit_query,
    switch_tab_by_ctrl_tab,
)

CODE = "IDLR68"
MAX_TAB_PRESSES = 6

# 查詢起始日的固定錨點(見上方說明)——按自己的需求調整,通常設成「你的下游
# 歷史資料最早可能需要回補的日期」就足夠安全。
REPORT_WINDOW_START_DATE = datetime.date(2026, 1, 1)


def determine_query_start_date():
    """
    決定這次查詢要往回抓多久。這裡示範最簡單的版本:永遠回到固定錨點。

    如果你的下游系統會持續記錄「目前已經確認安全處理過的最新一筆資料日期」,
    可以在這裡加上:讀取那個記錄 -> 該日期往前留一段安全緩衝天數 -> 跟固定
    錨點取較晚者(絕對不會比固定錨點更早,錨點是安全下限)。這樣一來,只要
    下游系統持續正常運作,查詢範圍會隨時間自動縮短,不會無限期變大;下游
    系統太久沒更新、不可信任時,就自動退回固定錨點,不會冒然縮小查詢範圍
    承擔漏資料的風險。
    """
    return REPORT_WINDOW_START_DATE


def fill_dates(query_win, out_lines):
    today = datetime.date.today()
    start_date = determine_query_start_date()
    start_str = roc_date_str(start_date)
    end_str = roc_date_str(today)
    log(f"  今天(民國)={end_str}  查詢起始日(民國)={start_str}(西元 {start_date})")

    panel = find_select_panel(query_win, "選擇銷貨日期")
    if panel is None:
        raise RuntimeError("查詢畫面裡沒找到text=='選擇銷貨日期'的SelectPanel。")

    edits = find_by_class(panel, "TDSYenMaskEdit.UnicodeClass", "Edit")
    if len(edits) < 2:
        raise RuntimeError(f"「選擇銷貨日期」panel底下只找到{len(edits)}個MaskEdit(預期至少2個)。")
    edits_sorted = sorted(edits, key=lambda c: c.rectangle().top)
    start_edit, end_edit = edits_sorted[0], edits_sorted[1]

    start_edit.set_edit_text(start_str)
    time.sleep(0.3)
    end_edit.set_edit_text(end_str)
    time.sleep(0.3)

    after_start = start_edit.window_text()
    after_end = end_edit.window_text()
    log(f"  填入後 起={after_start!r} 迄={after_end!r}")
    out_lines.append(f"[選擇銷貨日期] 起={after_start!r} 迄={after_end!r}")


def set_settle_status_all(query_win, out_lines):
    adv_sheet = find_tab_sheet(query_win, "進階選項")
    if adv_sheet is None:
        raise RuntimeError("查詢畫面裡沒找到text=='進階選項'的TabSheet。")
    combos = find_by_class(adv_sheet, "TDSUCComboBox.UnicodeClass", "ComboBox")
    if len(combos) < 2:
        raise RuntimeError(
            f"「進階選項」TabSheet底下只找到{len(combos)}個ComboBox(預期至少2個)。"
        )
    combos_sorted = sorted(combos, key=lambda c: c.rectangle().top)
    settle_status_combo = combos_sorted[1]  # 第2個(由上往下)= 選擇結帳狀態

    switched, presses = switch_tab_by_ctrl_tab(
        query_win, settle_status_combo, out_lines, max_presses=MAX_TAB_PRESSES
    )
    if not switched:
        raise RuntimeError(
            f"按了{presses}次Ctrl+Tab,「選擇結帳狀態」combo還是不可見,無法繼續。"
        )

    settle_status_combo.select("全部")
    time.sleep(0.3)
    new_text = settle_status_combo.window_text()
    log(f"  「選擇結帳狀態」選完之後: {new_text!r}")
    out_lines.append(f"[選擇結帳狀態] 選完之後: {new_text!r}")
    if new_text != "全部":
        out_lines.append("[注意] 選完之後的文字不是預期的'全部',請核對截圖。")


def run(main_win, out_lines, dry_run=False):
    query_win = jump_to_code(main_win, CODE, out_lines)
    if query_win is None:
        raise RuntimeError(f"沒有拿到{CODE}查詢畫面。")

    log(f"\n===== {CODE}: 填選擇銷貨日期 =====")
    fill_dates(query_win, out_lines)

    log(f"\n===== {CODE}: 切到進階選項,選擇結帳狀態=全部 =====")
    set_settle_status_all(query_win, out_lines)

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
