#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正式抓資料腳本 — IDLR69(庫存明細wh01 / 庫存明細wh02)
=========================================================

**2026-09-07更新(使用者人工實測驗證)**:原本這支會切到「進階選項」分頁
設定一組還沒實測過、憑設定表猜的checkbox目標值,結果切分頁的邏輯(Ctrl+
Tab+is_visible()驗證)在INVR18上被證實不可靠。使用者直接手動操作了一次
IDLR69完整流程,發現**根本不需要碰進階選項分頁**:輸入選擇庫別之後,
直接在「基本選項」分頁按4次Tab就能跳到「確認」鍵送出——代表這幾個
checkbox的ERP預設值本來就已經是需要的狀態,不用主動去設。這一版拿掉整段
猜測、未驗證的checkbox設定邏輯,改成跟IDLR16/17/68/MOCR10一樣的簡單流程
(選庫別 -> 直接送出),更貼近實際驗證過可行的操作。

「選擇庫別」用路線A(勾選「區間選擇」、起=迄=目標代號),IDLR69每次只選
「一個」庫別(wh01這次跑選WH01,wh02那次跑選WH02),不是INVR18「庫存
(全部)」那種多選,所以不用碰F2。

用法
----
    python report_idlr69.py                    # 真的跑完wh01+wh02兩組並送出
    python report_idlr69.py --only WH01        # 只跑其中一組
    python report_idlr69.py --dry-run          # 填完欄位按取消,不送出
"""

import argparse
import sys

from part2_inspect import (
    find_button,
    find_select_panel,
    jump_to_code,
    log,
    login,
    save_screenshot,
    select_single_via_range,
    submit_query,
)

CODE = "IDLR69"

# 兩次跑法,對照使用者設定表(庫存明細wh01 -> 選WH01;庫存明細wh02 -> 選WH02)
GROUPS = [
    {"name": "WH01"},
    {"name": "WH02"},
]


def run_one_group(main_win, group, out_lines, dry_run=False):
    log(f"\n========== IDLR69 跑法: {group['name']} ==========")
    out_lines.append(f"\n========== IDLR69 跑法: {group['name']} ==========")

    query_win = jump_to_code(main_win, CODE, out_lines)
    if query_win is None:
        raise RuntimeError(f"[{group['name']}] 沒有拿到{CODE}查詢畫面。")

    panel = find_select_panel(query_win, "選擇庫別")
    if panel is None:
        raise RuntimeError(f"[{group['name']}] 查詢畫面裡沒找到text=='選擇庫別'的SelectPanel。")

    log(f"選擇庫別: {group['name']}(路線A,單一代號,不用F2)")
    select_single_via_range(panel, group["name"], out_lines)

    log("維持ERP原本預設,不碰進階選項頁籤(2026-09-07 使用者實測確認不需要)。")

    save_screenshot(query_win, f"idlr69_{group['name']}_before_submit.png", out_lines)

    if dry_run:
        log(f"  --dry-run: 不送出,按取消關閉{CODE}查詢畫面。")
        cancel_btn = find_button(query_win, "取消")
        if cancel_btn is not None:
            cancel_btn.click()
        return None

    return submit_query(query_win, CODE, out_lines)


def run(main_win, out_lines, only=None, dry_run=False):
    """依序跑完GROUPS(WH01、WH02)。2026-09-08修正:原本一組失敗會直接
    中斷整個迴圈,改成每組try/except、失敗記錄後繼續跑下一組,避免INVR18
    9/7實測遇到的問題(一組壞掉害後面全部組別沒跑到)在IDLR69身上重演。"""
    groups = GROUPS if not only else [g for g in GROUPS if g["name"] == only]
    if only and not groups:
        raise RuntimeError(f"--only指定的名稱'{only}'不在GROUPS清單裡,有效值: "
                            f"{[g['name'] for g in GROUPS]}")
    results = []
    failures = []
    for group in groups:
        try:
            results.append(run_one_group(main_win, group, out_lines, dry_run=dry_run))
        except Exception as e:
            msg = f"[{group['name']}] 失敗: {e!r}"
            log(f"  [錯誤] {msg}(繼續跑下一組,不中斷整支IDLR69)")
            out_lines.append(f"  [錯誤] {msg}")
            results.append(None)
            failures.append(msg)
    if failures:
        raise RuntimeError(
            f"IDLR69: {len(failures)}/{len(groups)}組失敗(其餘組別已照常送出)-> "
            + " | ".join(failures)
        )
    return results


def main():
    parser = argparse.ArgumentParser(description=f"正式抓資料腳本 — {CODE}(2026-09-07已由使用者人工實測驗證WH01/WH02皆可成功送出)")
    parser.add_argument("--only", default=None, help="只跑其中一組,例如WH01")
    parser.add_argument("--dry-run", action="store_true",
                         help="填完欄位後按取消,不會真的送出查詢")
    args = parser.parse_args()

    out_lines = []
    try:
        app, main_win = login(out_lines)
        run(main_win, out_lines, only=args.only, dry_run=args.dry_run)
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
