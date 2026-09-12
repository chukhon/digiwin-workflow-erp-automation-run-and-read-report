#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
正式抓資料腳本 — INVR18(庫存明細表)
=======================================

狀態(2026-09-07更新,第7輪修正):9/7實測發現除WH01外其他5組都因為分頁
切換邏輯(Ctrl+Tab+is_visible()驗證)誤判卡在錯的分頁(「屬性類別選項」
而不是「進階選項」)而全部失敗——owner-drawn分頁控制項就算不在使用中的
分頁,is_visible()還是可能回傳True,這個驗證方式本質上不可靠。已改成完全
比照使用者人工實測驗證過的固定按鍵順序(純鍵盤操作,不找也不驗證任何控制
項),細節見`set_advanced_options()`。同時`jump_to_code()`(part2_inspect.py)
也補上「只認新冒出來的視窗」的判斷,`submit_query()`補上送出後ESC關閉查詢
畫面的步驟,避免同一個代號連續處理多組時抓到殘留的舊視窗。

路線A(WH05/WH06/WH08)後續又測出「起」欄位常被清空、分頁卡在「基本選項」
沒切過去的問題,第2/5/6輪陸續猜測是「填完欄位後焦點還停在文字框裡,把
Ctrl+Right吃掉」,依序試過補`{TAB}`、把焦點`.set_focus()`釘到checkbox上;
第7輪照使用者手動實測(打完代號原地按Ctrl+Right就會切分頁)改成完全不做
任何焦點操作——但使用者照第7輪的版本再測一次,WH05還是**一樣失敗**(起被
清空、卡在基本選項),證明「移開焦點」跟「什麼都不做」都不是正確方向。

第8輪(這次)重新比對才抓到真正被忽略的差異:使用者手動操作是「真的用鍵盤
打字」,打完之後鍵盤焦點自然停在「迄」欄位上;但我們填值用的
`set_edit_text()`是直接送WM_SETTEXT訊息改文字內容,**從頭到尾都不會真的
改變鍵盤焦點**——畫面看起來欄位填對了,但查詢視窗當下鍵盤焦點可能還停在
視窗剛開啟時的預設控制項上,跟「迄」欄位完全無關。另外,原本
`set_advanced_options()`把5個按鍵步驟拆成5次個別呼叫`send_key_sequence()`,
每次呼叫都會重新`set_focus()`一次,如果一開始焦點就不在對的地方,等於
後面每一步都各自從同一個(錯的)起點重新出發,而不是像使用者那樣連續按下去。
第8輪修法(兩處):
1. `select_single_via_range()`填完起/迄驗證過之後,明確把鍵盤焦點
   `.set_focus()`到「迄」欄位本身,重現使用者打完字後游標停在該處的真實
   狀態。
2. `set_advanced_options()`改成只呼叫一次`send_key_sequence()`,把全部
   7個按鍵(Ctrl+Right x2、Tab、Space、Tab、Down x2)一次送完,不再中途
   重新對焦。
`reverify_range_values()`送出前的保險檢查維持不動,當作最後一道防線。
**這一版尚未經過完整run_all_reports.py實機驗證,需要重新測一次WH05/WH06/
wh08確認畫面真的停在「進階選項」且checkbox/下拉都改對。**

**建議第一次先用--only參數只跑其中一組(例如--only 庫存(全部),這組
touch_advanced=True且用F2多選,涵蓋面最完整)測試過,確認沒問題之後,再
拿掉--only跑完整6組**。

6組設定(已由使用者 2026-09-03核對確認,對照erp-report-selection-methods-
2026-09-03.md):

| 分頁         | 選擇庫別代號                                    | 路線 | 動進階選項 |
|--------------|--------------------------------------------------|------|-----------|
| WH01         | WH01                                              | A    | 否(維持預設)|
| 庫存(全部)   | WH04, WH05, WH06, WH07, WH08, WH03, WH02          | B(F2)| 是         |
| WH03 WH02    | WH03, WH02                                        | B(F2)| 是         |
| WH05         | WH05                                              | A    | 是         |
| WH06         | WH06                                              | A    | 是         |
| WH08         | WH08                                              | A    | 是         |

「動進階選項」= 是:把「列印各庫明細」取消勾選,「選擇列印庫存」下拉選「全部」。

用法
----
    python report_invr18.py                  # 真的跑完整6組並送出
    python report_invr18.py --only WH01       # 只跑WH01這一組(建議第一次這樣測)
    python report_invr18.py --dry-run         # 全部6組都填完欄位、按取消,不送出
"""

import argparse
import sys

from part2_inspect import (
    find_button,
    find_by_class,
    find_select_panel,
    jump_to_code,
    log,
    login,
    reverify_range_values,
    save_screenshot,
    select_codes_via_f2,
    select_single_via_range,
    send_key_sequence,
    submit_query,
)

CODE = "INVR18"

# 每組: name(只是給人看的標籤) / codes(list) / route("A"或"F2") / touch_advanced(bool)
GROUPS = [
    {"name": "WH01", "codes": ["WH01"], "route": "A", "touch_advanced": False},
    {"name": "庫存(全部)", "codes": ["WH04", "WH05", "WH06", "WH07", "WH08", "WH03", "WH02"],
     "route": "F2", "touch_advanced": True},
    {"name": "WH03 WH02", "codes": ["WH03", "WH02"], "route": "F2", "touch_advanced": True},
    {"name": "WH05", "codes": ["WH05"], "route": "A", "touch_advanced": True},
    {"name": "WH06", "codes": ["WH06"], "route": "A", "touch_advanced": True},
    {"name": "WH08", "codes": ["WH08"], "route": "A", "touch_advanced": True},
]


def set_advanced_options(query_win, out_lines):
    """切到進階選項:「列印各庫明細」取消勾選,「選擇列印庫存」選「全部」。

    2026-09-07改版:原本用Ctrl+Tab+is_visible()驗證切分頁、再找checkbox/
    combo控制項操作,9/7實測發現除了WH01之外其他5組全部卡在「屬性類別
    選項」分頁而不是「進階選項」——owner-drawn的分頁控制項就算不在使用中
    的分頁,裡面的控制項is_visible()還是可能回傳True,這個驗證方式會誤判,
    導致INVR18當天只有WH01那1組真的送出去,後面5組都因為抓錯分頁失敗。

    改成完全比照使用者人工實測驗證過、可靠的固定按鍵順序,純鍵盤操作、不找
    也不驗證任何控制項:
        Ctrl+Right x2  切到「進階選項」分頁(從「基本選項」分頁起算)
        Tab            移到「列印各庫明細」checkbox(分頁裡第2個欄位)
        Space          取消勾選
        Tab            移到「選擇列印庫存」下拉選單
        Down x2        選到「全部」(第3個選項)

    2026-09-07第8輪修正:原本這5個步驟是分成5次個別呼叫`send_key_sequence()`,
    但`send_key_sequence()`每次一開始都會呼叫`win.set_focus()`——對使用者手動
    操作來說,這整串鍵盤操作是「一口氣」連續按下去、中間沒有任何重新
    對焦的動作,但我們的程式卻在5個步驟之間插了4次額外的`set_focus()`。
    如果查詢視窗當下並沒有真的把鍵盤焦點停在「迄」欄位上(見
    `select_single_via_range()`第8輪的修正說明——`set_edit_text()`本來就
    不會改變鍵盤焦點),那麼每一次重複呼叫`set_focus()`都有可能把焦點拉回
    某個「預設」控制項,等於後面4個步驟其實都是各自從同一個起點重新出發,
    而不是像使用者手動操作那樣接續進行。改成只在一開始呼叫一次
    `send_key_sequence()`,把全部按鍵一次送完,不再中途重新對焦,更貼近
    使用者實測有效的連續按鍵方式。
    """
    send_key_sequence(
        query_win,
        ["^{RIGHT}", "^{RIGHT}", "{TAB}", "{SPACE}", "{TAB}", "{DOWN}", "{DOWN}"],
        out_lines,
        label="切到進階選項+取消列印各庫明細+選擇列印庫存全部(一次連續送出)",
        pause=0.4,
    )
    # 注意:pywinauto的type_keys() mini-language裡,單獨一個" "(空白字元)
    # 會被當成「按鍵群組之間的分隔符號」處理,不會真的送出空白鍵——這是
    # 2026-09-07第2輪修正時漏掉的地方,結果log裡看起來「已送出」,但
    # 螢幕截圖顯示「列印各庫明細」checkbox其實根本沒被取消勾選(選擇列印
    # 庫存倒是有正確變成「全部」,因為Down/Tab都是用{}包起來的token,不受
    # 影響)。跟其他按鍵token一樣,空白鍵要送"{SPACE}"才會真的生效,所以
    # 上面合併後的序列裡用的是"{SPACE}"不是" "。


def run_one_group(main_win, group, out_lines, dry_run=False):
    log(f"\n========== INVR18 分頁組: {group['name']} ==========")
    out_lines.append(f"\n========== INVR18 分頁組: {group['name']} ==========")

    query_win = jump_to_code(main_win, CODE, out_lines)
    if query_win is None:
        raise RuntimeError(f"[{group['name']}] 沒有拿到{CODE}查詢畫面。")

    panel = find_select_panel(query_win, "選擇庫別")
    if panel is None:
        raise RuntimeError(f"[{group['name']}] 查詢畫面裡沒找到text=='選擇庫別'的SelectPanel。")

    log(f"選擇庫別: {group['codes']}(路線{group['route']})")
    if group["route"] == "A":
        if len(group["codes"]) != 1:
            raise RuntimeError(f"[{group['name']}] 路線A只支援單一代號,設定卻給了{group['codes']}。")
        select_single_via_range(panel, group["codes"][0], out_lines)
    else:
        list_view = find_by_class(panel, "TDSYenListView.UnicodeClass", "ListView")[0]
        select_codes_via_f2(list_view, group["codes"], out_lines)

    if group["touch_advanced"]:
        log("切到進階選項:取消列印各庫明細、選擇列印庫存=全部")
        set_advanced_options(query_win, out_lines)
    else:
        log("維持ERP原本預設,不碰進階選項頁籤。")

    # 2026-09-07第2輪新增:9/7實測發現路線A(WH05/WH06/WH08)即使
    # select_single_via_range()當下驗證起/迄都是對的,後面切到進階選項的
    # Ctrl+Right操作偶爾還是會把「起」的值弄不見。select_single_via_range()
    # 內部已經補了Tab把焦點移開這個主要肇因,這裡再做最後一道保險——真正
    # 送出前重新讀回「選擇庫別」起/迄目前的實際值,不對就修,還是修不好就
    # 直接擋下不送出。只有路線A需要檢查(F2路線的選庫別方式不同,不會踩到
    # 這個問題)。
    if group["route"] == "A":
        reverify_range_values(panel, group["codes"][0], out_lines, label="選擇庫別")

    save_screenshot(query_win, f"invr18_{group['name']}_before_submit.png".replace(" ", "_"), out_lines)

    if dry_run:
        log(f"  --dry-run: 不送出,按取消關閉{CODE}查詢畫面。")
        cancel_btn = find_button(query_win, "取消")
        if cancel_btn is not None:
            cancel_btn.click()
        return None

    return submit_query(query_win, CODE, out_lines)


def run(main_win, out_lines, only=None, dry_run=False):
    """依序跑完GROUPS。2026-09-08修正:原本任一組拋例外會直接中斷整個
    迴圈,剩下的組別完全不會跑到——9/7實測就是這樣:第2組(庫存全部)因為
    jump_to_code抓到舊視窗而失敗,導致後面WH05/WH06/WH08/WH03 wh02這4組
    當天完全沒跑(比對9/3人工跑的結果才發現少了這些)。現在改成每組各自
    try/except,一組失敗記錄下來、繼續跑下一組,全部跑完之後如果有失敗的
    才在最後一次性丟出RuntimeError(run_all_reports.py那層還是會把INVR18
    整支標成失敗,但至少沒失敗的組別都已經送出去了)。"""
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
            log(f"  [錯誤] {msg}(繼續跑下一組,不中斷整支INVR18)")
            out_lines.append(f"  [錯誤] {msg}")
            results.append(None)
            failures.append(msg)
    if failures:
        raise RuntimeError(
            f"INVR18: {len(failures)}/{len(groups)}組失敗(其餘組別已照常送出)-> "
            + " | ".join(failures)
        )
    return results


def main():
    parser = argparse.ArgumentParser(description=f"正式抓資料腳本 — {CODE}(6組分頁)")
    parser.add_argument("--only", default=None,
                         help="只跑其中一組(用GROUPS裡的name,例如WH01),建議第一次先這樣測")
    parser.add_argument("--dry-run", action="store_true",
                         help="每組都填完欄位後按取消,不會真的送出查詢")
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
