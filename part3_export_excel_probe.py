#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
閱覽報表 -> 存成EXCEL檔(全自動,含格式對話框+等待匯出完成+關閉視窗)
=======================================================================
2026-09-07

背景
----
接續`part2_reportview_probe.py`已經走通的鏈路(登入 -> MainMenu顯示報表
工作序列 -> 佇列工作管理員開啟檔案 -> 確認閱覽報表視窗變成MDI子視窗 ->
找到工具列)。使用者已經親自懸停確認過:閱覽報表工具列上相對座標x≈210這顆
圖示的tooltip就是「存成EXCEL檔」,第一版腳本點下去之後,確認彈出了
`TfrmToExcelDlg.UnicodeClass`格式設定視窗——證實座標抓對了。

使用者接著親自手動測試了一次完整流程,並回報操作方式:

1. 「存成EXCEL檔」格式設定視窗跳出來後,预設焦點在「Excel欄位格式設定」
   那組選項的第一顆(「不設定(依Excel預設)」),按**右鍵**切到第二顆
   (「依鼎新預設格式儲存」),再按**兩次TAB**切到「確定」按鈕,按
   **Enter**觸發匯出。
2. 匯出過程中會跳出一個「處理進度顯示」視窗,顯示百分比進度條,**每份
   報表資料量不同,跑多久沒有固定**——不能用固定的time.sleep()等,要用
   輪詢(每隔幾秒檢查一次)的方式,一直等到看到「完成」的訊號為止。
3. 進度跑完之後,會跳出「閱覽報表(...)  檢視Excel資料?」的確認對話框
   (Yes/No按鈕,雖然問句是中文,按鈕文字卻是英文Yes/No)——使用者按了
   **No**(不用另外開Excel顯示,保持流程單純)。
4. 使用者接著問下一步要「找出藍圈的位置把此報表關閉」——**這裡不需要**
   **靠像素座標點X按鈕**:前幾輪一直沒能用tooltip/掃描方式定出標題列
   關閉鈕的精確座標,但根本不需要那麼麻煩——直接送標準的Windows
   `WM_CLOSE`訊息給閱覽報表視窗本身,這是關閉視窗最標準、最可靠的方式,
   MDI子視窗一樣支援(Delphi會呼叫對應表單的Close/OnCloseQuery事件),
   完全不需要知道X按鈕在螢幕上的像素位置。

這支腳本把上面這整個流程串起來,全自動執行:

1. (跟之前一樣)登入 -> 開佇列工作管理員 -> 開啟檔案 -> 確認閱覽報表視窗
   ->找到工具列。
2. 記錄`C:\\My Documents\\`(使用者限定的匯出路徑,不能改設定)目前的
   .xls/.xlsx檔案,當作「匯出前」快照。
3. 點擊工具列相對座標(210, 工具列垂直中心)觸發「存成EXCEL檔」。
4. 等到`TfrmToExcelDlg.UnicodeClass`格式設定視窗跳出來,送出按鍵序列
   `{RIGHT}{TAB}{TAB}{ENTER}`(選「依鼎新預設格式儲存」->TAB兩次到
   「確定」->Enter),依使用者手動測試過的操作重現。
5. **輪詢等待**(不是固定sleep)Yes/No確認對話框出現,逾時上限拉長
   (預設20分鐘)因應報表資料量大小不一,並且每隔一段時間在console印一次
   「還在等」的心跳訊息,方便你確認腳本沒有卡死。
6. 找到確認對話框後點擊「No」按鈕(不開Excel)。
7. 再檢查一次`C:\\My Documents\\`,比對前後差異,列出新增的檔案。
8. 對閱覽報表視窗送出`WM_CLOSE`訊息關閉它(不用找X按鈕座標)。

第9輪修正(2026-09-07):第一版第7站原本用「視窗文字裡有沒有包含
『檢視Excel資料』這段字」去等,實測發現**等到逾時都等不到**,即使那個
確認對話框明明已經顯示在畫面上——研判是那段文字其實是對話框裡一顆Label
子控制項顯示的內容,不是視窗自己的標題,而原本邏輯對桌面頂層視窗只檢查
它自己的window_text(),沒有連子控制項一起搜。改成直接找「Yes」「No」
這兩顆按鈕本身(使用者的截圖已確認按鈕文字就是英文Yes/No),不管掛在哪個
容器底下都能找到,不用再猜文字位置。另外也把等待「閱覽報表」視窗出現的
逾時從20秒拉長到45秒(使用者這次實測有踩到20秒逾時、退回備案抓錯視窗的
情況,雖然後面的reacquire_rv_win()有把它救回來,但拉長逾時比較不會多繞
這一圈)。

執行需求
--------
    python part3_export_excel_probe.py
"""

import glob
import os
import sys
import time

from part2_inspect import dump_control_tree, find_by_class, log, login, safe_all_windows
from part2_reportview_probe import (
    MM_QUEUE_BTN,
    QM_CLASS,
    QM_OPEN_BTN,
    all_top_windows,
    click_toolbar_icon,
    save_shot,
    try_dismiss_error_dialog,
    wait_new_window,
)

# 使用者限定用現有的匯出路徑,不能改設定(erp-stageb-open-report-2026-09-07.md
# 裡記錄的兩個硬性限制之一)。
EXPORT_DIR = r"C:\My Documents"

# 使用者親自懸停確認過:閱覽報表工具列上相對x≈210(工具列自己左上角算起)
# 這顆圖示的tooltip文字是「存成EXCEL檔」。
EXCEL_BTN_REL_X = 210

# 「存成EXCEL檔」格式設定視窗的class name(第一版實測時確認過)。
EXCEL_DLG_CLASS = "TfrmToExcelDlg.UnicodeClass"

# 每份報表資料量不同,匯出時間沒有固定,這裡給一個寬鬆的輪詢逾時上限
# (秒)。真的跑到逾時代表可能卡住了,需要人工看一下畫面。
EXPORT_WAIT_TIMEOUT = 20 * 60
EXPORT_WAIT_POLL = 2.0
EXPORT_WAIT_HEARTBEAT = 15.0


def reacquire_rv_win(rv_win, qm_win, out):
    """比照part2_reportview_probe.py裡的同名邏輯:先確認原本抓到的視窗
    控制代碼是否還存在,不管存不存在都優先在桌面頂層視窗裡找同class的
    視窗,找不到才退回qm_win底下的子視窗清單(MDI子視窗)裡找。"""
    import win32gui

    still_valid = False
    try:
        still_valid = win32gui.IsWindow(rv_win.handle)
    except Exception:
        pass
    if not still_valid:
        out.append("[第3站] rv_win原本的控制代碼已經不存在了,改用class重新搜尋...")

    hits = [w for w in all_top_windows()
            if w.class_name() == "TfrmReportViewer.UnicodeClass"]
    if hits:
        return hits[0]
    try:
        child_hits = [c for c in qm_win.descendants()
                      if c.class_name() == "TfrmReportViewer.UnicodeClass"]
    except Exception as e:
        child_hits = []
        out.append(f"[第3站] 搜尋qm_win子視窗時發生例外: {e!r}")
    if child_hits:
        out.append("[第3站] 頂層視窗清單裡沒有,但在qm_win底下的子視窗裡找到了"
                    "——很可能已經變成MDI子視窗。")
        return child_hits[0]
    return rv_win if still_valid else None


def find_new_window_by_class(known_before, cls, qm_win, out, timeout=10, poll=0.3):
    """等一個「新出現、指定class」的視窗——先在桌面頂層視窗裡找,找不到
    再到qm_win底下的子視窗清單裡找(防範它也跟rv_win一樣被reparent成
    MDI子視窗)。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        news = [w for w in all_top_windows() if w.handle not in known_before]
        hits = [w for w in news if w.class_name() == cls]
        if hits:
            return hits[0]
        try:
            child_hits = [c for c in qm_win.descendants()
                          if c.class_name() == cls and c.handle not in known_before]
        except Exception:
            child_hits = []
        if child_hits:
            out.append(f"[提示] class={cls!r}的視窗是在qm_win子視窗裡找到的,不是頂層視窗。")
            return child_hits[0]
        time.sleep(poll)
    return None


def wait_for_yes_no_dialog(qm_win, out, timeout=EXPORT_WAIT_TIMEOUT,
                            poll=EXPORT_WAIT_POLL, heartbeat=EXPORT_WAIT_HEARTBEAT):
    """輪詢等待「檢視Excel資料?」確認對話框出現,回傳({'yes':控制項,
    'no':控制項}, 已等待秒數),沒等到就回傳(None, 已等待秒數)。

    2026-09-07第9輪修正:第一版用「桌面頂層視窗的window_text()裡有沒有
    包含『檢視Excel資料』這段文字」去等,結果使用者實測**等到逾時都等不到,
    即使那個確認對話框明明已經顯示在畫面上**。研判是「檢視Excel資料?」
    這行字其實是對話框裡一顆Label子控制項顯示的內容,不是這個視窗自己的
    標題(window_text()讀到的是caption,不是子控制項內容);而原本的
    邏輯對「桌面頂層視窗」只檢查它自己的window_text(),沒有連子控制項
    一起搜,自然永遠等不到。

    改法:不再去猜文字放在哪個層級,直接找**「Yes」跟「No」這兩顆按鈕**
    本身(使用者的截圖已經確認,不管問句是不是中文,按鈕文字就是英文
    Yes/No)——不管這兩顆按鈕實際上掛在qm_win.descendants()裡,還是掛在
    某個桌面頂層視窗自己的descendants()裡,都能找到,不用再猜文字位置。
    找到後直接回傳這兩顆按鈕控制項,呼叫端可以直接點,不用再另外定位。

    第10輪:改成廣泛找Yes/No按鈕之後,使用者實測**還是一樣卡住**,console
    持續印「還在等」,但畫面上確實顯示著「檢視Excel資料?」的Yes/No對話框。
    懷疑這次猜錯的地方換成了：Windows/Delphi標準對話框按鈕的文字資源字串
    通常內建「&」字元做為快捷鍵底線標記(例如Delphi VCL的SMsgDlgYes/
    SMsgDlgNo資源字串本身就是`'&Yes'`/`'&No'`,畫面上顯示成底線的Y/N,
    但`window_text()`讀到的是原始字串,**很可能是`'&Yes'`/`'&No'`,不是
    純`'Yes'`/`'No'`**)——之前兩次都用完全相等比對,只要多一個`&`字元就
    會比對失敗。這裡改成比對前先把`&`字元去掉再比。另外加一個一次性的
    診斷:等超過20秒還沒找到,就把目前看到的所有「短文字」(長度<=12)
    候選記錄進輸出檔,萬一這次假設又猜錯,至少能直接從記錄檔看到真正的
    文字內容是什麼,不用再靠螢幕截圖用眼睛看。"""
    def _norm(t):
        return (t or "").replace("&", "").strip()

    start = time.time()
    last_beat = start
    dumped_diag = False
    while time.time() - start < timeout:
        candidates = []
        try:
            candidates.extend(qm_win.descendants())
        except Exception:
            pass
        try:
            for w in all_top_windows():
                candidates.append(w)
                try:
                    candidates.extend(w.descendants())
                except Exception:
                    pass
        except Exception:
            pass

        yes_btn = None
        no_btn = None
        seen_texts = set()
        for c in candidates:
            try:
                t = (c.window_text() or "").strip()
            except Exception:
                continue
            if not t:
                continue
            seen_texts.add(t)
            t_norm = _norm(t)
            if t_norm == "Yes" and yes_btn is None:
                yes_btn = c
            elif t_norm == "No" and no_btn is None:
                no_btn = c
            if yes_btn is not None and no_btn is not None:
                return {"yes": yes_btn, "no": no_btn}, time.time() - start

        if not dumped_diag and time.time() - start > 20:
            short_texts = sorted({t for t in seen_texts if len(t) <= 12})
            out.append(f"[第7站診斷] 等待{int(time.time() - start)}秒後仍未找到Yes/No,"
                        f"目前看到的短文字候選(長度<=12): {short_texts}")
            log(f"  [診斷] 短文字候選: {short_texts}")
            dumped_diag = True

        if time.time() - last_beat > heartbeat:
            elapsed = int(time.time() - start)
            log(f"  ...還在等匯出完成(已等待{elapsed}秒,逾時上限{timeout}秒)...")
            last_beat = time.time()
        time.sleep(poll)
    return None, time.time() - start


def close_report_window(rv_win, out):
    """關閉閱覽報表視窗——不用去找標題列右上角那個X按鈕的精確座標(前幾輪
    一直沒能用tooltip/掃描方式定位到那個座標),直接送標準Windows
    `WM_CLOSE`訊息給視窗本身,這是關閉視窗最標準、最可靠的方式,MDI子
    視窗一樣支援(Delphi會呼叫對應表單的Close/OnCloseQuery事件),完全
    不需要知道X按鈕在螢幕上的像素位置。"""
    import win32con
    import win32gui

    try:
        win32gui.PostMessage(rv_win.handle, win32con.WM_CLOSE, 0, 0)
        out.append("[關閉閱覽報表視窗] 已送出WM_CLOSE訊息。")
        log("  已送出WM_CLOSE訊息關閉閱覽報表視窗。")
        return True
    except Exception as e:
        out.append(f"[關閉閱覽報表視窗] 送WM_CLOSE失敗: {e!r}")
        log(f"  [警告] 送WM_CLOSE失敗: {e!r}")
        return False


def main():
    out = []
    try:
        app, main_win = login(out)

        log("\n===== 第1站:點「顯示報表工作序列」 =====")
        known = {w.handle for w in all_top_windows()}
        click_toolbar_icon(main_win, MM_QUEUE_BTN, out, "MainMenu顯示報表工作序列")
        qm_win, _ = wait_new_window(known, out, "佇列工作管理員", timeout=25, want_class=QM_CLASS)
        if qm_win is None:
            out.append("[結果] 沒等到佇列工作管理員視窗,停止。")
            return
        log(f"  開啟成功: {qm_win.window_text()!r}")
        try:
            qm_win.set_focus()
        except Exception:
            pass
        time.sleep(2.5)

        log("\n===== 第2站:點「開啟檔案」開報表 =====")
        known2 = {w.handle for w in all_top_windows()}
        click_toolbar_icon(qm_win, QM_OPEN_BTN, out, "佇列工作管理員開啟檔案")
        # 2026-09-07第9輪:實測看到這一步逾時20秒的例子(報表資料量大,
        # 期間跳出好幾個過渡視窗:「刪除格式」提示、「設定數值顯示格式」
        # 對話框等),拉長到45秒降低誤觸備案邏輯的機率。就算真的等不到
        # 指定class、退回備案抓錯視窗也不影響最終結果——後面的
        # reacquire_rv_win()一定會重新用class名稱去搜,不會沿用抓錯的
        # 那個。
        rv_win, _ = wait_new_window(known2, out, "閱覽報表", timeout=45,
                                     want_class="TfrmReportViewer.UnicodeClass")
        if rv_win is None:
            out.append("[警告] 等不到指定class,改用通用邏輯當備案...")
            rv_win, _ = wait_new_window(known2, out, "閱覽報表(備案)", timeout=15)
        if rv_win is None:
            out.append("[結果] 按下「開啟檔案」之後沒有新視窗,停止。")
            return
        log(f"  開啟成功: class={rv_win.class_name()!r} text={rv_win.window_text()!r}")
        out.append(f"[第2站成功] class={rv_win.class_name()!r} text={rv_win.window_text()!r}")

        log("\n===== 第2站(續):先被動等5秒,再檢查有沒有錯誤訊息框 =====")
        time.sleep(5.0)
        try_dismiss_error_dialog(known2, out, wait=1.0)

        log("\n===== 第3站:重新確認閱覽報表視窗(可能已變成MDI子視窗) =====")
        rv_win_fresh = None
        for attempt in range(6):
            time.sleep(1.0)
            rv_win_fresh = reacquire_rv_win(rv_win, qm_win, out)
            if rv_win_fresh is None:
                continue
            try:
                r = rv_win_fresh.rectangle()
                if (r.right - r.left) > 10 and (r.bottom - r.top) > 10:
                    break
            except Exception:
                pass
        if rv_win_fresh is None:
            out.append("[結果] 多次重新搜尋後還是找不到閱覽報表視窗,停止,不會嘗試點擊。")
            log("  [錯誤] 找不到閱覽報表視窗,結束。")
            return
        rv_win = rv_win_fresh
        try:
            r = rv_win.rectangle()
            out.append(f"[第3站] 重新確認視窗: handle={rv_win.handle} rect={r}")
            log(f"  重新確認視窗: handle={rv_win.handle} rect={r}")
        except Exception as e:
            out.append(f"[第3站] 讀取視窗rect失敗: {e!r}")
            return

        tbs = find_by_class(rv_win, "TDSUCToolBar.UnicodeClass", "ToolBar")
        if not tbs:
            out.append("[結果] 找不到工具列,停止,不會嘗試點擊。")
            log("  [錯誤] 找不到工具列,結束。")
            return
        tb = tbs[0]
        tb_rect = tb.rectangle()
        y_rel = (tb_rect.bottom - tb_rect.top) // 2

        log(f"\n===== 第4站:存Excel前,先記錄{EXPORT_DIR}目前的.xls/.xlsx檔案 =====")
        try:
            before_files = set(glob.glob(os.path.join(EXPORT_DIR, "*.xls*")))
        except Exception as e:
            before_files = set()
            out.append(f"[警告] 讀取{EXPORT_DIR}目前檔案清單失敗: {e!r}")
        out.append(f"[第4站] {EXPORT_DIR} 目前有{len(before_files)}個.xls/.xlsx檔案。")
        log(f"  {EXPORT_DIR} 目前有{len(before_files)}個.xls/.xlsx檔案。")

        known3 = {w.handle for w in all_top_windows()}
        log(f"\n===== 第5站:點擊「存成EXCEL檔」(工具列相對座標({EXCEL_BTN_REL_X}, {y_rel}),"
            "使用者已親自懸停確認過tooltip文字) =====")
        click_toolbar_icon(rv_win, (EXCEL_BTN_REL_X, y_rel), out, "閱覽報表存成EXCEL檔")

        log("\n===== 第6站:等「存成EXCEL檔」格式設定視窗,依使用者手動測試過的操作送出按鍵 =====")
        excel_dlg = find_new_window_by_class(known3, EXCEL_DLG_CLASS, qm_win, out, timeout=10)
        if excel_dlg is None:
            out.append(f"[第6站] 等不到class={EXCEL_DLG_CLASS!r}的格式設定視窗,"
                        "可能按鈕座標偏了,或這次跳出方式不同,停止,不會亂送按鍵。")
            log("  [錯誤] 等不到格式設定視窗,結束。")
            return
        out.append(f"[第6站] 找到格式設定視窗: class={excel_dlg.class_name()!r} "
                    f"text={excel_dlg.window_text()!r}")
        log(f"  找到格式設定視窗: text={excel_dlg.window_text()!r}")
        try:
            excel_dlg.set_focus()
        except Exception:
            pass
        time.sleep(0.3)
        # 依使用者手動測試過的操作:按右鍵選「依鼎新預設格式儲存」,
        # 再按兩次TAB切到「確定」,最後按Enter觸發匯出。
        excel_dlg.type_keys("{RIGHT}{TAB}{TAB}{ENTER}")
        out.append("[第6站] 已送出按鍵序列 {RIGHT}{TAB}{TAB}{ENTER}"
                    "(選「依鼎新預設格式儲存」->TAB兩次到「確定」->Enter)。")
        log("  已送出按鍵序列,應該已經觸發匯出。")

        log("\n===== 第7站:輪詢等待匯出完成(每份報表耗時不同,不用固定sleep) =====")
        dialog_btns, waited = wait_for_yes_no_dialog(qm_win, out)
        if dialog_btns is None:
            out.append(f"[第7站] 等了{EXPORT_WAIT_TIMEOUT}秒還是沒看到Yes/No確認"
                        "對話框——可能還在跑、卡住了,需要人工確認畫面。")
            log(f"  [警告] 等了{EXPORT_WAIT_TIMEOUT}秒還是沒等到確認對話框。")
        else:
            out.append(f"[第7站] 等了約{waited:.0f}秒後,匯出完成,找到Yes/No確認對話框。")
            log(f"  匯出完成(等了約{waited:.0f}秒),找到Yes/No確認對話框。")

            log("\n===== 第8站:點擊確認對話框的「No」(依使用者手動測試過的操作,不開Excel) =====")
            try:
                dialog_btns["no"].click()
                out.append("[第8站] 已點擊「No」按鈕(不開啟Excel)。")
                log("  已點擊「No」按鈕。")
            except Exception as e:
                out.append(f"[第8站] 點擊No按鈕失敗: {e!r},請自己確認畫面、手動按掉這個對話框。")
                log(f"  [警告] 點擊No按鈕失敗: {e!r}")
            time.sleep(1.0)

        save_shot(rv_win, "export_probe_after_click.png", out)

        log(f"\n===== 第9站:比對{EXPORT_DIR}前後檔案差異 =====")
        try:
            after_files = set(glob.glob(os.path.join(EXPORT_DIR, "*.xls*")))
        except Exception as e:
            after_files = set()
            out.append(f"[警告] 讀取{EXPORT_DIR}事後檔案清單失敗: {e!r}")
        new_files = after_files - before_files
        if new_files:
            out.append(f"[結果] {EXPORT_DIR} 新增了{len(new_files)}個檔案:")
            for f in sorted(new_files):
                out.append(f"    {f}")
                log(f"  新增檔案: {f}")
        else:
            out.append(f"[結果] {EXPORT_DIR} 沒有偵測到新增的.xls/.xlsx檔案——"
                        "可能匯出失敗、還在處理中、或存到別的路徑/副檔名。")
            log(f"  {EXPORT_DIR} 沒有新增檔案。")

        log("\n===== 第10站:關閉閱覽報表視窗(送WM_CLOSE,不用找X按鈕座標) =====")
        close_report_window(rv_win, out)

        out.append("\n[到此為止] 已嘗試自動跑完整套流程:點存Excel->選格式->等匯出->"
                    "按No->關閉報表視窗。麻煩檢查C:\\My Documents\\裡新增的檔案內容"
                    "對不對,以及畫面上有沒有殘留任何沒處理到的對話框。")

    except Exception as e:
        out.append("")
        out.append("!" * 70)
        out.append(f"執行中發生錯誤: {e!r}")
        out.append("!" * 70)
        log(f"[錯誤] {e!r}")
        raise
    finally:
        with open("export_excel_probe_output.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(out))
        log("\n記錄已寫入: export_excel_probe_output.txt")
        log("請把 export_excel_probe_output.txt、export_probe_after_click.png、"
            "還有C:\\My Documents\\有沒有新檔案(檔名)都告訴我,並且留意畫面上"
            "有沒有卡住任何沒處理到的對話框。")


if __name__ == "__main__":
    sys.exit(main())
