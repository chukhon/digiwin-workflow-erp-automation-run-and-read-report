#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
批次跑完佇列裡所有12筆報表工作 -> 各自存成EXCEL檔(2026-09-07,第16輪簡化版)
=======================================================================

背景
----
上一版`part4_batch_export_excel.py`自己包了一層「判斷是Yes/No確認、還是
檔案已存在的覆蓋確認、還是Excel直接跳出」的複雜邏輯(`wait_for_export_
outcome()`),結果使用者實測時**第1筆就撞到「Excel直接跳出」的情況**,而
且原本「自動對這個Excel視窗送WM_CLOSE」的處理方式直接讓ERP噴出Access
violation、拖累第2筆完全開不起來(細節記錄在專案文件的「第15輪」)。

使用者看完之後明確指示:**不要包這層複雜邏輯,直接沿用`part3_export_excel_
probe.py`已經驗證過、簡單可靠的單筆流程就好,根本不用改太多**——那支
腳本本來就只找「Yes」「No」這兩顆按鈕、按下No,從來沒有主動去偵測或觸碰
過任何Excel視窗,所以本來就不會有「誤關Excel導致ERP crash」這個問題。

這一版的做法:**完全比照`part3_export_excel_probe.py`第2~10站的邏輯**
(確認MDI子視窗->找工具列->記錄匯出前檔案->點存Excel->送格式對話框
按鍵->輪詢等Yes/No->按No->比對新增檔案->送WM_CLOSE關閉視窗),只在外面
加一層迴圈處理12筆:

1. 第1筆(佇列清單最上面那筆,預設已選取)開啟前先送兩次TAB把鍵盤焦點
   移到清單本身(使用者第18輪實測發現的關鍵步驟,只有第1筆需要),再送
   ENTER開啟。
2. 第2筆以後,每一筆開始前先送`{DOWN}{ENTER}`移到下一筆並開啟(使用者
   說明的操作方式:處理完一筆、關閉閱覽報表視窗之後,焦點會回到佇列
   工作管理員清單,按DOWN、ENTER就能開下一筆)。
3. 總共12筆,所以是「開第1筆(TAB TAB+ENTER)+11次DOWN+ENTER開第2~12筆」。
4. 每一筆的匯出/確認/關閉邏輯,直接呼叫跟part3幾乎一模一樣的
   `process_one_job()`——一樣只找Yes/No按鈕、按No,不額外分辨「這是
   正常的檢視Excel確認,還是檔案已存在的覆蓋確認」。

**注意(沿用part3原本就有的限制,不是這次新增的簡化副作用)**:因為
`wait_for_yes_no_dialog()`只認Yes/No按鈕本身,不看對話框問的是什麼,如果
某一筆真的跳出「檔案已存在,是否覆蓋?」這類Yes/No對話框,這支腳本一樣會
點No——那不會讓ERP crash(單純不觸碰任何Excel視窗),但意義是「不覆蓋、
取消這次匯出」,那一筆可能就不會產生新檔案。如果使用者後續看到某幾筆
「沒有偵測到新增檔案」,這是需要留意、但不是什麼危險狀況的地方。真的遇到
「沒有跳出Yes/No、而是直接跳出Excel視窗」這種情況(使用者之前提過的
`excel_direct`),`wait_for_yes_no_dialog()`只會照part3原本的邏輯,一直
耐心輪詢等到20分鐘逾時,然後老實回報「等不到Yes/No」並停止整個批次——
不會去猜、不會主動碰那個Excel視窗,安全上跟part3單筆版本完全一致。

第17輪:實測發現DOWN+ENTER似乎沒有真的往下移動,一直卡在第1筆
------------------------------------------------------------
使用者實測上一版(第16輪)之後回報:打開第2筆、第3筆時,視窗標題其實
都還是第1筆的工作代號——log也證實這點(process_one_job重新確認視窗
時,第2筆、第3筆抓到的視窗標題都是「閱覽報表(庫存明細表(IDL)-:
20260101000001-...)」,跟第1筆一模一樣),而且使用者截圖顯示佇列工作
管理員清單裡反白選取的那一列,自始至終都停在最上面那筆(工作代號
20260101000001),完全沒有往下移動過。

目前最可能的根因:close_report_window()送出的WM_CLOSE訊息,用
win32gui.PostMessage()是非同步的,只是把「請關閉」這個訊息塞進視窗的
訊息佇列,不保證視窗馬上真的被銷毀。如果視窗實際上還沒真的關閉,接下來
qm_win.set_focus()不一定真的把鍵盤焦點移回佇列清單本身,送出的DOWN
因此完全沒有作用在清單上,選取列還是停在第1筆,接著送出的ENTER效果就
變成「重新開啟目前仍然選取的第1筆」,而不是「開啟已經往下移動一列的
下一筆」——難怪log裡看到的過渡視窗(設定數值顯示格式等)顯示報表引擎
確實有重新跑一次,但抓到的結果永遠是同一個工作代號。

這一輪的處理(以確認診斷為主,還沒有把握直接下定論、貿然改邏輯):

- 新增wait_until_window_gone():送出WM_CLOSE之後,明確輪詢
  win32gui.IsWindow(handle),確認視窗控制代碼是不是真的失效了(代表
  視窗真的被銷毀),而不是像原本那樣送出去就假設一定成功。如果等15秒
  後視窗控制代碼還有效,直接停止整個批次並回報這個關鍵診斷訊息。
- 把原本一次送出的DOWN+ENTER拆成兩步:先送DOWN,停0.8秒後對qm_win
  存一張截圖(batch_after_down_before_jobN.png),再送ENTER。這樣可以
  直接用眼睛比對「送出DOWN之後,清單裡反白選取的那一列到底有沒有真的
  往下移一格」。

已修好、已編譯通過。

第18輪:使用者自己實測找到真正的根因——鍵盤焦點一開始就沒有真的落在清單上
----------------------------------------------------------------------
使用者自己動手在畫面上實測(不是用腳本),直接看清楚了問題所在,並且用
截圖佐證:佇列工作管理員視窗其實同時有兩個清單——上面「工作代號/作業
名稱/...」那個真正的工作佇列清單(視窗1),跟下面「選擇項目/條件值」
那個報表參數清單(視窗2)。使用者發現:一開啟佇列工作管理員時,雖然畫面上
視窗1的第一列看起來是反白選取的,但鍵盤焦點其實預設落在視窗2上——所以
直接按ENTER確實可以開啟視窗1裡目前選取的那一筆(因為Enter某種程度上
是全域的「開啟目前選取項目」動作),但按DOWN的時候,移動的其實是視窗2
自己清單裡的選取項目(截圖為證:視窗2原本選取「選擇資料日期」,按DOWN
後變成「分類方式」),完全不影響視窗1的選取列——這正是第16、17輪一直
卡在第1筆的真正原因,跟WM_CLOSE有沒有真的關閉視窗其實無關。

使用者提供的解法(已用手動操作驗證過):開啟佇列工作管理員之後,**先按
兩次TAB**,焦點就會確實移到視窗1(清單)本身,之後只要ENTER開啟第1筆,
之後每一筆都是DOWN+ENTER,焦點就會一直留在視窗1上,不會再跑掉——**這個
TAB TAB的動作只有第1筆需要做一次**,不是每一筆都要做。

**這一輪的處理**:

- 第1筆的開啟方式,從原本點擊工具列「開啟檔案」圖示,改成完全用鍵盤:
  送出`{TAB}{TAB}`把焦點移到清單,再送`{ENTER}`開啟第1筆——跟第2筆
  以後的`{DOWN}{ENTER}`一樣走鍵盤這條路,不再混用滑鼠點擊,避免焦點
  再次跑到別的控制項上。
- 拿掉上一輪(第17輪)加的「DOWN、ENTER分兩次送、中間存截圖」那個
  純診斷用的寫法,改回跟第16輪一樣單純的`{DOWN}{ENTER}`——因為真正的
  根因已經由使用者手動確認清楚,不需要再靠截圖臆測。
- `wait_until_window_gone()`(確認WM_CLOSE有沒有真的關閉視窗)保留
  下來當作一般性的安全檢查,不是因為它是這次問題的根因,而是這種
  「送出關閉訊息就假設一定成功」的隱性假設本來就值得長期留著檢查。

已修好、已編譯通過。
第18輪修正後,使用者實測12筆已經可以依序打開(收到第5筆的截圖,確認
成功),不再卡在同一筆。

第19輪:12筆全部完成後,收尾關閉佇列工作管理員、關閉ERP
--------------------------------------------------------
使用者確認批次可以正常跑之後,提出下一步需求:全部12筆抓完報表之後,
要(1)把佇列工作管理員關掉、(2)把ERP本身(MainMenu)也關掉。這兩個
視窗跟閱覽報表視窗不一樣,之前完全沒測過關閉的時候會不會跳出確認
對話框(例如「是否要離開系統?」之類的提示)——照這個專案一貫的安全
原則,新增`close_window_gracefully()`:送出WM_CLOSE之後,用既有的
`wait_until_window_gone()`確認視窗是不是真的關閉了;如果沒關閉,再看
看有沒有跳出新視窗(可能是確認對話框),只記錄下來(class/文字),
**不會自動亂點**,交給使用者自己確認畫面、手動處理——跟這支腳本一路
以來「看不懂、沒把握的狀況就停下來,不猜」的原則一致。

只有在12筆全部順利處理完成(`all_completed`為True且`jobs_done`剛好
等於12)時,才會執行這個收尾動作——依序關閉佇列工作管理員、再關閉
ERP主畫面(如果佇列工作管理員沒有確實關閉,就不會繼續嘗試關ERP)。
如果批次中途因為任何原因停下來,會保留畫面原狀,不會自動關閉任何
東西,讓使用者可以自己確認目前卡在哪裡。

已修好、已編譯通過。
**這一版尚未經使用者實機驗證**——這是第一次讓腳本主動關閉佇列工作
管理員/ERP本身,需要實際測試才知道關閉過程中會不會跳出什麼對話框、
`close_window_gracefully()`的判斷邏輯夠不夠用。

第20輪(2026-09-08):把main()的批次邏輯抽成run_batch(),供整合腳本重用
------------------------------------------------------------------------
使用者要把Part1(清除/複製Excel)、Stage A(run_all_reports.py,送出全部
報表查詢)、Stage B(這支腳本,批次匯出成Excel+收尾關閉)三段串成一支
`run_full_pipeline.py`一次跑完,而且只登入ERP一次(Stage A、Stage B
共用同一個main_win,不是分別各自登入/登出)。

為了不重寫這支腳本已經驗證過的批次邏輯,把原本`main()`裡「開佇列工作
管理員->迴圈處理12筆->收尾關閉」這一整段,抽成獨立函式`run_batch(main_win,
out)`(回傳True/False代表是否全部完成+收尾成功)。`main()`本身改成單純
呼叫`login()`拿到`main_win`,再呼叫`run_batch(main_win, out)`,其餘
(寫入`batch_export_excel_output.txt`、印結尾提示)完全不變——單獨執行
`python part4_batch_export_excel.py`的行為跟這次改之前完全一樣,只是
內部多了一層函式邊界,方便`run_full_pipeline.py`直接`from
part4_batch_export_excel import run_batch`重用,不必複製貼上這段邏輯、
也不會讓兩份拷貝之後各自修各自的、漸漸不同步。

執行需求
--------
    python part4_batch_export_excel.py
"""

import glob
import os
import sys
import time

import win32con
import win32gui

from part2_inspect import find_by_class, log, login
from part2_reportview_probe import (
    MM_QUEUE_BTN,
    QM_CLASS,
    all_top_windows,
    click_toolbar_icon,
    save_shot,
    try_dismiss_error_dialog,
    wait_new_window,
)
from part3_export_excel_probe import (
    EXCEL_BTN_REL_X,
    EXCEL_DLG_CLASS,
    EXPORT_DIR,
    EXPORT_WAIT_TIMEOUT,
    close_report_window,
    find_new_window_by_class,
    reacquire_rv_win,
    wait_for_yes_no_dialog,
)

# 今天run_all_reports.py總共產生了12筆工作。
NUM_JOBS_TO_PROCESS = 12


def wait_until_window_gone(handle, out, job_label, timeout=15.0, poll=0.5):
    """2026-09-07第17輪新增:使用者實測發現,批次跑第2筆、第3筆時,重新
    抓到的閱覽報表視窗標題**一直都是第1筆的工作代號**,懷疑是`close_
    report_window()`送出的WM_CLOSE其實沒有真的把視窗關掉/銷毀,導致
    後面的`reacquire_rv_win()`每次都重新抓到同一個還活著的舊視窗。

    這裡在送出WM_CLOSE之後,明確輪詢`win32gui.IsWindow(handle)`,確認
    視窗控制代碼是不是真的失效(代表視窗真的被銷毀了),而不是像原本
    那樣送出WM_CLOSE後就假設它一定會關閉、直接往下一筆走。回傳True代表
    視窗確實消失了,False代表逾時後視窗仍然存在(這是這次要抓的關鍵
    診斷資訊)。"""
    start = time.time()
    while time.time() - start < timeout:
        try:
            still_there = win32gui.IsWindow(handle)
        except Exception:
            still_there = False
        if not still_there:
            out.append(f"[{job_label}] 送出WM_CLOSE後,約{time.time() - start:.1f}秒視窗確實消失了。")
            return True
        time.sleep(poll)
    out.append(f"[{job_label}] 【重要診斷】送出WM_CLOSE後等了{timeout}秒,視窗控制代碼仍然有效"
                "(win32gui.IsWindow()回傳True)——很可能WM_CLOSE沒有真的把視窗關掉/銷毀,"
                "這可能就是後面重新抓到的視窗一直是同一份工作的原因。")
    return False


def close_window_gracefully(win, out, label, wait_seconds=10):
    """2026-09-07第19輪新增:12筆全部跑完之後,使用者要求收尾時依序關掉
    佇列工作管理員、再關掉ERP本身(MainMenu)。這兩個視窗跟閱覽報表視窗
    不一樣,我們從來沒實際測過關閉之後會不會跳出確認對話框(例如「是否
    要離開系統?」之類的提示)——依這個專案一貫的安全原則,送出WM_CLOSE
    之後只確認視窗是否真的關閉,如果沒關閉、或期間跳出了新視窗,一律
    只記錄下來、不自動亂點,交給使用者自己確認畫面、手動處理。"""
    handle = win.handle
    try:
        known_before = {w.handle for w in all_top_windows()}
    except Exception:
        known_before = set()
    try:
        win32gui.PostMessage(handle, win32con.WM_CLOSE, 0, 0)
        out.append(f"[{label}] 已送出WM_CLOSE訊息。")
    except Exception as e:
        out.append(f"[{label}] 送WM_CLOSE失敗: {e!r}")
        return False

    closed = wait_until_window_gone(handle, out, label, timeout=wait_seconds)
    if closed:
        out.append(f"[{label}] 視窗已確實關閉。")
        return True

    try:
        new_wins = [w for w in all_top_windows() if w.handle not in known_before]
    except Exception:
        new_wins = []
    if new_wins:
        out.append(f"[{label}] 視窗還沒關閉,而且期間跳出了{len(new_wins)}個新視窗"
                    "(可能是確認對話框,例如「是否要離開?」之類的提示)——"
                    "為安全起見不自動點擊,請自己確認畫面、手動處理:")
        for w in new_wins:
            try:
                out.append(f"    class={w.class_name()!r} text={w.window_text()!r}")
            except Exception:
                pass
    else:
        out.append(f"[{label}] 視窗還沒關閉,期間也沒有跳出新視窗,請自己確認畫面。")
    return False


def process_one_job(qm_win, rv_win, out, job_label):
    """處理「目前已經開啟」的這一筆閱覽報表視窗——完全比照
    part3_export_excel_probe.py的main()第3~10站:重新確認MDI子視窗->
    找工具列->記錄匯出前檔案->點存Excel->送格式對話框按鍵->輪詢等
    Yes/No確認->按No->比對新增檔案->送WM_CLOSE關閉視窗。

    回傳True代表這一筆順利跑完,可以繼續下一筆;False代表遇到問題,
    呼叫端應該停止整個批次。"""
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
        out.append(f"[{job_label}] 多次重新搜尋後還是找不到閱覽報表視窗,停止整個批次。")
        return False
    rv_win = rv_win_fresh
    out.append(f"[{job_label}] 視窗標題: {rv_win.window_text()!r}")

    tbs = find_by_class(rv_win, "TDSUCToolBar.UnicodeClass", "ToolBar")
    if not tbs:
        out.append(f"[{job_label}] 找不到工具列,停止整個批次。")
        return False
    tb = tbs[0]
    tb_rect = tb.rectangle()
    y_rel = (tb_rect.bottom - tb_rect.top) // 2

    try:
        before_files = set(glob.glob(os.path.join(EXPORT_DIR, "*.xls*")))
    except Exception as e:
        before_files = set()
        out.append(f"[{job_label}] 警告:讀取匯出前檔案清單失敗: {e!r}")

    known = {w.handle for w in all_top_windows()}
    click_toolbar_icon(rv_win, (EXCEL_BTN_REL_X, y_rel), out, f"{job_label} 存成EXCEL檔")

    excel_dlg = find_new_window_by_class(known, EXCEL_DLG_CLASS, qm_win, out, timeout=10)
    if excel_dlg is None:
        out.append(f"[{job_label}] 等不到格式設定視窗,停止整個批次。")
        return False
    out.append(f"[{job_label}] 找到格式設定視窗: text={excel_dlg.window_text()!r}")
    try:
        excel_dlg.set_focus()
    except Exception:
        pass
    time.sleep(0.3)
    excel_dlg.type_keys("{RIGHT}{TAB}{TAB}{ENTER}")
    out.append(f"[{job_label}] 已送出格式對話框按鍵序列。")

    dialog_btns, waited = wait_for_yes_no_dialog(qm_win, out)
    if dialog_btns is None:
        out.append(f"[{job_label}] 等了{EXPORT_WAIT_TIMEOUT}秒還是沒看到Yes/No確認對話框,"
                    "可能卡住了或跳出了非預期的視窗(例如直接跳出Excel視窗,沒有經過"
                    "確認對話框),停止整個批次,請自己確認畫面。")
        return False
    out.append(f"[{job_label}] 等了約{waited:.0f}秒後,匯出完成,找到Yes/No確認對話框。")
    try:
        dialog_btns["no"].click()
        out.append(f"[{job_label}] 已點擊「No」按鈕(不開啟Excel)。")
    except Exception as e:
        out.append(f"[{job_label}] 點擊No按鈕失敗: {e!r},停止整個批次,請自己確認畫面。")
        return False
    time.sleep(1.0)

    save_shot(rv_win, f"batch_job{job_label.replace('批次第', '').replace('筆', '')}_after_click.png", out)

    try:
        after_files = set(glob.glob(os.path.join(EXPORT_DIR, "*.xls*")))
    except Exception as e:
        after_files = set()
        out.append(f"[{job_label}] 警告:讀取匯出後檔案清單失敗: {e!r}")
    new_files = after_files - before_files
    if new_files:
        out.append(f"[{job_label}] 新增了{len(new_files)}個檔案:")
        for f in sorted(new_files):
            out.append(f"    {f}")
    else:
        out.append(f"[{job_label}] 沒有偵測到新增檔案,請留意(可能是這一筆匯出前"
                    "檔案就已經存在、剛剛跳出的Yes/No其實是覆蓋確認,按No變成取消"
                    "匯出,不是本來預期的「檢視Excel資料?」確認)。")

    rv_handle = rv_win.handle
    close_report_window(rv_win, out)
    closed_ok = wait_until_window_gone(rv_handle, out, job_label)
    if not closed_ok:
        out.append(f"[{job_label}] 因為視窗疑似沒有真的被關閉/銷毀,為安全起見停止整個批次"
                    "(避免下一筆又重複開到同一份工作)——請自己確認畫面上這個閱覽報表視窗"
                    "是不是還在、佇列工作管理員清單目前停在哪一筆。")
        return False
    return True


def run_batch(main_win, out):
    """開佇列工作管理員 -> 批次處理NUM_JOBS_TO_PROCESS筆 -> 全部成功才收尾
    關閉佇列工作管理員/ERP。

    2026-09-08第20輪抽出:這是原本`main()`的批次邏輯本體,抽成獨立函式後,
    可以被這支腳本自己的`main()`呼叫(獨立執行,`login()`是這支腳本自己
    開的新ERP),也可以被`run_full_pipeline.py`呼叫(整合流程裡,main_win
    是run_full_pipeline.py自己登入、Stage A已經用過的同一個ERP session,
    不會重新登入)——邏輯完全沒變,只是抽出來給兩邊共用,避免複製貼上。

    回傳True代表全部NUM_JOBS_TO_PROCESS筆都成功處理完、且收尾關閉也做了;
    False代表批次中途停止,或收尾關閉沒有完全成功。"""
    jobs_done = 0
    log("\n===== 第1站:點「顯示報表工作序列」 =====")
    known = {w.handle for w in all_top_windows()}
    click_toolbar_icon(main_win, MM_QUEUE_BTN, out, "MainMenu顯示報表工作序列")
    qm_win, _ = wait_new_window(known, out, "佇列工作管理員", timeout=25, want_class=QM_CLASS)
    if qm_win is None:
        out.append("[結果] 沒等到佇列工作管理員視窗,停止。")
        return False
    log(f"  開啟成功: {qm_win.window_text()!r}")
    try:
        qm_win.set_focus()
    except Exception:
        pass
    time.sleep(2.5)

    all_completed = True
    for i in range(1, NUM_JOBS_TO_PROCESS + 1):
        job_label = f"批次第{i}筆"
        log(f"\n===== {job_label}:開啟這一筆工作 =====")
        try:
            qm_win.set_focus()
        except Exception:
            pass
        time.sleep(1.0)

        known_open = {w.handle for w in all_top_windows()}
        if i == 1:
            # 2026-09-07第18輪:使用者自己實測找到真正的根因——一開啟
            # 佇列工作管理員時,畫面上清單(視窗1)雖然看起來是反白
            # 選取的,但鍵盤焦點其實可能落在下面「選擇項目/條件值」
            # 那個參數清單(視窗2)上。之前用滑鼠點「開啟檔案」工具列
            # 圖示可以正常開第1筆,是因為工具列點擊本來就不需要鍵盤
            # 焦點在清單上;但也正因為這樣,後面改用DOWN+ENTER移到
            # 第2筆時,DOWN其實作用在視窗2、不是視窗1的清單,才會一直
            # 卡在第1筆。使用者實測確認的解法:開啟佇列工作管理員之後,
            # 先按兩次TAB把焦點移到清單本身(只有第1次需要),之後
            # 全程都用鍵盤操作(不再用滑鼠點工具列開檔案),DOWN/ENTER
            # 就會一直正確作用在清單上。
            qm_win.type_keys("{TAB}{TAB}")
            out.append(f"[{job_label}] 已送出兩次TAB,把焦點移到佇列清單本身"
                        "(使用者實測確認的做法,只有第1筆需要)。")
            time.sleep(0.3)
            qm_win.type_keys("{ENTER}")
            out.append(f"[{job_label}] 已送出ENTER,開啟第1筆工作。")
        else:
            qm_win.type_keys("{DOWN}{ENTER}")
            out.append(f"[{job_label}] 已送出" + "{DOWN}{ENTER}" + "。")

        rv_win, _ = wait_new_window(known_open, out, f"{job_label} 閱覽報表", timeout=45,
                                     want_class="TfrmReportViewer.UnicodeClass")
        if rv_win is None:
            rv_win, _ = wait_new_window(known_open, out, f"{job_label} 閱覽報表(備案)", timeout=15)
        if rv_win is None:
            out.append(f"[{job_label}] 找不到閱覽報表視窗,停止整個批次。")
            all_completed = False
            break

        time.sleep(5.0)
        try_dismiss_error_dialog(known_open, out, wait=1.0)

        ok = process_one_job(qm_win, rv_win, out, job_label)
        jobs_done += 1
        if not ok:
            all_completed = False
            break
        time.sleep(2.0)

    out.append(f"\n[批次結果] 這次共處理了{jobs_done}筆工作(從第1筆開始)。")

    if all_completed and jobs_done == NUM_JOBS_TO_PROCESS:
        log("\n===== 收尾:12筆全部處理完成,關閉佇列工作管理員、關閉ERP =====")
        out.append("\n[收尾] 12筆全部順利完成,開始關閉佇列工作管理員、關閉ERP。")
        # 2026-09-11:使用者實測發現「佇列工作管理員」其實關閉了,只是花的時間比原本
        # 預設的10秒更久(等到第11版才真正觸發win32gui.IsWindow()變False),導致
        # 誤判成「沒確實關閉」而不敢繼續關ERP。這裡把這個視窗的等待時間拉長到20秒,
        # 給多一點緩衝;ERP主畫面(下面那個close_window_gracefully呼叫)維持預設
        # 10秒不變,之後如果也遇到類似的誤判再一併調整。
        qm_closed = close_window_gracefully(qm_win, out, "關閉佇列工作管理員", wait_seconds=20)
        if qm_closed:
            time.sleep(1.5)
            main_closed = close_window_gracefully(main_win, out, "關閉ERP主畫面(MainMenu)")
        else:
            out.append("[收尾] 因為佇列工作管理員沒有確實關閉,先不繼續嘗試關閉ERP本身,"
                        "請自己確認畫面狀況。")
            main_closed = False
        return bool(qm_closed and main_closed)
    else:
        out.append("\n[收尾] 因為這次批次沒有完整跑完12筆,先不自動關閉佇列工作管理員/ERP,"
                    "保留畫面讓你自己確認狀況。")
        return False


def main():
    out = []
    try:
        app, main_win = login(out)
        run_batch(main_win, out)
    except Exception as e:
        out.append("")
        out.append("!" * 70)
        out.append(f"執行中發生錯誤: {e!r}")
        out.append("!" * 70)
        log(f"[錯誤] {e!r}")
        raise
    finally:
        with open("batch_export_excel_output.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(out))
        log("\n記錄已寫入: batch_export_excel_output.txt")
        log(f"請把 batch_export_excel_output.txt、batch_job*_after_click.png、"
            f"還有{EXPORT_DIR}裡新增的檔案清單都告訴我,並且留意畫面上有沒有"
            "卡住任何沒處理到的對話框。")


if __name__ == "__main__":
    sys.exit(main())
