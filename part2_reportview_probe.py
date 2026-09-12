#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
閱覽報表(ReportView)視窗偵察 — 找「存成EXCEL檔」「關閉」等按鈕座標
=======================================================================

背景
----
`part2_stageb_chain_probe.py`已經走通到「閱覽報表視窗打開」這一步(登入 →
點MainMenu工具列開佇列工作序列 → 點佇列工作管理員工具列「開啟檔案」),但
上次那一輪閱覽報表視窗一開就跳出Windows錯誤訊息「Access violation ...
module 'ReportViewerC.dll'」,報表內容是空的,控制項清單裡也只找到2個
不相關的控制項(TDSUCMemo、空的TDSUCTreeView),**沒有看到工具列**——不確定
是這個視窗本來就沒有工具列,還是因為Access violation導致視窗沒有完全初始化。

這支腳本要做的事
----------------
跟之前MainMenu/佇列工作管理員工具列偵察一樣的做法(owner-drawn、沒有
tooltip,只能用截圖量座標),延伸到閱覽報表視窗:
1. 走到閱覽報表視窗打開為止(沿用chain_probe的路線)。
2. 如果跳出Windows錯誤訊息視窗(class通常是`#32770`),嘗試按掉它(找
   「確定」或「OK」按鈕點擊,或送ESC/Enter),讓閱覽報表視窗有機會完整
   顯示,不要讓這個錯誤訊息擋住後面的偵察。
3. 不管報表內容有沒有正常顯示,都把閱覽報表視窗**目前的**完整控制項清單
   +全螢幕截圖+(如果找得到工具列)放大編號截圖存下來——如果這次工具列
   有出現,就能比照之前MainMenu/佇列工作管理員的方式,量出「存成EXCEL檔」
   「關閉」等按鈕的相對座標;如果還是只有2個不相關控制項、沒有工具列,
   這也是重要資訊,代表Access violation不只是跳出一個訊息框而已,而是
   報表內容真的沒有正常初始化,需要先解決這個更根本的問題才能繼续往下走。

**不會**做轉Excel、不會刪除任何工作、不會關掉閱覽報表視窗本身(只會嘗試
關掉錯誤訊息視窗,如果有跳出來的話)。

狀態(2026-09-07更新):9/7實測這次**沒有跳出Access violation錯誤訊息**,
使用者自己截圖也確認報表視窗有正常顯示工具列跟報表資料(庫存明細表(IDL))
——但這支腳本自己抓到的控制項清單卻是空的、截圖也失敗(`capture_as_image()`
回傳`None`)。研判是這份報表資料量比較大,「開啟檔案」之後畫面可能先跳出
一個內容還沒畫好的過渡狀態,子控制項要等資料查完、畫面畫出來才會補齊;
如果中途視窗被整個關掉重開(不是同一個控制代碼),沿用第2站當時抓到的
舊物件參照就會撲空。已加上重新確認/重新搜尋視窗的邏輯(`reacquire_rv_win()`),
在真正診斷之前多等幾秒、確認視窗控制代碼還存在且有實際大小,拿最新的
視窗物件來用。

第2輪追測又發現另一個問題:第2站原本用「抓第一個冒出來、不是雜訊視窗」
這種通用邏輯去等閱覽報表視窗,結果這次抓到一個完全不相關的小過渡視窗
(class='TForm',空白,400x300),導致後面不管怎麼重新確認都在對著錯的
視窗打轉。已改成優先明確指定`want_class="TfrmReportViewer.UnicodeClass"`
去等(這是第1次測試已經確認過的正確class name),等不到才退回通用邏輯
當備案。

第3輪:使用者先把畫面上所有閱覽報表視窗關掉、確保是乾淨狀態再測,排除掉
「同一個工作重複開視窗」這個變數之後,結果還是一樣——第2站正確抓到
class/標題都對的視窗,但幾秒後控制代碼就失效、重新搜尋也找不到,視窗
真的整個消失了。這代表問題不在偵察邏輯抓錯視窗,而是**這份報表資料量大,
畫面內容還在陸續繪製時,只要我們的腳本對這個視窗做任何程式互動(哪怕只是
看似無害的`.set_focus()`),就可能干擾到ReportViewerC.dll內部還在跑的流程,
間接觸發那個一直沒解決的Access violation,讓視窗整個關掉**——使用者自己
全程用滑鼠手動操作、完全不觸碰程式時,報表可以正常顯示、不會消失。已移除
`rv_win.set_focus()`這個多餘的互動,改成先完全不碰視窗、被動等5秒讓報表
引擎有充分時間畫完內容,才開始做只讀性質的診斷。

第4輪:拿掉`.set_focus()`、多等5秒之後,結果還是一模一樣——原本的控制
代碼確實完全失效(`win32gui.IsWindow()`回傳`False`,不是還在但抓不到),
重新搜尋桌面所有頂層視窗也找不到同class的視窗。這排除了「自動化互動害它
掛掉」的猜測(這次全程沒碰它)。改指向另一個方向:視窗標題「佇列工作
控制台[...] - [閱覽報表(...)]」是Windows MDI子視窗被放大時的經典標題
格式——很可能這個視窗一開始短暫是個一般頂層視窗,之後被`SetParent()`
重新掛到佇列工作管理員(qm_win)底下變成真正的MDI子視窗,一旦完成這個
轉換,只列頂層視窗的`Desktop().windows()`自然再也找不到它,但它其實還
活著、還在畫面上。已在原本的頂層搜尋之外,新增一條路:改在`qm_win.
descendants()`(往下遍歷所有子視窗,不限頂層)裡找同class name的視窗。

**第4輪使用者實測驗證成功**:這次確實在`qm_win`底下的子視窗裡找到了
(log:「頂層視窗清單裡沒有,但在qm_win底下的子視窗裡找到了」),截圖也
成功存下來,而且**這次真的找到了工具列**(`TDSUCToolBar.UnicodeClass`,
1個),量出座標的刻度圖也存好了。MDI子視窗的推測證實正確。

第5輪:使用者的截圖意外拍到一個「刪除格式」的黃色提示框浮在報表內容上方、
緊貼工具列下方——這個位置符合Windows標準tooltip的浮現位置(貼著滑鼠游標
下方),不是StatusBar提示文字的位置(StatusBar在視窗最下面)。代表**這個
工具列(跟MainMenu/佇列工作管理員那兩個不一樣)其實是有tooltip的**——
先前「這個ERP完全沒有tooltip」的結論只驗證過那兩個owner-drawn工具列,
沒驗證過這個。既然有tooltip,就不用猜圖示形狀,新增`sweep_toolbar_tooltips()`
——把滑鼠依序移到工具列每個x座標懸停,讀取有沒有跳出`tooltips_class32`
視窗,把每個座標對應的真正按鈕名稱記錄下來,程式化找出「存成EXCEL檔」
「關閉」等按鈕的位置,不用再靠人眼比對圖示。

第6輪:使用者實測`sweep_toolbar_tooltips()`,結果**一個tooltip都沒掃到**,
跟第5輪那張「刪除格式」截圖的證據矛盾。同時使用者也提供一張標註截圖,用
藍色圈圈圈出閱覽報表視窗**標題列右上角的「關閉(X)」按鈕**,要求這個
位置也要一併掃tooltip。這裡一次處理兩件事:

1. 懷疑0個tooltip的原因:第3輪之後為了排除「我們的互動害它崩潰」這個
   (後來證明是錯的)理論,一直沒有對rv_win呼叫`set_focus()`。但Windows
   有些自訂/owner-drawn控制項的tooltip邏輯,只有在**視窗是作用中視窗**時
   才會顯示提示框(這跟WM_MOUSEMOVE本身無關,而是控制項內部可能會判斷
   `GetActiveWindow()`/`GetForegroundWindow()`再決定要不要顯示hint)。
   既然第4輪已經證實set_focus()本身不會造成Access violation(這是MDI
   子視窗的問題,跟互動與否無關),這裡在開始掃tooltip之前先補回
   `rv_win.set_focus()`,看看是不是這個造成之前掃不到。同時把原本掃描
   邏輯抽成共用的`hover_and_read_tooltip()`,並把`hover_wait`從1.0秒
   拉長到1.3秒,增加保險。
2. 新增`sweep_titlebar_buttons()`:標題列的關閉/最大化/最小化按鈕屬於
   視窗的非客戶區(non-client area),不是`TDSUCToolBar`底下的子控制項,
   沒辦法像工具列一樣直接抓子控制項的rectangle量座標,只能比照工具列的
   做法,在視窗矩形右上角一小塊範圍內(多試幾個y偏移,因為不確定標題列
   實際厚度)逐點掃描、讀tooltip文字。

第7輪:使用者測了補上`set_focus()`+拉長懸停時間那一版,**工具列跟標題列
兩邊都還是0個tooltip**,排除了「視窗不是作用中視窗」這個猜測。回頭想:
一直以來都只找`class='tooltips_class32'`(Windows COMCTL32標準tooltip
視窗),但`TDSUCToolBar`是Delphi VCL的owner-drawn自訂控制項,很可能根本
不是用這套標準機制顯示提示,而是VCL自己的`THintWindow`(舊版Delphi這個
提示框元件的window class name字面上就是`'THintWindow'`)或其他自訂彈出
視窗——一直用錯的class name去找,當然永遠是0個。

**改法**:不再限定class name。`_find_new_text_windows()`改成單純比對
「懸停前/懸停後」兩次桌面頂層視窗handle快照的差集,只要是新冒出來、
而且有文字內容的視窗都算候選,連真正的class name都一併記錄下來,寫進
輸出檔(`[候選彈出視窗] ... class=... text=...`)。這樣就算這次挑到的
「代表文字」還是不準,至少能從候選清單裡看到真正的提示框class name是
什麼,之後才能精準鎖定,不用繼續憑空猜測。已修好。**這一版尚未經實機驗證。**

執行需求
--------
    python part2_reportview_probe.py
"""

import sys
import time

import pipeline_flags
from part2_inspect import dump_control_tree, find_by_class, log, login, safe_all_windows

# MainMenu工具列:「顯示報表工作序列」= 第3顆圖示(使用者紅圈確認),相對工具列座標
MM_QUEUE_BTN = (297, 13)
# 佇列工作管理員工具列:「開啟檔案」= 第3顆圖示(黃色開啟資料夾圖示)
QM_OPEN_BTN = (90, 16)

QM_CLASS = "DSC_Conductor_QueueManager.UnicodeClass"

# Windows標準訊息框/錯誤對話框的class name
MSGBOX_CLASS = "#32770"


def all_top_windows():
    # 2026-09-07改用part2_inspect.safe_all_windows():直接呼叫
    # Desktop(backend="win32").windows()實測發現偶爾會因為某個過渡視窗剛好
    # 在枚舉完、包裝前的瞬間被關掉而整支腳本crash(InvalidWindowHandle),
    # 跟這支腳本本身要找的視窗完全無關。safe_all_windows()會在遇到這種
    # race condition時重試整次枚舉,詳見part2_inspect.py裡的說明。
    return safe_all_windows()


def save_shot(win, path, out):
    if not pipeline_flags.SCREENSHOTS_ENABLED:
        out.append(f"[截圖已略過(SCREENSHOTS_ENABLED=False): {path}]")
        return None
    try:
        img = win.capture_as_image()
        img.save(path)
        out.append(f"[截圖: {path}]")
        log(f"  已存 {path}")
        return img
    except Exception as e:
        out.append(f"[截圖失敗 {path}: {e!r}]")
        log(f"  [警告] 截圖失敗 {path}: {e!r}")
        return None


def save_numbered_toolbar(win, img, out, tag):
    """把視窗裡每個工具列切出來、放大4倍、每30px畫一條刻度線並標數字。"""
    if img is None:
        return 0
    try:
        from PIL import Image as PILImage, ImageDraw
        wr = win.rectangle()
        tbs = [c for c in win.descendants()
               if c.class_name() == "TDSUCToolBar.UnicodeClass"]
        for n, tb in enumerate(tbs):
            r = tb.rectangle()
            box = (max(0, r.left - wr.left), max(0, r.top - wr.top),
                   min(img.width, r.right - wr.left),
                   min(img.height, r.bottom - wr.top))
            if box[2] <= box[0] or box[3] <= box[1]:
                continue
            S = 4
            crop = img.crop(box).resize(((box[2] - box[0]) * S, (box[3] - box[1]) * S))
            canvas_h = crop.height + 34
            canvas = PILImage.new("RGB", (crop.width, canvas_h), (255, 255, 255))
            canvas.paste(crop, (0, 34))
            d = ImageDraw.Draw(canvas)
            step = 30
            x = 0
            while x * S < crop.width:
                px = x * S
                d.line([(px, 34), (px, canvas_h)], fill=(255, 0, 0), width=1)
                d.text((px + 2, 10), str(x), fill=(200, 0, 0))
                x += step
            path = f"{tag}_toolbar{n}_ruler.png"
            canvas.save(path)
            out.append(f"[工具列刻度圖: {path}  工具列rect={r}]")
            log(f"  已存 {path}")
        return len(tbs)
    except Exception as e:
        out.append(f"[工具列刻度圖失敗: {e!r}]")
        return 0


def _snapshot_handles():
    """回傳目前桌面所有頂層視窗的handle集合,當作「這個時間點還沒出現」
    的基準快照,用來之後做差集找出新冒出來的視窗。"""
    try:
        return {w.handle for w in safe_all_windows()}
    except Exception:
        return set()


def _find_new_text_windows(baseline_handles):
    """找出目前所有頂層視窗裡,baseline快照沒有、而且有文字內容的新視窗,
    回傳[(class_name, text, handle), ...]。**不限定class name。**

    2026-09-07第7輪新增,背景:前兩輪(第5、6輪)都只找class=
    'tooltips_class32'的視窗,結果連續兩次、兩個位置(工具列+標題列)都是
    0個,但使用者的截圖證明真的有「刪除格式」提示文字浮現過。懷疑這個
    owner-drawn工具列(TDSUCToolBar,Delphi VCL自訂控制項)根本不是用
    Windows標準COMCTL32的tooltip機制,而是VCL自己的`THintWindow`(這是
    Delphi VCL內建的提示框元件,舊版Delphi常見的window class name就是
    'THintWindow'字面本身,不是'tooltips_class32')或其他自訂彈出視窗——
    class name的猜測一直錯,与其繼續猜,不如改成**不限定class**,只要是
    baseline快照裡沒有、而且有文字的新視窗都當作候選記錄下來(連class
    name一起記),這樣至少能找出真正的class name是什麼,之後才能精準鎖定。"""
    hits = []
    try:
        for w in safe_all_windows():
            if w.handle in baseline_handles:
                continue
            try:
                t = (w.window_text() or "").strip()
            except Exception:
                continue
            if not t:
                continue
            try:
                cls = w.class_name()
            except Exception:
                cls = "?"
            hits.append((cls, t, w.handle))
    except Exception:
        pass
    return hits


# 已知/推測可能的提示框class name,依優先順序挑選——沒有命中就退回用第一個
# 有文字的候選(不限class),當作診斷用途。
_HINT_CLASS_PRIORITY = ("tooltips_class32", "THintWindow")


def hover_and_read_tooltip(x, y, away_dxdy=(0, -80), hover_wait=1.3, settle=0.05):
    """把滑鼠先移到(x,y)附近但明顯不同的位置,再移回(x,y),停留hover_wait
    秒後讀有沒有跳出新的、有文字的彈出視窗。先移開再移回是為了避免Windows
    因為滑鼠沒真的離開過同一個控制項而不重複顯示同一個提示。

    2026-09-07第6輪抽成共用函式(原本寫在sweep_toolbar_tooltips()裡面),
    第7輪改成不限定class name(見_find_new_text_windows()的說明),回傳
    (text, hits)——text是挑出來的「這次的提示文字」(空字串代表沒找到),
    hits是這次看到的所有候選(class, text, handle),讓呼叫端可以把還沒見
    過的class name記錄下來,方便之後鎖定真正的提示框class。"""
    from pywinauto import mouse

    away = (x + away_dxdy[0], max(y + away_dxdy[1], 0))
    try:
        mouse.move(coords=away)
        time.sleep(settle)
    except Exception:
        return "", []
    baseline = _snapshot_handles()
    try:
        mouse.move(coords=(x, y))
    except Exception:
        return "", []
    time.sleep(hover_wait)
    hits = _find_new_text_windows(baseline)
    if not hits:
        return "", []
    for cls, t, _h in hits:
        if cls in _HINT_CLASS_PRIORITY:
            return t, hits
    # 沒有命中已知的class,退回用第一個有文字的候選(診斷用途)。
    return hits[0][1], hits


def sweep_toolbar_tooltips(tb, out, step=15, hover_wait=1.3):
    """把滑鼠依序移到工具列每個x座標懸停,讀取當下有沒有跳出Windows標準
    tooltip視窗(class='tooltips_class32'),記錄每個座標對應的提示文字。

    2026-09-07新增,背景:使用者的截圖意外拍到一個「刪除格式」的黃色提示框
    浮在報表內容上方、緊貼在工具列下方——這個位置(貼著滑鼠游標下方,不是
    視窗最下面的StatusBar位置)符合Windows標準tooltip的浮現位置,不是
    StatusBar提示文字。這代表**閱覽報表視窗自己的這個工具列,跟MainMenu/
    佇列工作管理員那兩個工具列不一樣,是有tooltip的**(先前結論「這個ERP
    完全沒有tooltip」只驗證過那兩個,沒驗證過這個)。既然有tooltip,就不用
    靠猜圖示形狀,可以直接程式化讀出每顆按鈕真正的名稱。

    第6輪:第一版實測掃到0個tooltip,懷疑跟視窗當時不是作用中視窗有關
    (呼叫端在main()裡,掃描前已補上rv_win.set_focus())。這裡也把
    hover_wait從1.0秒拉長到1.3秒,並改用共用的hover_and_read_tooltip()。

    第7輪:補了set_focus()之後還是0個,改成hover_and_read_tooltip()不再
    限定class name(見該函式說明),這裡額外把每個「第一次見到」的候選
    (class, text)組合記錄進out,就算這次抓到的候選不算數,至少能知道
    真正的提示框class name是什麼。"""
    r = tb.rectangle()
    y = (r.top + r.bottom) // 2
    results = []
    last_text = None
    seen_candidates = set()
    x = r.left + 5
    while x < r.right:
        text, hits = hover_and_read_tooltip(x, y, hover_wait=hover_wait)
        for cls, t, _h in hits:
            key = (cls, t)
            if key not in seen_candidates:
                seen_candidates.add(key)
                rel_x0 = x - r.left
                out.append(f"    [候選彈出視窗] 相對x={rel_x0} class={cls!r} text={t!r}")
                log(f"  [候選彈出視窗] 相對x={rel_x0} class={cls!r} text={t!r}")
        rel_x = x - r.left
        if text and text != last_text:
            results.append({"rel_x": rel_x, "text": text})
            log(f"  相對x={rel_x:>4}  tooltip={text!r}")
            out.append(f"  相對x={rel_x:>4}  tooltip={text!r}")
            last_text = text
        elif not text:
            last_text = None
        x += step
    return results


def sweep_titlebar_buttons(win, out, hover_wait=1.3):
    """2026-09-07第6輪新增,回應使用者指令:「你也要用tooltip 掃藍色圈圈的
    位置」——使用者在標註截圖上用藍色圈圈標出閱覽報表視窗**標題列右上角的
    「關閉(X)」按鈕**,要求這個位置也要一併掃tooltip,不能只掃主工具列。

    這個按鈕屬於視窗的非客戶區(non-client area),不是TDSUCToolBar底下的
    子控制項,沒辦法像工具列一樣直接抓子控制項的rectangle量出精確座標。
    做法比照工具列:用win.rectangle()的右上角往回推,在一小塊範圍內
    (多試幾個垂直偏移,因為不確定這個縮放環境下標題列實際厚度)逐點掃描、
    讀tooltip文字——如果掃到「關閉」之類的文字,就能同時確認關閉按鈕的
    精確座標,不用再靠肉眼比對截圖。

    第7輪:改用不限class的hover_and_read_tooltip()(回傳(text, hits)),
    同樣把「第一次見到」的候選(class, text)記錄下來。"""
    r = win.rectangle()
    results = []
    last_text = None
    seen_candidates = set()
    for dy in (10, 16, 22, 28, 34):
        y = r.top + dy
        if y >= r.bottom:
            continue
        x = r.right - 8
        x_limit = max(r.right - 160, r.left)
        while x > x_limit:
            text, hits = hover_and_read_tooltip(x, y, hover_wait=hover_wait)
            rel_x = x - r.left
            rel_y = y - r.top
            for cls, t, _h in hits:
                key = (cls, t)
                if key not in seen_candidates:
                    seen_candidates.add(key)
                    out.append(f"    [候選彈出視窗] 相對(x={rel_x},y={rel_y}) "
                                f"class={cls!r} text={t!r}")
                    log(f"  [候選彈出視窗] 相對(x={rel_x},y={rel_y}) "
                        f"class={cls!r} text={t!r}")
            if text and text != last_text:
                results.append({"rel_x": rel_x, "rel_y": rel_y, "text": text})
                log(f"  相對(x={rel_x},y={rel_y})  tooltip={text!r}")
                out.append(f"  相對(x={rel_x},y={rel_y})  tooltip={text!r}")
                last_text = text
            elif not text:
                last_text = None
            x -= 12
    return results


def click_toolbar_icon(win, rel_xy, out, label):
    tbs = find_by_class(win, "TDSUCToolBar.UnicodeClass", "ToolBar")
    if not tbs:
        raise RuntimeError(f"[{label}] 找不到TDSUCToolBar控制項。")
    tb = tbs[0]
    r = tb.rectangle()
    abs_xy = (r.left + rel_xy[0], r.top + rel_xy[1])
    log(f"  [{label}] 工具列rect={r},點擊相對座標{rel_xy} -> 螢幕座標{abs_xy}")
    out.append(f"[{label}] 工具列rect={r} 點擊相對{rel_xy} 螢幕{abs_xy}")
    from pywinauto import mouse
    mouse.click(button="left", coords=abs_xy)


def wait_new_window(known, out, label, timeout=20, want_class=None):
    deadline = time.time() + timeout
    while time.time() < deadline:
        news = [w for w in all_top_windows() if w.handle not in known]
        if want_class:
            hits = [w for w in news if w.class_name() == want_class]
            if hits:
                return hits[0], news
        elif news:
            noise = ("IME", "MSCTFIME UI", "tooltips_class32", "OleDdeWndClass",
                     "TPUtilWindow", "TThreadWindow", "ADODB.AsyncEventMessenger")
            real = [w for w in news if w.class_name() not in noise]
            if real:
                return real[0], news
        time.sleep(0.5)
    news = [w for w in all_top_windows() if w.handle not in known]
    out.append(f"[{label}] 逾時{timeout}秒沒等到新視窗。期間新出現的視窗:")
    for w in news:
        try:
            out.append(f"    class={w.class_name()!r} text={w.window_text()!r}")
        except Exception:
            pass
    return None, news


def try_dismiss_error_dialog(known_before_open, out, wait=3.0):
    """開啟報表之後如果跳出Windows錯誤訊息框(例如Access violation那個),
    嘗試按掉它,讓閱覽報表視窗有機會完整顯示。找不到就算了,不當成錯誤。"""
    time.sleep(wait)
    candidates = [w for w in all_top_windows()
                  if w.handle not in known_before_open and w.class_name() == MSGBOX_CLASS]
    if not candidates:
        out.append("[錯誤訊息框] 沒有偵測到#32770訊息框(這次可能沒跳錯誤,或跳的方式不同)。")
        log("  沒偵測到Windows訊息框,繼續往下。")
        return False

    dlg = candidates[0]
    text = ""
    try:
        text = dlg.window_text()
    except Exception:
        pass
    log(f"  偵測到訊息框: text={text!r},嘗試按掉它...")
    out.append(f"[錯誤訊息框] 偵測到: text={text!r}")

    dismissed = False
    for label in ("確定", "OK", "關閉", "是(Y)", "是"):
        try:
            btns = [c for c in dlg.descendants()
                    if (c.window_text() or "").strip() == label]
            if btns:
                btns[0].click()
                dismissed = True
                out.append(f"[錯誤訊息框] 已點擊按鈕文字={label!r}")
                log(f"  已點擊「{label}」按鈕關閉訊息框。")
                break
        except Exception as e:
            out.append(f"[錯誤訊息框] 嘗試點擊{label!r}失敗: {e!r}")

    if not dismissed:
        try:
            dlg.set_focus()
            dlg.type_keys("{ENTER}")
            dismissed = True
            out.append("[錯誤訊息框] 找不到明確按鈕文字,改送Enter關閉。")
            log("  找不到明確按鈕文字,改送Enter。")
        except Exception as e:
            out.append(f"[錯誤訊息框] 送Enter也失敗: {e!r}")

    time.sleep(1)
    return dismissed


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

        # 2026-09-07第2輪修正:上一版這裡沒指定want_class,靠「第一個冒出來
        # 、不是雜訊視窗」這個通用邏輯去抓——第1次測試剛好抓對(class=
        # TfrmReportViewer.UnicodeClass),但第2次測試卻抓到一個完全不相關、
        # 空白的過渡視窗(class='TForm',text='',rect只有400x300那麼小,
        # 明顯不是真正的閱覽報表視窗),導致後面不管怎麼重新確認都在對著
        # 這個錯的小視窗打轉,永遠找不到工具列。研判是「開啟檔案」之後,
        # 真正的閱覽報表視窗完整顯示之前,ERP可能會先短暫跳出一個類似
        # loading/過渡用的小TForm,通用比對邏輯運氣不好就會抓到這個。
        # 改成優先明確指定want_class="TfrmReportViewer.UnicodeClass"(已經
        # 從第1次測試確認過這是正確的class name)去等,只有真的等不到才
        # 退回用通用邏輯(而且會在log裡清楚標示是走哪一條路,方便之後追查)。
        rv_win, _ = wait_new_window(known2, out, "閱覽報表", timeout=20,
                                     want_class="TfrmReportViewer.UnicodeClass")
        if rv_win is None:
            out.append("[警告] 等不到class='TfrmReportViewer.UnicodeClass'的視窗,"
                        "改用通用邏輯(抓第一個非雜訊的新視窗)當備案...")
            log("  [警告] 指定class等不到,退回通用邏輯。")
            rv_win, _ = wait_new_window(known2, out, "閱覽報表(備案)", timeout=15)
        if rv_win is None:
            out.append("[結果] 按下「開啟檔案」之後沒有新視窗,停止。")
            return
        log(f"  開啟成功: class={rv_win.class_name()!r} text={rv_win.window_text()!r}")
        out.append(f"[第2站成功] class={rv_win.class_name()!r} text={rv_win.window_text()!r}")

        # 2026-09-07第3輪修正:前兩輪不管怎麼修reacquire/等待邏輯,結果都
        # 一樣——視窗一開始抓到的class/標題都對,但幾秒後控制代碼就失效、
        # 重新搜尋也找不到,代表視窗真的整個關掉/消失了,不是我們抓錯視窗
        # 的問題。回頭比對:使用者自己用滑鼠手動操作、全程沒有任何程式介入時,
        # 報表視窗可以正常顯示、留著不會消失;但只要是這支腳本自動化去處理
        # (含這裡原本緊接著就對rv_win呼叫的.set_focus()),就會在數秒內
        # 出問題。懷疑是這份報表資料量大、內容還在陸續查詢/畫面陸續繪製時,
        # 過早對這個視窗做任何程式互動(即使只是.set_focus()這種看似無害
        # 的操作)可能干擾到ReportViewerC.dll內部還在跑的流程,間接觸發
        # Access violation讓視窗整個關掉——這就是這幾輪一直沒解決的
        # Access violation問題本體,不是偵察腳本邏輯寫錯。
        #
        # 修法:先完全不去碰rv_win(不set_focus、不做任何互動),單純被動
        # 等一段夠長的時間(5秒)給報表引擎充分時間把內容畫完,再開始做
        # 「檢查錯誤訊息框」這種只讀動作,最後才進入下面的reacquire流程。
        log("\n===== 第3站:先被動等待,不碰視窗,給報表引擎足夠時間畫完內容 =====")
        time.sleep(5.0)

        log("\n===== 第3站(續):檢查有沒有跳錯誤訊息框,嘗試按掉 =====")
        try_dismiss_error_dialog(known2, out, wait=1.0)

        # 2026-09-07新增:9/7這次實測發現報表其實有正常顯示(使用者自己截圖
        # 確認畫面上有工具列、有報表資料),但這支腳本自己的
        # capture_as_image()卻回傳None、descendants()也是空的——推測是
        # 這份報表資料量大,「開啟檔案」之後畫面是先跳出一個還沒畫好內容
        # 的過渡視窗(跟第2站抓到的rv_win是同一個handle,但這時候子控制項
        # 可能還沒建好),等資料真的查完、畫面畫出來之後才「補齊」內容——
        # 如果rv_win這個控制代碼中途被整個換掉(視窗關掉重開,不是同一個
        # handle),或只是內容還在陸續加入,直接沿用第2站當下抓到的
        # rv_win物件就可能撲空。這裡改成:診斷之前先確認視窗控制代碼還
        # 存在,並且多等一段時間、視窗有實際大小之後,直接用「目前」的
        # class+標題重新在桌面上找一次最新的視窗物件,不要沿用第2站當時
        # 那個可能已經過時的參照。
        import win32gui

        # 2026-09-07第4輪修正:第3輪拿掉.set_focus()、多等5秒之後,結果還是
        #一模一樣——連續好幾次測試都是同一個模式:原本的控制代碼真的完全
        # 失效(win32gui.IsWindow()回傳False,不是還在但只是暫時抓不到而
        # 已),而且重新用class在**桌面所有頂層視窗**裡搜尋也一次都找不到。
        # 這排除了「我們自己的互動害它掛掉」的猜測(這次全程沒碰它),改指向
        # 另一種可能:閱覽報表視窗標題列的格式是「佇列工作控制台[...] -
        # [閱覽報表(...)]」,這是Windows MDI(多重文件介面)子視窗被放大時的
        # 經典標題格式——Delphi的fsMDIChild表單,常見做法是先短暫建立成一個
        # 一般的頂層視窗,建立完成後才呼叫SetParent()把它重新掛到佇列工作
        # 管理員(qm_win)底下的MDI Client容器裡,變成真正的MDI子視窗。一旦
        # 完成這個轉換,它就不再是「頂層視窗」,`Desktop(backend="win32")
        # .windows()`(以EnumWindows為基礎,只列頂層視窗)自然再也找不到它,
        # 但它其實還活著、還在畫面上——只是要往qm_win底下的子視窗清單裡找,
        # 不是繼續在桌面頂層視窗清單裡找。
        #
        # 這裡在原本的頂層搜尋之外,新增一條路:改在`qm_win.descendants()`
        # (完整往下遍歷所有子視窗,不限頂層)裡面找同樣class name的視窗。
        def reacquire_rv_win():
            still_valid = False
            try:
                still_valid = win32gui.IsWindow(rv_win.handle)
            except Exception:
                pass
            if not still_valid:
                out.append("[第4站] rv_win原本的控制代碼已經不存在了,改用class重新搜尋...")
                log("  [警告] rv_win控制代碼已經不存在,重新搜尋同class的視窗。")
            hits = [w for w in all_top_windows()
                    if w.class_name() == "TfrmReportViewer.UnicodeClass"]
            if hits:
                out.append("[第4站] 在桌面頂層視窗清單裡找到了。")
                return hits[0]
            try:
                child_hits = [c for c in qm_win.descendants()
                              if c.class_name() == "TfrmReportViewer.UnicodeClass"]
            except Exception as e:
                child_hits = []
                out.append(f"[第4站] 搜尋qm_win子視窗時發生例外: {e!r}")
            if child_hits:
                out.append("[第4站] 頂層視窗清單裡沒有,但在qm_win底下的子視窗裡找到了"
                            "——很可能已經變成MDI子視窗。")
                log("  [重要] 在qm_win的子視窗裡找到閱覽報表視窗(MDI子視窗)。")
                return child_hits[0]
            return rv_win if still_valid else None

        rv_win_fresh = None
        for attempt in range(6):
            time.sleep(1.0)
            rv_win_fresh = reacquire_rv_win()
            if rv_win_fresh is None:
                continue
            try:
                r = rv_win_fresh.rectangle()
                if (r.right - r.left) > 10 and (r.bottom - r.top) > 10:
                    break
            except Exception:
                pass
        if rv_win_fresh is None:
            out.append("[第4站] 多次重新搜尋後還是找不到閱覽報表視窗,放棄診斷。")
            log("  [錯誤] 找不到閱覽報表視窗,結束。")
            out.append("\n[到此為止] 不會做轉Excel、不會關閉閱覽報表視窗。")
            return
        rv_win = rv_win_fresh
        try:
            r = rv_win.rectangle()
            out.append(f"[第4站] 重新確認視窗: handle={rv_win.handle} rect={r}")
            log(f"  重新確認視窗: handle={rv_win.handle} rect={r}")
        except Exception as e:
            out.append(f"[第4站] 讀取視窗rect失敗: {e!r}")

        log("\n===== 第4站:偵察閱覽報表視窗目前的狀態 =====")
        img = save_shot(rv_win, "rv_probe_full.png", out)
        dump_control_tree(rv_win, out, "閱覽報表視窗(ReportView,錯誤訊息框處理後)")
        n_tb = save_numbered_toolbar(rv_win, img, out, "rv_probe")

        if n_tb == 0:
            out.append(
                "\n[重要] 這次還是沒有在閱覽報表視窗底下找到TDSUCToolBar控制項——"
                "代表報表內容/工具列可能真的沒有正常初始化(不只是跳一個訊息框而已),"
                "需要請使用者手動走一次同樣的步驟(開佇列工作管理員->選最上面那筆工作->"
                "開啟檔案),看看正常情況下這個視窗長怎樣、工具列在哪裡,而不是繼續"
                "用自動化去猜。"
            )
            log("\n[重要] 沒找到工具列,可能報表內容本身沒有正常初始化,請看記錄檔。")
        else:
            out.append(f"\n[結果] 這次找到{n_tb}個工具列,座標可以從rv_probe_toolbar*_ruler.png量出來。")
            log(f"\n[結果] 找到{n_tb}個工具列,截圖已存,可以量座標了。")

            log("\n===== 第5站(前):先讓閱覽報表視窗成為作用中的MDI子視窗 =====")
            # 2026-09-07第6輪:上一版在這裡完全沒碰rv_win就直接掃tooltip,
            # 結果掃到0個。第4輪已經證實set_focus()本身不會造成Access
            # violation(那是MDI子視窗偵測邏輯的問題,跟互動與否無關),這裡
            # 補回set_focus(),測試是不是「視窗不是作用中視窗」導致tooltip
            # 邏輯不顯示提示框。
            try:
                rv_win.set_focus()
                out.append("[第5站前] 已對rv_win呼叫set_focus()(第4輪已確認"
                            "這不會導致Access violation)。")
                log("  已對rv_win呼叫set_focus()。")
            except Exception as e:
                out.append(f"[第5站前] set_focus()失敗: {e!r}")
                log(f"  [警告] set_focus()失敗: {e!r}")
            time.sleep(0.5)

            log("\n===== 第5站:掃過工具列,讀每個位置的tooltip文字 =====")
            tbs = find_by_class(rv_win, "TDSUCToolBar.UnicodeClass", "ToolBar")
            if tbs:
                hints = sweep_toolbar_tooltips(tbs[0], out)
                out.append("\n[第5站] 工具列每個圖示的tooltip(依相對x座標由左到右,"
                            "座標是相對工具列自己左上角):")
                if hints:
                    for h in hints:
                        out.append(f"    相對x={h['rel_x']:>4}  tooltip={h['text']!r}")
                else:
                    out.append("    沒有掃到任何tooltip文字——可能這個工具列其實也沒有"
                                "tooltip(之前那個「刪除格式」也許是別的東西),或者滑鼠"
                                "移動/停留的時機需要調整。")
            else:
                out.append("[第5站] 找不到工具列,略過tooltip掃描。")

            log("\n===== 第6站:掃標題列右上角(關閉/最大化/最小化按鈕)的tooltip文字 =====")
            # 2026-09-07第6輪新增,回應使用者指令:「你也要用tooltip 掃藍色
            # 圈圈的位置」——標註截圖裡藍色圈圈圈的是標題列右上角的關閉(X)
            # 按鈕。這個按鈕不在TDSUCToolBar底下,座標用視窗矩形右上角
            # 往回推、逐點掃描。
            title_hints = sweep_titlebar_buttons(rv_win, out)
            out.append("\n[第6站] 標題列右上角區域的tooltip(座標是相對rv_win"
                        "視窗自己左上角):")
            if title_hints:
                for h in title_hints:
                    out.append(f"    相對(x={h['rel_x']},y={h['rel_y']})  "
                                f"tooltip={h['text']!r}")
            else:
                out.append("    沒有掃到任何tooltip文字——標題列按鈕的座標可能"
                            "跟猜測的範圍(右邊160px、y偏移10~34px)對不上,"
                            "需要再調整掃描範圍,或這幾顆按鈕本身也沒有tooltip。")

        out.append("\n[到此為止] 不會做轉Excel、不會關閉閱覽報表視窗,視窗留著給你確認。")

    except Exception as e:
        out.append("")
        out.append("!" * 70)
        out.append(f"執行中發生錯誤: {e!r}")
        out.append("!" * 70)
        log(f"[錯誤] {e!r}")
        raise
    finally:
        with open("reportview_probe_output.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(out))
        log("\n記錄已寫入: reportview_probe_output.txt")
        log("請把 reportview_probe_output.txt、rv_probe_full.png、"
            "rv_probe_toolbar*_ruler.png(如果有)都傳給我。")


if __name__ == "__main__":
    sys.exit(main())
