#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Part 2 — 第0步:控制項偵察工具(erp_inspect.py)
================================================

目的
----
在真的寫「送查詢、處理F2庫別選擇、按確認、擷取工作編號」這些自動化邏輯之前,
我們需要先知道這個鼎新ERP client每一個畫面(登入視窗/主畫面/程式代號欄位/
查詢對話框/F2開窗查詢)在Windows API層級長什麼樣子——class name、
control id、文字內容、視窗階層結構。

這支腳本**只做安全、唯讀的動作**:
1. 用Windows認證管理員裡的 DigiwinERP 帳密自動登入
2. 把登入後主畫面的控制項結構印出來,存成文字檔
3. 把左上角「檢視類別」下拉選單切到「系統+作業別_IDL」(使用者確認:切換前
   工具列上完全沒有「程式代號」欄位,切換後才會出現),再把切換後的主畫面
   控制項結構也印出來
4. 在主視窗「程式代號」欄位打入你指定的代號(預設 IDLR17,這是我們已經
   確認過「不需要任何選擇、也不會跳出F2視窗」的最簡單報表,用來測試安全)
   按Enter,等查詢畫面跳出來,把這個查詢畫面的控制項結構也印出來
5. **不會**按「確認」送出查詢、不會處理F2視窗、不會產生任何ERP工作

跑完之後,把這支腳本印出來的文字檔(預設存在同資料夾的 inspect_output.txt)
傳給我,我就可以照著真實的control class name/id把後面幾支腳本(選庫別、
按確認、擷取工作編號、轉EXCEL)寫準,不用用猜的、少走很多冤枉路。

執行需求
--------
    pip install pywinauto pywin32
    python part2_inspect.py                  # 用預設代號 IDLR17 測試
    python part2_inspect.py --code IDLR68    # 也可以換其他不需要選擇的代號測試

安全性
------
帳密一樣是從Windows「認證管理員」讀,不寫死在腳本裡(跟Part 1一致的做法)。

已知的警告訊息(可以先不理)
--------------------------
執行時可能會看到:
    UserWarning: 32-bit application should be automated using 32-bit Python
    (you use 64-bit Python)
這是因為MainMenu.exe是32位元的舊版Delphi程式,你電腦裝的Python是64位元。
大部分操作(找控制項、讀文字、填欄位)跨32/64位元還是能動,這個警告不是這次
「找不到帳密欄位」那個錯誤的原因(那個是另一個bug,已修好)。如果之後遇到
「填欄位填不進去/勾選核取方塊沒反應」這種更詭異的問題,才需要考慮改裝
32位元版Python(python.org下載頁面選32-bit/Win32版本)來澈底排除這個因素,
現在還不用急著弄。
"""

import argparse
import ctypes
import os
import sys
import time
from ctypes import wintypes

import pipeline_flags

# ============================================================
# 設定區
# ============================================================
# 下面兩個值請依你自己的環境調整。也可以不改原始碼,直接設定同名的環境變數
# (CRED_TARGET / ERP_EXE)覆蓋這裡的預設值,方便不同台電腦、不同安裝路徑
# 共用同一份程式碼。

CRED_TARGET = os.environ.get("CRED_TARGET", "DigiwinERP")  # Windows認證管理員裡,你自己存的那筆帳密的目標名稱
ERP_EXE = os.environ.get("ERP_EXE", r"C:\Conductor\C_dsbin\MainMenu.exe")  # 你的鼎新ERP主程式路徑
OUTPUT_FILE = "inspect_output.txt"
VIEW_CATEGORY = "系統+作業別_IDL"     # 登入後「檢視類別」下拉要切到這個,「程式代號」欄位才會出現

LOGIN_WAIT_TIMEOUT = 30      # 秒,等登入視窗跳出來的逾時時間
MAIN_WINDOW_WAIT_TIMEOUT = 30  # 秒,等登入成功、主畫面出現的逾時時間
QUERY_DIALOG_WAIT_TIMEOUT = 15  # 秒,等打完代號+Enter後查詢畫面跳出來的逾時時間


# ============================================================
# 從Windows認證管理員讀帳密(Python版,不需要PowerShell)
# ============================================================

class FILETIME(ctypes.Structure):
    _fields_ = [("dwLowDateTime", wintypes.DWORD),
                ("dwHighDateTime", wintypes.DWORD)]


class CREDENTIAL(ctypes.Structure):
    _fields_ = [
        ("Flags", wintypes.DWORD),
        ("Type", wintypes.DWORD),
        ("TargetName", wintypes.LPWSTR),
        ("Comment", wintypes.LPWSTR),
        ("LastWritten", FILETIME),
        ("CredentialBlobSize", wintypes.DWORD),
        ("CredentialBlob", ctypes.POINTER(ctypes.c_byte)),
        ("Persist", wintypes.DWORD),
        ("AttributeCount", wintypes.DWORD),
        ("Attributes", ctypes.c_void_p),
        ("TargetAlias", wintypes.LPWSTR),
        ("UserName", wintypes.LPWSTR),
    ]


def get_credential(target_name):
    """從Windows認證管理員讀出「Windows認證」(不是「一般/Web認證」)。
    對應之前PAD用的PowerShell CredRead手法,這裡直接用Python ctypes呼叫同一個
    Win32 API,不需要另外開PowerShell行程。"""
    advapi32 = ctypes.WinDLL("advapi32", use_last_error=True)
    cred_read = advapi32.CredReadW
    cred_read.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                           ctypes.POINTER(ctypes.POINTER(CREDENTIAL))]
    cred_read.restype = wintypes.BOOL
    cred_free = advapi32.CredFree
    cred_free.argtypes = [ctypes.c_void_p]

    cred_ptr = ctypes.POINTER(CREDENTIAL)()
    # CRED_TYPE_GENERIC=1 是「一般認證」,Windows認證是CRED_TYPE_DOMAIN_PASSWORD=2。
    # 因為之前PAD筆記說是手動新增「Windows認證」,這裡兩種都試,哪個成功用哪個。
    for cred_type in (2, 1):
        ok = cred_read(target_name, cred_type, 0, ctypes.byref(cred_ptr))
        if ok:
            break
    else:
        err = ctypes.get_last_error()
        raise OSError(
            f"CredRead失敗(錯誤碼 {err})。請確認Windows認證管理員裡有一筆"
            f"目標名稱為 '{target_name}' 的認證(認證管理員 -> Windows認證 -> 新增)。"
        )

    cred = cred_ptr.contents
    blob_size = cred.CredentialBlobSize
    if blob_size:
        blob = ctypes.string_at(cred.CredentialBlob, blob_size)
        password = blob.decode("utf-16-le", errors="ignore")
    else:
        password = ""
    username = cred.UserName or ""
    cred_free(cred_ptr)
    return username, password


# ============================================================
# 登入 + 控制項傾印
# ============================================================

def log(msg):
    print(msg, flush=True)


def dump_control_tree(window, out_lines, label):
    """把一個視窗的完整控制項結構(class name/control id/文字/座標/階層)
    印成文字,附加到out_lines清單裡。用pywinauto內建的print_control_identifiers
    機制,但導向字串而不是直接印到終端機,這樣才能寫進檔案。"""
    import io
    import contextlib

    out_lines.append("")
    out_lines.append("=" * 70)
    out_lines.append(f"[{label}]")
    out_lines.append("=" * 70)

    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf):
            window.print_control_identifiers(depth=None)
    except Exception as e:
        buf.write(f"(print_control_identifiers 失敗: {e})\n")
    out_lines.append(buf.getvalue())

    # 額外印出所有子控制項的「class name / 文字 / 矩形座標」精簡列表,
    # 這個通常比print_control_identifiers更容易一眼看出「哪個是帳號欄、
    # 哪個是密碼欄、程式代號欄在哪」。
    out_lines.append("-" * 70)
    out_lines.append(f"[{label} - 精簡控制項清單(class / rectangle / text)]")
    out_lines.append("-" * 70)
    try:
        descendants = window.descendants()
    except Exception as e:
        descendants = []
        out_lines.append(f"(descendants() 失敗: {e})")
    for i, ctrl in enumerate(descendants):
        try:
            raw_cls = ctrl.class_name()
        except Exception:
            raw_cls = "?"
        try:
            friendly_cls = ctrl.friendly_class_name()
        except Exception:
            friendly_cls = "?"
        # 印「原始class_name」為主,因為之前發現friendly_class_name()會把不同的
        # 原始class簡化成同一個籠統名稱(例如把TDSUCEdit.UnicodeClass跟純Edit都
        # 顯示成'Edit'),容易誤導判斷。兩個都印,原始的在前面。
        cls = f"{raw_cls}(friendly={friendly_cls})" if raw_cls != friendly_cls else raw_cls
        try:
            rect = ctrl.rectangle()
        except Exception:
            rect = "?"
        try:
            text = ctrl.window_text()
        except Exception:
            text = "?"
        try:
            handle = ctrl.handle
        except Exception:
            handle = "?"
        out_lines.append(f"  [{i:3d}] class={cls!r:35s} handle={handle} rect={rect} text={text!r}")


def safe_all_windows(retries=5, delay=0.2):
    """比照`Desktop(backend="win32").windows()`,但改成對race condition有
    抵抗力。

    2026-09-07新增,起因:實測發現`Desktop(backend="win32").windows()`會
    先枚舉「這一刻」所有頂層視窗的handle,再依序把每個handle包裝成
    HwndWrapper物件——如果枚舉完到包裝完這段極短的時間內,剛好有某個
    視窗被關掉(例如某個一閃即逝的過渡/loading小視窗,不見得是我們關心
    的對象),pywinauto目前的實作會直接對那個視窗拋出
    `InvalidWindowHandle`,而且沒有任何保護,整支呼叫端腳本(包括login()
    這種從一開始就會用到的共用函式)都會被這個跟我們真正要找的視窗完全
    無關的例外中斷。實測就是在`login()`裡的`enum_process_windows()`踩到
    這個問題而crash掉。

    這裡改成:呼叫`Desktop(backend="win32").windows()`失敗就直接重試整次
    枚舉(不是重試包裝單一視窗,因為pywinauto沒有提供逐一包裝的接口)——
    下一次重新枚舉時,那個剛好消失的視窗多半已經不在最新的頂層視窗清單
    裡了,重試個幾次通常就能繞過這種瞬間的race condition。真的每次都失敗
    (不像是純運氣問題)才把最後一次的例外往外丟。
    """
    from pywinauto import Desktop

    last_err = None
    for _ in range(retries):
        try:
            return Desktop(backend="win32").windows()
        except Exception as e:
            last_err = e
            time.sleep(delay)
    raise last_err


def enum_process_windows(pid):
    """列出「目前」屬於指定pid的所有頂層視窗。

    不用 app.top_window()/app.windows() 的原因:MainMenu.exe啟動後會先跳出
    一個沒有任何子控制項的LOGO/啟動畫面(class TfrmLogo.UnicodeClass),過幾秒
    才會出現真正的登入視窗。pywinauto的 top_window() 是挑「目前z-order最上面」
    的視窗,實測發現它在LOGO視窗還沒關閉、登入視窗已經出現的這段期間,兩次呼叫
    可能拿到不同的視窗(這支腳本第一版就是因為這樣,明明wait_until_passes通過了,
    緊接著再呼叫一次top_window()卻拿回LOGO視窗,導致找不到帳密欄位)。
    改成用 win32process.GetWindowThreadProcessId 直接比對pid、掃過「這個process
    當下所有的頂層視窗」,不管Z-order,就不會有這個問題。"""
    import win32process

    result = []
    for w in safe_all_windows():
        try:
            _, w_pid = win32process.GetWindowThreadProcessId(w.handle)
        except Exception:
            continue
        if w_pid == pid:
            result.append(w)
    return result


def find_by_class(root, exact_class, substr):
    """先用精確的原始class name(exact_class)找,找不到才退回用substring比對
    (包在整個descendants()清單裡找class_name()含substr的)。

    背景:這個ERP的控制項幾乎都是自訂的Delphi元件(例如TDSUCEdit.UnicodeClass、
    TDSUCComboBox.UnicodeClass),但pywinauto的friendly_class_name()會把這些
    簡化成通用名稱(Edit、ComboBox)。如果直接用descendants(class_name="Edit")
    這種通用名稱去找,實際上是在找「原始class name剛好等於Edit」的控制項,
    根本找不到任何TDSUCEdit.UnicodeClass——這正是先前兩次踩到的同一種坑
    (登入欄位、工具列程式代號欄位都是這樣),這裡統一用「先精確找、找不到再
    退回substring」來避免第三次犯同樣的錯。"""
    ctrls = root.descendants(class_name=exact_class)
    if ctrls:
        return ctrls
    return [c for c in root.descendants() if substr in (c.class_name() or "")]


# ============================================================
# checkbox勾選狀態讀取(截圖比對顏色,取代不可靠的BM_GETCHECK)
# ============================================================
#
# 背景:這個ERP的checkbox(class TDSUCBitCheckBox.UnicodeClass)是owner-drawn
# 的自訂控制項。2026-09-03在INVR18「進階選項」頁籤測試時,第一版曾經用
# Win32原生BM_GETCHECK訊息去問checkbox目前是不是勾選狀態,結果讀回來的值
# 是假的——使用者截圖明明顯示「列印各庫明細」是打勾狀態,BM_GETCHECK卻回報
# 0(未勾選),導致腳本誤判「已經是要的狀態不用點」,實際上完全沒把勾去掉,
# 是一個靠使用者肉眼比對截圖才抓到的真正bug。
#
# 改成「截圖比對顏色」的方式,已用使用者的實際截圖完整驗證成功(INVR18的
# 「列印各庫明細」:變更前偵測到66個藍色像素判定為勾選,點擊後偵測到0個
# 判定為未勾選,收尾前再確認一次仍是0,三次讀取都跟畫面實際狀態吻合):
# 對checkbox控制項本身capture_as_image(),只看最左邊
# CHECKBOX_GLYPH_WIDTH像素寬的範圍(勾勾圖示畫在這裡),數有多少個像素是
# 「藍色勾勾筆畫」(B分量比R/G分量高出BLUE_DOMINANCE_THRESHOLD以上;沒勾選
# 的方塊只有灰階像素,R=G=B)。藍色像素數量超過
# MIN_BLUE_PIXELS_FOR_CHECKED就判定為「勾選」。
#
# 這個函式取代了原本各腳本各自土法煉鋼、還一度用錯BM_GETCHECK的做法,
# 之後任何要讀取TDSUCBitCheckBox勾選狀態的腳本都應該直接呼叫這個函式,
# 不要再自己重新發明或使用BM_GETCHECK。

CHECKBOX_GLYPH_WIDTH = 28
BLUE_DOMINANCE_THRESHOLD = 40
MIN_BLUE_PIXELS_FOR_CHECKED = 5


def is_checkbox_checked(checkbox, save_path=None):
    """判斷一個TDSUCBitCheckBox.UnicodeClass控制項目前是不是勾選狀態。

    回傳True=勾選、False=未勾選、None=讀取失敗(例如截圖失敗)。
    如果有給save_path,會把截圖存檔,方便事後肉眼比對這個判斷準不準。
    """
    try:
        img = checkbox.capture_as_image()
        if img is None:
            return None
        if save_path:
            img.save(save_path)

        w, h = img.size
        glyph_w = min(CHECKBOX_GLYPH_WIDTH, w)
        blue_count = 0
        for y in range(h):
            for x in range(glyph_w):
                r, g, b = img.getpixel((x, y))[:3]
                if b - max(r, g) > BLUE_DOMINANCE_THRESHOLD:
                    blue_count += 1
        return blue_count >= MIN_BLUE_PIXELS_FOR_CHECKED
    except Exception:
        return None


def set_checkbox_checked(checkbox, want_checked, out_lines, label):
    """確保checkbox達到want_checked這個狀態:已經是的話不點,不是才點一下切換。

    會把每一步的判斷結果跟截圖路徑記進out_lines,方便你事後核對。
    """
    import time

    before_path = f"checkbox_{label}_before.png"
    before = is_checkbox_checked(checkbox, save_path=before_path)
    log(f"  [{label}] 變更前狀態: {before}(截圖:{before_path})")
    out_lines.append(f"  [{label}] 變更前狀態: {before}(截圖:{before_path})")

    if before is None:
        log(f"  [{label}] [警告] 讀取狀態失敗,保險起見不自動點擊。")
        out_lines.append(f"  [{label}] [警告] 讀取狀態失敗,不自動點擊。")
        return None

    if before == want_checked:
        log(f"  [{label}] 已經是要的狀態({want_checked}),不用點。")
        out_lines.append(f"  [{label}] 已經是要的狀態,不用點。")
        return before

    log(f"  [{label}] 目前是{before},要變成{want_checked},點擊切換...")
    checkbox.click()
    time.sleep(0.5)

    after_path = f"checkbox_{label}_after.png"
    after = is_checkbox_checked(checkbox, save_path=after_path)
    log(f"  [{label}] 點擊後狀態: {after}(截圖:{after_path})")
    out_lines.append(f"  [{label}] 點擊後狀態: {after}(截圖:{after_path})")
    if after != want_checked:
        log(f"  [{label}] [警告] 點擊後狀態({after})跟預期({want_checked})不符,"
            "請務必人工核對截圖。")
        out_lines.append(f"  [{label}] [警告] 點擊後狀態跟預期不符,請人工核對。")
    return after


def find_login_edits(login_dlg):
    """在已經確定是登入視窗的物件裡,找兩個Edit欄位(class TDSUCEdit.UnicodeClass),
    依畫面上的垂直位置排序(由上到下),回傳 (帳號欄, 密碼欄)。"""
    edits = find_by_class(login_dlg, "TDSUCEdit.UnicodeClass", "Edit")
    if len(edits) < 2:
        raise RuntimeError(
            f"預期在登入視窗裡找到2個TDSUCEdit.UnicodeClass欄位,實際找到{len(edits)}個。"
            f"請把 {OUTPUT_FILE} 的內容回報,才能確認正確的欄位怎麼找。"
        )
    edits_sorted = sorted(edits, key=lambda c: c.rectangle().top)
    return edits_sorted[0], edits_sorted[1]


def login(out_lines):
    from pywinauto.application import Application

    log(f"讀取帳密(Windows認證管理員,目標名稱={CRED_TARGET})...")
    username, password = get_credential(CRED_TARGET)
    log(f"  帳號: {username}(密碼已讀取,不顯示)")

    log(f"啟動ERP: {ERP_EXE}")
    app = Application(backend="win32").start(ERP_EXE)
    pid = app.process

    log(f"等待登入視窗出現(逾時 {LOGIN_WAIT_TIMEOUT} 秒)...")
    login_dlg = None
    seen = {}
    deadline = time.time() + LOGIN_WAIT_TIMEOUT
    while time.time() < deadline and login_dlg is None:
        for w in enum_process_windows(pid):
            try:
                cls, txt = w.class_name(), w.window_text()
            except Exception:
                cls, txt = "?", "?"
            seen[(cls, txt)] = True
            try:
                if len(w.descendants(class_name="TDSUCEdit.UnicodeClass")) >= 2:
                    login_dlg = w
                    break
            except Exception:
                continue
        if login_dlg is None:
            time.sleep(0.5)

    if login_dlg is None:
        out_lines.append(f"{LOGIN_WAIT_TIMEOUT}秒內,這個process出現過的所有頂層視窗(class | title):")
        for cls, txt in seen:
            out_lines.append(f"  class={cls!r} title={txt!r}")
        raise RuntimeError(
            f"{LOGIN_WAIT_TIMEOUT}秒內沒有找到含2個TDSUCEdit.UnicodeClass欄位的視窗。"
            f"已把這段時間看到的所有視窗class/title記錄到{OUTPUT_FILE},請回報。"
        )

    dump_control_tree(login_dlg, out_lines, "登入視窗(填帳密之前)")

    account_edit, password_edit = find_login_edits(login_dlg)
    log("找到帳號欄與密碼欄,填入帳密...")
    account_edit.set_edit_text(username)
    password_edit.set_edit_text(password)

    log("送出登入(按Enter)...")
    password_edit.type_keys("{ENTER}", pause=0.05)

    log(f"等待主畫面出現(逾時 {MAIN_WINDOW_WAIT_TIMEOUT} 秒)...")
    main_win = None
    deadline = time.time() + MAIN_WINDOW_WAIT_TIMEOUT
    while time.time() < deadline and main_win is None:
        for w in enum_process_windows(pid):
            try:
                if w.handle == login_dlg.handle:
                    continue
                txt = w.window_text() or ""
            except Exception:
                continue
            if "conductor" in txt.lower():
                main_win = w
                break
        if main_win is None:
            time.sleep(1)

    if main_win is None:
        raise RuntimeError(
            f"登入後{MAIN_WINDOW_WAIT_TIMEOUT}秒內沒有找到標題含'Conductor'的主視窗。"
            "可能是帳密錯誤、視窗標題跟預期不同,或登入時彈出了其他訊息視窗。"
            "請檢查ERP畫面實際狀況,並把目前畫面截圖回報。"
        )

    log(f"找到主畫面: {main_win.window_text()!r}")
    dump_control_tree(main_win, out_lines, "主畫面(登入成功後,切換檢視類別之前)")

    select_view_category(main_win, VIEW_CATEGORY, out_lines)
    dump_control_tree(main_win, out_lines, f"主畫面(切換檢視類別為{VIEW_CATEGORY}之後)")

    return app, main_win


def select_view_category(main_win, category, out_lines):
    """把主畫面左上角「檢視類別」下拉選單切到指定的類別。

    使用者回報:登入後預設是「系統+作業別」,「程式代號」欄位要切到
    「系統+作業別_IDL」之後才會出現在工具列上(截圖比對:切換前的工具列上
    完全沒有「程式代號」這幾個字跟輸入框,切換後才冒出來)。所以這一步是
    「打程式代號跳轉」之前的必要前置動作,不能省略。"""
    log(f"切換「檢視類別」下拉選單為 {category!r}...")
    combos = find_by_class(main_win, "TDSUCComboBox.UnicodeClass", "ComboBox")
    if not combos:
        raise RuntimeError("主畫面找不到任何ComboBox控制項(檢視類別/公司別下拉選單應該都在裡面)。")

    # 畫面上由左到右是:檢視類別 -> 公司別,所以取rect.left最小的那個。
    combos_sorted = sorted(combos, key=lambda c: c.rectangle().left)
    view_combo = combos_sorted[0]
    log(f"  找到候選「檢視類別」下拉:目前值={view_combo.window_text()!r} rect={view_combo.rectangle()}")

    try:
        view_combo.select(category)
    except Exception as e:
        raise RuntimeError(
            f"嘗試選擇「{category}」失敗:{e!r}。"
            f"請檢查{OUTPUT_FILE}裡主畫面切換前的ComboBox清單,確認是不是找對了下拉選單、"
            f"清單裡有沒有這個確切的選項文字(可能因為語言設定差1個字,例如全形/半形括號)。"
        )

    time.sleep(2)  # 給畫面一點時間重新載入樹狀選單/工具列版面
    log(f"  切換後的值: {view_combo.window_text()!r}")


def jump_to_code(main_win, code, out_lines):
    """在主畫面找「程式代號」欄位,打入代號+Enter,等查詢畫面跳出來。

    第一輪偵察(見inspect_output.txt)已經確認:主畫面上方有一條
    class='TDSUCToolBar.UnicodeClass'(text='ToolBar1')的工具列,裡面由左到右
    依序排著「系統+作業別」下拉、「公司別」下拉、幾個圖示按鈕、最後是一個空白
    的Edit欄位——這個Edit就是「程式代號」欄位(畫面上看到的標籤跟輸入框都在
    這條工具列裡)。原本用「畫面右上角+Y在視窗上方15%內」這種座標啟發式去猜,
    因為視窗實際寬度沒抓對,誤判成不在候選範圍內——改成直接鎖定
    TDSUCToolBar.UnicodeClass底下的Edit控制項,不用再猜座標比例。"""
    log(f"尋找「程式代號」欄位...")

    toolbars = main_win.descendants(class_name="TDSUCToolBar.UnicodeClass")
    if toolbars:
        search_roots = toolbars
    else:
        log("  [警告] 沒找到TDSUCToolBar.UnicodeClass,退回在整個主畫面裡找Edit。")
        search_roots = [main_win]

    candidates = []
    for root in search_roots:
        for ctrl in find_by_class(root, "TDSUCEdit.UnicodeClass", "Edit"):
            candidates.append(ctrl)

    if not candidates:
        raise RuntimeError(
            "工具列裡沒有找到任何Edit欄位。"
            f"請查看 {OUTPUT_FILE} 裡主畫面的完整控制項清單,人工確認正確的欄位,"
            "並回報給我(class name / 座標 / 是否有control id)。"
        )

    # 工具列裡如果有多個Edit,「程式代號」欄位是最右邊那一個(畫面上它在
    # 公司別/檢視類別兩個下拉選單的右側)。
    candidates.sort(key=lambda c: c.rectangle().left, reverse=True)
    code_edit = candidates[0]
    log(f"  候選欄位: class={code_edit.class_name()!r} rect={code_edit.rectangle()}"
        f"(工具列裡共找到{len(candidates)}個Edit,取最右邊那個)")

    # 2026-09-08新增:先記住「輸入代號之前」桌面上已經存在的視窗控制代碼。
    # 起因:INVR18/IDLR69這種同一個代號要連續jump_to_code好幾次的報表,如果
    # 上一組的查詢畫面因故沒關乾淨(參見submit_query的說明),桌面上會同時
    # 有好幾個標題含這個代號的視窗——單純「標題含代號」去配一定會抓錯(可能
    # 抓到舊的、已經送出過的那個),9/7實測INVR18第2組(庫存全部)就是抓到
    # 舊視窗,導致找不到「列印各庫明細」checkbox整支失敗、後面4組全部沒跑到。
    # 修法:优先只認「這次新冒出來的」視窗,舊視窗一律不列入候選。
    known_handles = {w.handle for w in safe_all_windows()}

    code_edit.set_focus()
    code_edit.set_edit_text(code)
    code_edit.type_keys("{ENTER}", pause=0.05)

    log(f"已輸入代號 {code} 並按Enter,等待查詢畫面出現(逾時 {QUERY_DIALOG_WAIT_TIMEOUT} 秒)...")
    time.sleep(QUERY_DIALOG_WAIT_TIMEOUT)  # 先用固定等待,之後有更準的偵測方式再改成輪詢

    # 查詢畫面標題通常會包含報表名稱跟代號,例如「訂單預計出貨明細表(IDL)(IDLR17)...」
    all_wins = safe_all_windows()
    new_dialogs = [w for w in all_wins
                   if w.handle not in known_handles and code in (w.window_text() or "")]

    if new_dialogs:
        query_win = new_dialogs[0]
    else:
        # 退回舊邏輯(標題含代號就算數),但這代表沒偵測到「新視窗」,可能是
        # 這次真的沒開出查詢畫面,也可能是新視窗判斷本身有問題——兩種都要
        # 明顯警告,不要默默抓一個可能是舊的視窗來用。
        dialogs = [w for w in all_wins if code in (w.window_text() or "")]
        if not dialogs:
            log("  [警告] 沒有找到標題包含代號的新視窗,改成傾印整個主視窗底下的所有視窗給你人工比對。")
            dump_control_tree(main_win, out_lines, f"主畫面(輸入{code}之後,找不到明確的查詢視窗)")
            return None
        log(f"  [警告] 沒有偵測到「新冒出來」的{code}查詢視窗,退回用標題比對"
            f"(候選共{len(dialogs)}個,可能有沒關乾淨的舊視窗殘留,請核對截圖確認選對)。")
        out_lines.append(f"  [警告] {code}查詢視窗退回標題比對,候選數={len(dialogs)}(可能有殘留舊視窗)")
        query_win = dialogs[0]

    log(f"找到查詢畫面: {query_win.window_text()!r}")
    dump_control_tree(query_win, out_lines, f"查詢畫面({code})")
    return query_win


# ============================================================
# 查詢畫面共用小工具(2026-09-03新增)
# ------------------------------------------------------------
# 之前每一支part2_*_test.py都各自複製貼上同一套find_select_panel/
# find_button/find_checkbox_by_text/find_tab_sheet/save_screenshot,
# 現在收斂成這裡的共用函式。正式的report_*.py一律從這裡import,
# 不要再各自複製一份。
# ============================================================

def find_select_panel(root, label):
    for p in find_by_class(root, "TDSUCSelectPanel.UnicodeClass", "SelectPanel"):
        try:
            text = (p.window_text() or "").strip()
        except Exception:
            continue
        if text == label:
            return p
    return None


def find_tab_sheet(root, label):
    for s in find_by_class(root, "TDSUCTabSheet.UnicodeClass", "TabSheet"):
        try:
            text = (s.window_text() or "").strip()
        except Exception:
            continue
        if text == label:
            return s
    return None


def find_button(root, text):
    for b in find_by_class(root, "TDSUCBitBtn.UnicodeClass", "Button"):
        try:
            t = (b.window_text() or "").strip()
        except Exception:
            continue
        if t == text:
            return b
    return None


def find_checkbox_by_text(root, text):
    for cb in find_by_class(root, "TDSUCBitCheckBox.UnicodeClass", "CheckBox"):
        try:
            t = (cb.window_text() or "").strip()
        except Exception:
            continue
        if t == text:
            return cb
    return None


def save_screenshot(window, path, out_lines):
    if not pipeline_flags.SCREENSHOTS_ENABLED:
        out_lines.append(f"[截圖已略過(SCREENSHOTS_ENABLED=False): {path}]")
        return None
    try:
        img = window.capture_as_image()
        if img is None:
            raise RuntimeError("capture_as_image()回傳None")
        img.save(path)
        log(f"  已存截圖: {path}")
        out_lines.append(f"[截圖已存: {path}]")
        return img
    except Exception as e:
        log(f"  [警告] 截圖失敗({path}): {e!r}")
        out_lines.append(f"[截圖失敗: {path} -> {e!r}]")
        return None


def roc_date_str(d):
    """西元date物件 -> 民國「YYY/MM/DD」字串,例如2026-09-03 -> '115/09/03'。"""
    roc_year = d.year - 1911
    return f"{roc_year:03d}/{d.month:02d}/{d.day:02d}"


# ============================================================
# owner-drawn分頁條(TDSUCExtPageControl):Ctrl+Tab切換
# ------------------------------------------------------------
# 2026-09-03在IDLR68上驗證成功、取代猜像素座標的正確做法(詳見
# erp-pywinauto-techniques.md第3節)。INVR18原本用click_input(coords=...)
# 猜座標切分頁,之後的正式腳本一律改用這裡的switch_tab_by_ctrl_tab()。
# ============================================================

def is_really_visible(ctrl):
    """用pywinauto的is_visible()檢查控制項是不是「真的」在畫面上可見,
    不能只看find_by_class找不找得到——有些程式(IDLR68/IDLR69)的分頁內容
    不是懶加載的,一開始就找得到,只是不可見,一定要用這個才是正確驗證。"""
    try:
        return bool(ctrl.is_visible())
    except Exception:
        return False


def switch_tab_by_ctrl_tab(query_win, verify_ctrl, out_lines, max_presses=6):
    """對query_win送Ctrl+Tab切換分頁,每按一次就用verify_ctrl.is_visible()
    檢查有沒有切到目標分頁,直到成功或按到max_presses次上限為止。

    verify_ctrl請傳「目標分頁裡一定會出現的某個控制項」(例如目標分頁上
    第一個ComboBox或CheckBox)。回傳(switched, presses)。
    """
    query_win.set_focus()
    time.sleep(0.3)
    switched = is_really_visible(verify_ctrl)
    presses = 0
    while not switched and presses < max_presses:
        presses += 1
        log(f"  按第{presses}次 Ctrl+Tab...")
        query_win.type_keys("^{TAB}")
        time.sleep(0.5)
        switched = is_really_visible(verify_ctrl)
        if switched:
            log(f"  [成功] 按了{presses}次Ctrl+Tab之後,目標控制項變成可見了。")
    out_lines.append(f"Ctrl+Tab切換分頁: {'成功' if switched else '失敗'}(按了{presses}次)")
    return switched, presses


def send_key_sequence(win, keys, out_lines, label="", pause=0.3):
    """對win依序送出keys清單裡的每一個按鍵token(pywinauto type_keys語法,
    例如'^{RIGHT}'、'{TAB}'、' '(space)、'{DOWN}'、'{ENTER}'、'{ESC}'),
    每個之間停pause秒,純鍵盤操作、不找也不驗證任何控制項。

    2026-09-07新增,起因:INVR18原本用switch_tab_by_ctrl_tab()(Ctrl+Tab +
    is_visible()驗證)切到「進階選項」分頁,結果使用者實測發現除了WH01之外
    其他5組全部卡在「屬性類別選項」分頁——owner-drawn的分頁控制項(
    TDSUCExtPageControl)就算不是目前使用中的分頁,裡面的控制項is_visible()
    還是可能回傳True,靠這個驗證會誤判「已經切到對的分頁」,實際上沒有,
    導致後面找checkbox/combo抓錯東西或直接找不到。

    使用者人工操作驗證出來的可靠替代方案:不去驗證任何控制項狀態,單純按
    使用者實測過的固定按鍵次數(例如「Ctrl+Right按2次一定會從基本選項切到
    進階選項」「Tab後Space可以取消勾選當前欄位」「Down按2次可以選到選單
    第3個選項」),完全比照人工操作的按鍵順序、不做任何「這樣是不是對」的
    判斷——這是目前唯一在owner-drawn控制項上驗證有效的做法。
    """
    win.set_focus()
    time.sleep(0.3)
    for k in keys:
        win.type_keys(k, pause=0.05)
        time.sleep(pause)
    tag = f"[{label}] " if label else ""
    log(f"  {tag}已送出鍵盤序列: {keys}")
    out_lines.append(f"  {tag}鍵盤序列: {keys}")


# ============================================================
# 選擇庫別(或類似需要對照主檔代號的欄位):路線A/路線B
# ------------------------------------------------------------
# 路線A(區間選擇,單一代號):勾選panel上的「區間選擇」checkbox
# (class TDSYenCheckBox.UnicodeClass),清單框變成起/迄兩個MaskEdit,
# 起=迄=目標代號即可完成單選,不用碰F2。
#
# 路線B(F2,多代號):對清單控制項按F2鍵跳出獨立視窗,每個代號各自
# 獨立開一次F2(搜尋+勾選第一列+按確定)。**已證實**同一個F2視窗內連續
# 搜尋多個代號行不通(每次重查reQuery會把已勾選的狀態沖掉)。
# ============================================================

def select_single_via_range(select_panel, code, out_lines):
    """路線A:panel底下勾選「區間選擇」、填入起=迄=code。"""
    range_cb = None
    for cb in find_by_class(select_panel, "TDSYenCheckBox.UnicodeClass", "CheckBox"):
        try:
            t = (cb.window_text() or "").strip()
        except Exception:
            continue
        if t == "區間選擇":
            range_cb = cb
            break
    if range_cb is None:
        raise RuntimeError("panel底下找不到text=='區間選擇'的checkbox。")

    edits = find_by_class(select_panel, "TDSYenMaskEdit.UnicodeClass", "Edit")
    if len(edits) < 2:
        log("  「區間選擇」目前不是起/迄輸入框模式,點擊「區間選擇」切換...")
        range_cb.click()
        time.sleep(1)
        edits = find_by_class(select_panel, "TDSYenMaskEdit.UnicodeClass", "Edit")

    if len(edits) < 2:
        raise RuntimeError(
            f"勾選「區間選擇」之後,panel底下只找到{len(edits)}個MaskEdit(預期至少2個)。"
        )
    edits_sorted = sorted(edits, key=lambda c: c.rectangle().top)

    # 2026-09-07新增:9/7實測發現INVR18的WH05/WH06/wh08這幾組,「起」欄位
    # 常常沒填進去(只有「迄」有值),但當時log卻顯示兩次set_edit_text都
    # 執行過——research推測是剛從清單模式切成起/迄輸入框模式後,「起」欄位
    # 有時還沒完全就緒,set_edit_text()當下靜默沒生效(同一支腳本裡IDLR69
    # 的WH01/WH02、INVR18自己的WH01都沒踩到,顯示這是時好時壞的timing問題,
    # 不是每次都會發生)。**這不只是畫面好不好看的問題**——「起」空白+
    # 「迄」=code,在這個ERP的區間查詢邏輯下很可能代表「從最開頭一路查到
    # code」,範圍比預期的「只查code這一個庫別」大上很多,等於報表內容是
    # 錯的。所以這裡改成每次設定後讀回來驗證,沒對上就重設,最多重試3次,
    # 兩個欄位都確認寫對了才放行,寫不對就直接RuntimeError擋下來,不要讓
    # 這種錯誤範圍的查詢繼續送出去。
    def set_and_verify(edit_ctrl, label):
        actual = ""
        for attempt in range(3):
            edit_ctrl.set_edit_text(code)
            time.sleep(0.3)
            actual = (edit_ctrl.window_text() or "").strip()
            if actual == code:
                return actual
            log(f"  [警告] 「{label}」第{attempt + 1}次設定後讀回是{actual!r}"
                f"(預期{code!r}),重試...")
            time.sleep(0.3)
        return actual

    start_actual = set_and_verify(edits_sorted[0], "起")
    end_actual = set_and_verify(edits_sorted[1], "迄")

    log(f"  [路線A] 起={start_actual!r} 迄={end_actual!r} 填入完成。")
    out_lines.append(f"  [路線A] 起={start_actual!r} 迄={end_actual!r}")

    if start_actual != code or end_actual != code:
        raise RuntimeError(
            f"「選擇庫別」起/迄重試3次後還是設定失敗:起={start_actual!r} "
            f"迄={end_actual!r}(預期都是{code!r}),範圍會查錯,不送出這筆查詢。"
        )

    # 2026-09-07第2/5/6輪都在這裡加過「把焦點移開迄欄位」的動作(先是猜
    # 普通Tab、後來改成.set_focus()釘在區間選擇checkbox上),想解決Ctrl+
    # Right切不到「進階選項」分頁、「起」被清空的問題。第7輪(移除所有
    # 焦點操作,原地不動)也還是同樣失敗,使用者再測一次還是「起」被清空、
    # 卡在「基本選項」。
    #
    # 第8輪(這次)重新比對使用者的手動操作跟我們的自動化,抓到真正被忽略的
    # 差異:使用者是「真的用鍵盤在迄欄位打字」,打完之後鍵盤焦點自然、真實地
    # 停在這個欄位上,這時候按Ctrl+Right,Windows知道要把這個按鍵送去哪裡。
    # 但我們的`set_and_verify()`是用`set_edit_text()`(直接送WM_SETTEXT訊息
    # 改文字內容),**這個方式完全不會改變鍵盤焦點**——畫面上看起來欄位填對
    # 了,但整個查詢視窗的「目前鍵盤焦點在哪個控制項」這件事,其實從頭到尾
    # 都沒有被我們設定過,可能還停在視窗剛開啟時的預設控制項(例如某個按鈕
    # 或第1個欄位),跟「迄」欄位完全無關。之前第5/6輪的Tab/set_focus()想
    # 解決的其實不是「焦點卡在迄欄位裡」,而是反過來——根本沒有焦點在正確
    # 的地方,所以那兩次「移開」的方向本身就是錯的方向;第7輪什麼都不做,
    # 一樣是錯的,因為本來就沒有焦點在該在的地方。
    #
    # 這裡改成:填完值、驗證通過之後,明確把鍵盤焦點`.set_focus()`到「迄」
    # 這個欄位本身(不是移開,是移過去),重現使用者手動打完字之後、游標自然
    # 停在這裡的狀態,讓後面接著送的Ctrl+Right有正確的目標可以送。
    edits_sorted[1].set_focus()
    time.sleep(0.3)


def reverify_range_values(select_panel, code, out_lines, label="選擇庫別"):
    """路線A送出前的最後一道保險:重新讀回目前畫面上「起/迄」的實際值,
    跟code對不上就重設。用在select_single_via_range()設定完之後、還做過
    切分頁/設定進階選項等其他鍵盤操作,真正按確認送出之前——9/7實測發現
    即使select_single_via_range()當下驗證是對的,後續的Ctrl+Right切分頁
    操作偶爾還是會把「起」的值弄不見,補了Tab把焦點移開這個主要肇因之後,
    這裡再做最後一次「讀回來看看,不對就修」當保險,不管中間發生什麼都要
    在送出前守住這一關。

    這裡假設panel已經在起/迄輸入框模式(不會重新檢查/點擊「區間選擇」),
    因為呼叫這個函式的時候一定是select_single_via_range()已經成功切換過
    一次了。"""
    edits = find_by_class(select_panel, "TDSYenMaskEdit.UnicodeClass", "Edit")
    if len(edits) < 2:
        raise RuntimeError(
            f"[{label}] 送出前重新驗證時,只找到{len(edits)}個起/迄輸入框(預期至少2個)。"
        )
    edits_sorted = sorted(edits, key=lambda c: c.rectangle().top)

    def check_and_fix(edit_ctrl, field_label):
        actual = (edit_ctrl.window_text() or "").strip()
        if actual == code:
            return actual
        log(f"  [警告] [{label}·{field_label}] 送出前重新檢查,發現值變成{actual!r}"
            f"(預期{code!r}),重新設定...")
        out_lines.append(f"  [警告] [{label}·{field_label}] 送出前重新檢查發現值跑掉"
                          f"(變成{actual!r}),重新設定為{code!r}。")
        for attempt in range(3):
            edit_ctrl.set_edit_text(code)
            time.sleep(0.3)
            actual = (edit_ctrl.window_text() or "").strip()
            if actual == code:
                return actual
            time.sleep(0.3)
        return actual

    start_actual = check_and_fix(edits_sorted[0], "起")
    end_actual = check_and_fix(edits_sorted[1], "迄")
    out_lines.append(f"  [送出前重新驗證] 起={start_actual!r} 迄={end_actual!r}")
    log(f"  [送出前重新驗證] 起={start_actual!r} 迄={end_actual!r}")

    if start_actual != code or end_actual != code:
        raise RuntimeError(
            f"[{label}] 送出前重新驗證仍失敗:起={start_actual!r} 迄={end_actual!r}"
            f"(預期都是{code!r}),範圍會查錯,不送出這筆查詢。"
        )


def select_codes_via_f2(list_view, codes, out_lines,
                         header_guess=26, row_height_guess=23,
                         checkbox_x_fraction=0.06):
    """路線B:對list_view(TDSYenListView.UnicodeClass)按F2,codes清單裡每
    個代號各自獨立開一次F2視窗(搜尋+勾選第一列+按確定+關閉),依序累加。"""

    def all_top_windows():
        return safe_all_windows()

    for i, code in enumerate(codes):
        log(f"  [路線B] 第{i + 1}/{len(codes)}個代號 {code}(獨立開F2)...")
        known_handles = {w.handle for w in all_top_windows()}

        list_view.set_focus()
        list_view.type_keys("{F2}", pause=0.05)
        time.sleep(2)

        candidates = [
            w for w in all_top_windows()
            if w.handle not in known_handles and w.class_name() == "TfrmF2Window.UnicodeClass"
        ]
        if not candidates:
            raise RuntimeError(f"代號{code}: 按F2後沒有找到F2視窗(TfrmF2Window.UnicodeClass)。")
        popup = candidates[0]

        search_edits = find_by_class(popup, "TDSUCMaskEdit.UnicodeClass", "Edit")
        if not search_edits:
            raise RuntimeError(f"代號{code}: F2視窗裡沒找到重查搜尋框。")
        search_edit = search_edits[0]

        grids = find_by_class(popup, "TDSUCwwDBGrid", "Grid")
        if not grids:
            raise RuntimeError(f"代號{code}: F2視窗裡沒找到grid。")
        grid = grids[0]

        search_edit.set_edit_text(code)
        time.sleep(0.3)
        search_edit.type_keys("{ENTER}", pause=0.1)
        time.sleep(1.2)

        rect = grid.rectangle()
        checkbox_x = int(rect.width() * checkbox_x_fraction)
        checkbox_y = header_guess + int(row_height_guess * 0.5)
        grid.click_input(coords=(checkbox_x, checkbox_y))
        time.sleep(0.5)

        confirm_btn = find_button(popup, "確定")
        if confirm_btn is None:
            raise RuntimeError(f"代號{code}: F2視窗裡沒找到確定按鈕。")
        confirm_btn.click()
        time.sleep(1)
        log(f"    代號{code}完成。")
        out_lines.append(f"  [路線B] 代號{code}: F2搜尋+勾選+確定 完成")


# ============================================================
# 送出查詢(真的按確認,產生真實ERP工作)
# ------------------------------------------------------------
# 從part2_submit_test.py抽出來的驗證過邏輯,供正式的report_*.py呼叫。
# ============================================================

def find_confirm_button(query_win):
    btns = find_by_class(query_win, "TDSUCBitBtn.UnicodeClass", "Button")
    for b in btns:
        try:
            text = (b.window_text() or "").strip()
        except Exception:
            continue
        if text == "確認":
            return b
    raise RuntimeError("查詢畫面裡沒有找到文字為「確認」的按鈕。")


def submit_query(query_win, code, out_lines, confirm_wait_timeout=20):
    """按下查詢畫面的「確認」按鈕,真的送出查詢(產生真實ERP工作),然後
    掃整個Desktop抓新冒出來的訊息視窗,自動按OK關閉。

    訊息視窗裡「伺服器已收到您的需求,並已登錄為XXX!」這段文字是owner-drawn
    畫上去的,讀不到,所以這裡**不會**回傳工作編號——Stage B(轉Excel)直接從
    「佇列工作管理員」清單去對應剛送出的工作(最新的永遠排最上面)。
    """

    def all_top_windows():
        return safe_all_windows()

    known_handles = {w.handle for w in all_top_windows()}

    confirm_btn = find_confirm_button(query_win)
    log("  找到「確認」按鈕,按下去送出查詢...")
    confirm_btn.click()

    log(f"  已按下確認,等待跳出訊息視窗(逾時{confirm_wait_timeout}秒)...")
    popup = None
    deadline = time.time() + confirm_wait_timeout
    while time.time() < deadline and popup is None:
        new_windows = [w for w in all_top_windows() if w.handle not in known_handles]
        if new_windows:
            matched = [w for w in new_windows if code in (w.window_text() or "")]
            popup = matched[0] if matched else new_windows[0]
        if popup is None:
            time.sleep(0.5)

    if popup is None:
        raise RuntimeError(
            f"按下確認後{confirm_wait_timeout}秒內沒有偵測到任何新視窗。"
            "請截圖回報ERP畫面目前的實際狀況。"
        )

    log(f"  偵測到新視窗: class={popup.class_name()!r} text={popup.window_text()!r}")
    out_lines.append(
        f"  送出查詢後跳出視窗: class={popup.class_name()!r} text={popup.window_text()!r}"
    )

    ok_btn = None
    for c in find_by_class(popup, "TDSUCButton.UnicodeClass", "Button"):
        if (c.window_text() or "").strip().upper() == "OK":
            ok_btn = c
            break
    if ok_btn is None:
        log("  [警告] 跳出視窗裡沒找到文字為OK的按鈕,不會自動關閉,請人工確認畫面。")
        out_lines.append("  [警告] 沒有自動按OK,請人工確認畫面。")
    else:
        ok_btn.click()
        time.sleep(1)
        log("  已按OK關閉訊息視窗。")
        out_lines.append("  已按OK關閉訊息視窗,查詢送出流程完成。")

    # 2026-09-08新增,2026-09-07(第2輪)修正:關閉訊息視窗之後,查詢畫面
    # 本身**不會自動關掉**——使用者實測發現IDLR69、INVR18跑完之後查詢畫面還
    # 留在畫面上,手動按ESC就能關掉。第1版在這裡用query_win.exists()判斷,
    # 但query_win是從Desktop(backend="win32").windows()拿到的**具體控制項
    # 物件(HwndWrapper/DialogWrapper)**,不是pywinauto的WindowSpecification
    # 代理物件,沒有.exists()這個方法——呼叫就丟AttributeError,被下面的
    # except吃掉變成一句「可忽略」的警告,實際上ESC**根本沒送出去過**,
    # 這就是使用者回報「還是沒關掉」的真正原因(同一種錯誤先前在
    # dump_control_tree呼叫print_control_identifiers()時也出現過,是同一個
    # 物件型別問題,只是這次沒有備援邏輯把它蓋過去)。
    #
    # 改用win32gui.IsWindow(handle)直接查詢窗控代碼是否還存在,這個函式對
    # 任何win32視窗代碼都能用,不依賴pywinauto包裝物件的類型。
    try:
        import win32gui
        if win32gui.IsWindow(query_win.handle):
            query_win.set_focus()
            time.sleep(0.3)
            query_win.type_keys("{ESC}")
            time.sleep(0.8)
            still_there = win32gui.IsWindow(query_win.handle)
            log(f"  已送ESC關閉查詢畫面(關閉{'失敗,仍存在' if still_there else '成功'})。")
            out_lines.append(f"  已送ESC關閉查詢畫面(仍存在={still_there})。")
        else:
            log("  查詢畫面視窗代碼已經不存在,不用送ESC。")
            out_lines.append("  查詢畫面已不存在,不用送ESC。")
    except Exception as e:
        log(f"  [警告] 關閉查詢畫面時發生例外(可能本來就已經關了,可忽略): {e!r}")
        out_lines.append(f"  [警告] 關閉查詢畫面時發生例外(可忽略): {e!r}")

    return popup


def main():
    parser = argparse.ArgumentParser(description="Part 2 控制項偵察工具(唯讀,不會送出任何查詢)")
    parser.add_argument("--code", default="IDLR17",
                         help="測試用的程式代號,預設IDLR17(已確認不需要任何選擇、安全)")
    parser.add_argument("--output", default=OUTPUT_FILE, help="輸出文字檔路徑")
    parser.add_argument("--skip-code", action="store_true",
                         help="只做登入+主畫面傾印,不測試打代號跳轉這一步")
    args = parser.parse_args()

    out_lines = []
    try:
        app, main_win = login(out_lines)
        if not args.skip_code:
            jump_to_code(main_win, args.code, out_lines)
    except Exception as e:
        out_lines.append("")
        out_lines.append("!" * 70)
        out_lines.append(f"執行中發生錯誤: {e!r}")
        out_lines.append("!" * 70)
        log(f"[錯誤] {e!r}")
        raise
    finally:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write("\n".join(out_lines))
        log(f"\n控制項結構已寫入: {args.output}")
        log("請把這個檔案傳給我(或貼出內容),我再照真實的control class name/id寫後續腳本。")
        log("(這支腳本執行到這裡就停了,不會按確認送出查詢,也不會關閉ERP——"
            "你可以自己看一下畫面目前停在哪裡,確認跟預期一致。)")


if __name__ == "__main__":
    sys.exit(main())
