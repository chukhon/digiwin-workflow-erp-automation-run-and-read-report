#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
共用旗標——目前只有一個:SCREENSHOTS_ENABLED。

2026-09-08新增,起因:user要求「auto weekly」正式版不要抓圖、不要產出
報告檔案,但同時想保留run_full_pipeline.py(會抓圖、會產生
run_full_pipeline_output.txt)當作除錯/驗證用的版本,不要動它的行為。

「不產出報告檔案」這件事,不需要旗標——run_full_pipeline.py跟
part5_paste_reports.py都已經改成:只有各自的`main()`(獨立執行的
命令列入口)才會真的把記錄寫成output.txt檔案,可重用的核心函式
(`run_pipeline()`/`run()`)本身不寫檔案。auto_weekly.py直接呼叫這些
核心函式、不呼叫`main()`,自然就不會產生任何output.txt。

但「不抓圖」沒辦法用同樣的方式處理,因為screenshot是在流程中途
(report_*.py的run()、part4_batch_export_excel.py的process_one_job())
就會呼叫,不是只在最後才發生的動作,沒辦法用「呼叫核心函式而不是
main()」來自動避開。所以這裡用一個全域旗標,`save_screenshot()`
(part2_inspect.py)跟`save_shot()`(part2_reportview_probe.py)——
目前所有腳本裡的截圖動作,都是透過這兩個函式做的——在動手截圖之前先
檢查這個旗標,是False就直接跳過。

預設True,代表report_*.py、run_all_reports.py、
part4_batch_export_excel.py、run_full_pipeline.py不管是單獨執行還是被
run_full_pipeline.py呼叫,只要沒有人主動把這個旗標改成False,行為都
跟這次修改之前完全一樣。只有auto_weekly.py會在一開始把它設成False。
"""

SCREENSHOTS_ENABLED = True
