CMP Module 3 CCL Mapping V1.1 覆蓋說明

目的：
1. CCL 係數對應改為背景 Job，避免長時間 POST 造成 Failed to fetch。
2. CCL 處理中畫面改用與 Module 1 / Module 2 相同的 processing-status + progress bar 格式。
3. CCL Mapping 保留 Dictionary Index：先建立 Material → Factor 對應表，再逐列寫入 raw material bulk，避免重複掃描對照表。
4. 前端輪詢 /module3/ccl-job/{job_id}，完成後才開啟下載連結。
5. 前端新增 JSON 檢查，若後端回傳 HTML 錯誤頁，顯示「伺服器回傳錯誤頁面，請查看 Render log」，不再只顯示 Unexpected token '<'。

覆蓋檔案：
- main.py
- factor_selector.py
- templates/index.html
- data/factor_library/factors.db
- README/README_MODULE3_SQLITE_INDEX_V1.txt
- README/README_MODULE3_CCL_JOB_PROGRESS_V1.txt

新增 API：
- POST /module3/apply-ccl-factors-job
- GET  /module3/ccl-job/{job_id}

保留舊 API：
- POST /module3/apply-ccl-factors

注意：
- factors.db 為 Module 3 Factor Search Center 的 SQLite 索引檔，需與 factor_selector.py 一起覆蓋。
- 若 Render instance 重啟，進行中的背景工作會中斷，使用者需重新執行該次 CCL 對應。
