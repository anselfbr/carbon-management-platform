CMP Module 3 SQLite Index V1 - 2026/07/04

覆蓋內容：
1. factor_selector.py
   - Module 3 Factor Search 改查 SQLite factors.db，不再把 APOS/Cut-off Excel 整包載入 Python list cache。
   - 保留原本 /module3/search-factor-library API 參數，前端介面不需要修改。
   - 保留 APOS 優先、Cut-off 次之的排序。
   - 保留 Geography 與 Source 篩選。
   - 僅生產 production_only 判定改為 Activity Name 包含 production。

2. data/factor_library/factors.db
   - 已由目前 ZIP 內 APOS / Cut-off LCIA Excel 建立。
   - 筆數：53,730 rows。
   - 若未來 APOS / Cut-off Excel 更新，系統會偵測檔案大小與修改時間，並自動重建 index。

未修改：
- templates/index.html
- static/style.css
- static/i18n.js
- main.py
- Module 1 / Module 2 相關檔案

部署方式：
直接將覆蓋檔內相同路徑檔案覆蓋到專案。
