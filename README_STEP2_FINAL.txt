# cmp_step2_bulk_formatting_v1_2_FINAL

這版修正 Render 404 Not Found 問題。

原因：
Render log 顯示 POST /generate-bulk-file = 404，
代表 main.py 沒有建立 /generate-bulk-file endpoint。

請覆蓋 / 新增：
1. main.py
2. bulk_formatter.py
3. templates/index.html
4. static/style.css

部署後測試：
1. Render Manual Deploy > Deploy latest commit
2. 開啟網站
3. Step 2 上傳：
   - Step 1 Output: 年度產品產量與分類結果_v6_ALL...
   - Bulk Template: product_activity_data_bulk_create_template_v1...
4. 點 Generate Formatted Bulk File

輸出規則：
- WIP 不寫入
- Activity Data:
  Product Name = Material Number
  Doc. Start Date = YYYY/01/01
  Doc. End Date = YYYY/12/31
  Product Type = Target Product
  Production Site:
    NB = 常州廠(A2)-IPS
    TP = 常州廠(A9)-IPS
    其他 = 石碣廠-IPS
  Production/ Service Quantity = 年度生產量
  Data Source = SAP
- Products:
  Product Name = Material Number
  System Boundary = Cradle-to-Gate
  Declared Unit = PC
