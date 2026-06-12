# cmp_step2_bulk_formatting_v1_4_PRODUCT_DESCRIPTION

本版延續方案 1：
直接複製原始 Bulk Template，再只覆蓋下列兩個分頁的儲存格內容：

1. Input Sheet Activity Data
2. Input Sheet Products

不重新建立 Workbook。
不刪除分頁。
不重建格式。
不改動 Instructions / Dropdown Values 等其他分頁。
原本 bulk template 的下拉選項可保留；前提是 template 的 Data Validation 原本就涵蓋寫入列數。

本版新增：
- Input Sheet Products 的 C 欄 Product Description
  = Step1 Output 的 Material Description

請覆蓋 / 新增：
- main.py
- bulk_formatter.py
- templates/index.html
- static/style.css

寫入內容：

Activity Data：
- A Product Name = Material Number
- B Doc. Start Date = YYYY/01/01
- C Doc. End Date = YYYY/12/31
- D Product Type = Target Product
- E Production Site:
  - NB = 常州廠(A2)-IPS
  - TP = 常州廠(A9)-IPS
  - 其他 = 石碣廠-IPS
- F Production/ Service Quantity = 年度生產量
- G Data Source = SAP
- H Data Source Other = blank

Products：
- A Product Name = Material Number
- C Product Description = Material Description
- D System Boundary = Cradle-to-Gate
- F Declared Unit = PC

備註：
- WIP 不寫入 bulk file
- 不處理 Material Number 重複，因 Step1 Output 已去除重複值
