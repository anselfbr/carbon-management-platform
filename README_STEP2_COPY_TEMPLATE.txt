# cmp_step2_bulk_formatting_v1_3_COPY_TEMPLATE

本版改為方案 1：
直接複製原始 Bulk Template，再只覆蓋下列兩個分頁的儲存格內容：

1. Input Sheet Activity Data
2. Input Sheet Products

不重新建立 Workbook。
不刪除分頁。
不重建格式。
不改動 Instructions / Dropdown Values 等其他分頁。

請覆蓋 / 新增：
- main.py
- bulk_formatter.py
- templates/index.html
- static/style.css

重點修改：
- bulk_formatter.py 先使用 shutil.copy2() 複製原始 template
- 再用 openpyxl 開啟複製後檔案
- 僅清除並覆蓋：
  - Input Sheet Activity Data：A:H，自第 3 列起
  - Input Sheet Products：A、D、F，自第 3 列起

寫入內容：
- WIP 不寫入
- Activity Data:
  A Product Name = Material Number
  B Doc. Start Date = YYYY/01/01
  C Doc. End Date = YYYY/12/31
  D Product Type = Target Product
  E Production Site:
    NB = 常州廠(A2)-IPS
    TP = 常州廠(A9)-IPS
    其他 = 石碣廠-IPS
  F Production/ Service Quantity = 年度生產量
  G Data Source = SAP
  H Data Source Other = blank

- Products:
  A Product Name = Material Number
  D System Boundary = Cradle-to-Gate
  F Declared Unit = PC
