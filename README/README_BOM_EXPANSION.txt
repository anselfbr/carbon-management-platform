# CMP BOM Expansion v1

本包建立平台第二部分：2. BOM Expansion

覆蓋：
- templates/index.html
- static/style.css

新增：
- bom_formatter.py

修改 main.py：
- 將 main_bom_patch.py 內容貼到 main.py 最下方
- 不建議直接覆蓋 main.py，避免覆蓋掉你目前已完成的 Step 1 / Step 2 endpoint

平台畫面：
- Upload Standard BOM
- Upload Raw Material Bulk Template
- Process BOM Expansion
- Progress bar / processing stage / estimated remaining time
- Download Raw Material Bulk

BOM 欄位定義：
- Parent Material = Parent Node
- Component = Component
- Quantity = CS03 Qty
- Unit = CS03 UoM

半品判斷：
- Component 同時存在於 Parent Node => 半品，繼續往下展開

Raw Material Bulk 寫入：
- Input Sheet Activity Data
  - A Raw Material Name = 展開後最終原物料
  - B Doc. Start Date = BOM Valid From / valid from
  - C Doc. End Date = 同年度 12/31
  - D Document Type = BOM
  - F Usage = 多階累積用量
  - G Activity Data Unit = CS03 UoM
  - K Data Source = SAP
  - P Allocated Target Product/Service = 最上層 Parent Product

- Input Sheet Raw Material
  - A Raw Material Name = 原物料
  - B Raw Material Code = 原物料
  - F Raw Material Description = Component Description
