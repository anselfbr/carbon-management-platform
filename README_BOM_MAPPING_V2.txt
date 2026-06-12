# CMP BOM Expansion Mapping v2

本版新增 BOM Mapping Configuration 可編輯空白欄位。

覆蓋：
- templates/index.html
- static/style.css
- bom_formatter.py
- main.py（若你目前 main.py 是這次上傳的版本，可直接覆蓋）

畫面新增空白輸入欄位：
- Parent Material
- Component
- Quantity
- Unit
- Component Description
- Material Group
- Valid From

欄位可留空，後端會使用預設：
- Parent Node
- Component
- CS03 Qty
- CS03 UoM
- Component Description
- Material group
- BOM Valid From
