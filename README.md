# Annual Output Platform v6

判斷順序：

1. `rule_master.csv` 依 Priority 由小到大判斷
2. `product_series_master.csv` 補充分類
3. Default = WIP

啟動：

```powershell
python -m pip install -r requirements.txt
python -m uvicorn main:app --host 0.0.0.0 --port 10000
```

網址：

```text
http://localhost:10000
```


## v6 更新
- Product Series 解析改成雙軌：先保留逗號/分號/底線切段，再全文 Regex Prefix Search。
- Series Prefix 從 data/rule_master.csv 的 Series Prefix 動態讀取。
- 可正確處理：AssyCH SN5372BL、Assy,CHSN4396BL1、SP2B20XF0AssyHapticForce Touchpad module。


## v6 updates
- 支援一次上傳多個 SAP 生產工單 Excel。
- 平台名稱：Annual Output Platform v6。
- Rule Master 新增 Touch pad module → TP、SCMC → WIP。
- Product Series 抽取保留標點切段，並使用全文 Regex Prefix Search。
