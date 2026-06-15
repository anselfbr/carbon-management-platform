/* CMP i18n v1.1
   Safe implementation:
   - Standalone file
   - Does not modify existing platform functions
   - Does not overwrite existing click handlers
   - Uses MutationObserver to translate texts created dynamically by the original page JS
*/
(function () {
  const button = document.getElementById("langToggle");
  if (!button) return;

  const STORAGE_KEY = "cmp_lang";
  let currentLang = localStorage.getItem(STORAGE_KEY) || "en";
  let isApplying = false;

  const zh = {
    "Carbon Management Platform | Product Data Preparation": "碳管理平台 | 產品資料準備",
    "Reading labor work orders": "讀取生產工時工單",
    "Labor HR.Act + FOH-Others.Act": "人員+設備工時",
    "FOH-Others.Act Only": "設備工時",
    "Labor HR.Act Only": "人員工時",
    "Labor Allocation Source": "工時來源",
    "Working Hour Source": "工時來源",
    "Upload Working Hour Orders": "生產工時工單",
    "Production Quantity Work Orders": "生產數量工單",

    "Carbon Management Platform": "碳管理平台",
    "Manufacturing Data Layer": "製造資料層",
    "Modules": "模組",
    "Product Data Preparation": "產品資料準備",
    "Rule Management": "規則管理",
    "BOM Expansion": "BOM 展開",
    "Carbon Emission Factor Selection": "碳排放係數選擇",
    "PCF Calculation": "產品碳足跡計算",
    "Production Environment": "正式環境",
    "Online | Version 1.0": "線上｜版本 1.0",
    "Rule Master Enabled": "Rule Master 已啟用",
    "Multi-file Upload": "多檔上傳",
    "Version 1.0": "版本 1.0",

    "Prepare manufacturing data, expand BOM structures, select emission factors, and calculate product carbon footprints.": "準備製造資料、展開 BOM 結構、選擇碳排放係數，並計算產品碳足跡。",
    "Prepare manufacturing data, expand BOM structures, select emission factors, and calculate product carbon footprints.": "準備製造資料、展開 BOM 結構、選擇碳排放係數，並計算產品碳足跡。",
    "Prepare production output and batch data for product carbon footprint workflows.": "準備產品碳足跡流程所需的生產產出與批次資料。",
    "Maintain Product Data Preparation rules, including Rule Master and Product Series Master.": "維護產品資料準備規則，包含 Rule Master 與 Product Series Master。",
    "Expand multi-level BOM structures, roll up raw material quantities, and generate raw material bulk files.": "展開多階 BOM 結構、彙總原物料數量，並產生原物料 Bulk 檔。",

    "Module": "模組",
    "Work order & batch preparation": "工單與批次資料準備",
    "Multi-level BOM explosion": "多階 BOM 展開",
    "Factor mapping and selection": "係數對應與選擇",
    "Product carbon footprint": "產品碳足跡",

    "Step 1 · Work Order Processing": "步驟 1 · 工單處理",
    "Upload one or multiple SAP production work order files.": "上傳一份或多份 SAP 生產工單檔案。",
    "Step 1": "步驟 1",
    "Upload SAP Production Work Orders": "生產數量工單",
    "Reporting Year": "報告年度",
    "Run Consolidation & Classification": "執行合併與分類",
    "Ready for processing": "準備處理",
    "Upload SAP work orders and start the classification workflow.": "上傳 SAP 工單並開始分類流程。",
    "Upload production quantity work orders and optional labor work orders to start the classification workflow.": "上傳生產數量工單與選填生產工時工單後開始分類流程。",
    "Idle": "待命",
    "Download Step 1 Output Excel": "下載 Step 1 輸出 Excel",

    "Step 2 · Batch Data Formatting": "步驟 2 · 批次資料格式化",
    "Convert Step 1 output and batch file template into a standardized batch file.": "將 Step 1 輸出與批次範本轉換為標準化批次檔。",
    "Step 2": "步驟 2",
    "Step 1 Output": "Step 1 輸出",
    "Annual output & classification result": "年度產量與分類結果",
    "Batch Template": "批次範本",
    "Required batch file format": "必要批次檔格式",
    "Formatted Batch File": "已格式化批次檔",
    "Ready for downstream processing": "可供後續處理",
    "Upload Step 1 Output File": "上傳 Step 1 輸出檔",
    "Upload Bulk Template File": "上傳 Bulk 範本檔",
    "Generate Formatted Bulk File": "產生格式化 Bulk 檔",
    "Ready for batch formatting": "準備批次格式化",
    "Upload Step 1 output and bulk template to generate formatted bulk file.": "上傳 Step 1 輸出與 Bulk 範本以產生格式化檔案。",
    "Download Formatted Bulk File": "下載格式化 Bulk 檔",
    "Step 2 extracts required fields from the Step 1 output and writes them into the bulk template.": "Step 2 會從 Step 1 輸出擷取必要欄位並寫入 Bulk 範本。",
    "Activity Data and Products sheets will be populated automatically.": "系統會自動填入 Activity Data 與 Products 分頁。",

    "Execution Log": "執行紀錄",
    "Summary": "摘要",
    "Rules": "規則",
    "Ready. Upload SAP production work orders and start processing.": "準備完成。請上傳 SAP 生產工單並開始處理。",
    "Ready. Upload production quantity work orders and optional labor work orders to start processing.": "準備完成。請上傳生產數量工單與選填生產工時工單後開始處理。",
    "Version 1.0 Decision Flow": "版本 1.0 判斷流程",
    "Product Series Engine": "產品系列引擎",
    "punctuation-based segmentation, then full-text Regex Prefix Search.": "以標點符號分段，再進行全文 Regex 前綴搜尋。",
    "Rule Master": "規則主檔",
    "priority-based classification logic.": "以優先順序為基礎的分類邏輯。",
    "Product Series Master": "產品系列主檔",
    "fallback mapping by product series.": "依產品系列進行備援對應。",
    "Default WIP": "預設 WIP",
    "classify as WIP if no rule is matched.": "若沒有命中任何規則，則分類為 WIP。",
    "Required SAP Fields": "必要 SAP 欄位",

    "Upload Standard BOM": "上傳標準 BOM",
    "Upload Raw Material Bulk Template": "上傳原物料 Bulk 範本",
    "Process BOM Expansion": "執行 BOM 展開",
    "Ready for BOM Expansion": "準備 BOM 展開",
    "Upload standard BOM and raw material bulk template to start processing.": "上傳標準 BOM 與原物料 Bulk 範本後開始處理。",
    "Download Raw Material Bulk": "下載原物料 Bulk",
    "BOM Expansion Logic": "BOM 展開邏輯",
    "Configure source column names for BOM Expansion.": "設定 BOM 展開來源欄位名稱。",
    "Use": "使用",
    "for system settings or": "套用系統設定，或使用",
    "to apply your input.": "套用輸入內容。",
    "Parent Material": "母件料號",
    "Component": "元件料號",
    "Quantity": "數量",
    "Unit": "單位",
    "Component Description": "元件描述",
    "Material Group": "物料群組",
    "Valid From": "有效起始日",
    "Default": "預設",
    "Confirm": "確認",
    "Current Setting:": "目前設定：",
    "Parent Node": "Parent Node",
    "CS03 Qty": "CS03 Qty",
    "CS03 UoM": "CS03 UoM",
    "Material group": "Material group",
    "BOM Valid From": "BOM Valid From",
    "Semi-finished Rule": "半成品判斷規則",
    "Component also exists as Parent Node": "Component 同時存在於 Parent Node",
    "Quantity Roll-up": "數量展開邏輯",
    "Multiply quantities across all BOM levels": "跨 BOM 階層累乘數量",
    "Raw Material Bulk Output": "原物料 Bulk 輸出",
    "Input Sheet Activity Data and Input Sheet Raw Material will be populated.": "系統會填入 Input Sheet Activity Data 與 Input Sheet Raw Material。",
    "Optional fields are not written.": "Optional 欄位不會寫入。",
    "BOM Expansion is ready.": "BOM 展開已準備完成。",

    "Import and download classification rule masters for Product Data Preparation.": "匯入與下載產品資料準備使用的分類規則主檔。",
    "Rule Management is a sub-function under Product Data Preparation.": "規則管理是產品資料準備下的子功能。",
    "It controls product type classification, product series mapping, WIP judgment and customer mapping logic.": "用於控制產品類型分類、產品系列對應、WIP 判斷與客戶對應邏輯。",
    "Upload Rule Master": "上傳 Rule Master",
    "Import Rule Master": "匯入 Rule Master",
    "Download Rule Master": "下載 Rule Master",
    "Download Product Series Master": "下載 Product Series Master",
    "Rule Master Fields": "Rule Master 欄位",
    "Default Rules": "預設規則",
    "Rule Management is ready.": "規則管理已準備完成。",

    "1. Product Data Preparation": "1. 產品資料準備",
    "Complete Step 1 Work Order Processing, Step 2 Batch Data Formatting and Rule Management.": "完成 Step 1 工單處理、Step 2 批次資料格式化與規則管理。",
    "2. BOM Expansion": "2. BOM 展開",
    "Reserved module. This area can be extended for multi-level BOM explosion.": "預留模組。此區可延伸為多階 BOM 展開。",
    "3. Carbon Emission Factor Selection": "3. 碳排放係數選擇",
    "Reserved module. This area can be extended for emission factor mapping and selection.": "預留模組。此區可延伸為排放係數對應與選擇。",
    "4. PCF Calculation": "4. 產品碳足跡計算",
    "Reserved module. This area can be extended for product carbon footprint calculation.": "預留模組。此區可延伸為產品碳足跡計算。",

    "Processing...": "處理中...",
    "Reading Excel files": "讀取 Excel 檔案",
    "Merging work orders": "合併工單",
    "Filtering year": "篩選年度",
    "Annual output summary": "年度產量彙總",
    "WIP decision": "WIP 判斷",
    "Excel export": "Excel 匯出",
    "Completed.": "已完成。",
    "Year": "年度",
    "Uploaded files": "上傳檔案數",
    "Labor files": "工時檔案數",
    "Total hours": "總工時",
    "Uploaded Files": "上傳檔案數",
    "Labor Files": "工時檔案數",
    "Work order rows": "工單筆數",
    "Work Order Rows": "工單筆數",
    "Summary rows": "彙總筆數",
    "Summary Rows": "彙總筆數",
    "Total output": "總產量",
    "Total Output": "總產量",
    "Total Hours": "總工時",
    "WIP rows": "WIP 筆數",
    "WIP Rows": "WIP 筆數",
    "Processing completed": "處理完成",
    "Classification result is ready. You can download the Excel output.": "分類結果已完成，可下載 Excel 輸出檔。",
    "Completed": "已完成",
    "Processing failed": "處理失敗",
    "Please review the error message and input file format.": "請確認錯誤訊息與輸入檔格式。",

    "BOM Expansion in progress": "BOM 展開處理中",
    "Reading standard BOM and validating raw material bulk template.": "讀取標準 BOM 並檢查原物料 Bulk 範本。",
    "Reading BOM structure...": "讀取 BOM 結構...",
    "Detecting semi-finished components...": "判斷半成品元件...",
    "Expanding multi-level BOM...": "展開多階 BOM...",
    "Calculating material quantity roll-up...": "計算物料數量累乘...",
    "Copying raw material bulk template...": "複製原物料 Bulk 範本...",
    "Writing Activity Data and Raw Material sheets...": "寫入 Activity Data 與 Raw Material 分頁...",
    "BOM Expansion completed": "BOM 展開完成",
    "BOM Expansion failed": "BOM 展開失敗",
    "Please review the BOM and raw material bulk template.": "請檢查 BOM 與原物料 Bulk 範本。",
    "Processing BOM Expansion...": "正在處理 BOM 展開...",

    "Batch formatting in progress": "批次格式化處理中",
    "Reading Step 1 output and validating the bulk template.": "讀取 Step 1 輸出並檢查 Bulk 範本。",
    "Reading Step 1 output file...": "讀取 Step 1 輸出檔...",
    "Copying original bulk template...": "複製原始 Bulk 範本...",
    "Writing Input Sheet Activity Data...": "寫入 Input Sheet Activity Data...",
    "Writing Input Sheet Products...": "寫入 Input Sheet Products...",
    "Preserving template formatting and validation...": "保留範本格式與資料驗證...",
    "Generating formatted bulk file...": "產生格式化 Bulk 檔...",
    "Batch formatting completed": "批次格式化完成",
    "Batch formatting failed": "批次格式化失敗",
    "Please review the input files and template format.": "請檢查輸入檔案與範本格式。",

    "Please select a rule file first.": "請先選擇規則檔。",
    "Importing Rule Master...": "正在匯入 Rule Master...",
    "Rule Master imported successfully.": "Rule Master 匯入成功。",
    "Total rules:": "規則總數：",
    "Import failed": "匯入失敗"
  };

  const en = {};
  Object.keys(zh).forEach(function (key) {
    en[zh[key]] = key;
  });

  const preserveExact = new Set([
    "CMP", "SAP", "BOM", "PCF", "WIP", "NB", "TP", "SCMC", "SN", "SP", "SM", "SK",
    "Order、Plant、Material Number、Material description、Delivered quantity (GMEIN)、Actual finish date",
    "Priority、Rule Type、Key、Product Type、Customer、Customer Code Logic、Is_WIP、Enabled",
    "851- / 852- prefix → WIP",
    "Material description contains Touch pad module → TP",
    "SCMC prefix → default WIP, can be overridden by description rules",
    "Material description contains ASSY → WIP",
    "SN → NB；SP → TP；SM → DT Mouse；SK → DT Keyboard"
  ]);

  function isSkippableText(text) {
    const t = text.trim();
    if (!t) return true;
    if (preserveExact.has(t)) return true;
    if (/^~?\d+s remaining$/.test(t)) return true;
    return false;
  }

  function getDict(targetLang) {
    return targetLang === "zh" ? zh : en;
  }

  function translateString(input, targetLang) {
    if (!input || isSkippableText(input)) return input;

    const dict = getDict(targetLang);
    let output = input;

    // Exact normalized match first
    const normalized = input.replace(/\s+/g, " ").trim();
    if (dict[normalized]) {
      return input.replace(input.trim(), dict[normalized]);
    }

    // Phrase replacement, longest first
    const keys = Object.keys(dict).sort(function (a, b) { return b.length - a.length; });
    keys.forEach(function (key) {
      const val = dict[key];
      output = output.split(key).join(val);
    });

    return output;
  }


  function normalizeWorkingHourTexts() {
    const isZh = currentLang === "zh";

    document.querySelectorAll("label").forEach(function (label) {
      const t = label.textContent.replace(/\s+/g, " ").trim();
      if (
        t === "Labor Allocation Source" ||
        t === "Working Hour Source" ||
        t === "工時擷取選項" ||
        t === "工時來源"
      ) {
        label.textContent = isZh ? "工時來源" : "Working Hour Source";
      }
    });

    const select = document.getElementById("laborMode");
    if (!select) return;

    const optionText = isZh
      ? {
          both: "人員+設備工時",
          labor_hr: "人員工時",
          foh: "設備工時"
        }
      : {
          both: "Labor HR.Act + FOH-Others.Act",
          labor_hr: "Labor HR.Act Only",
          foh: "FOH-Others.Act Only"
        };

    Array.from(select.options).forEach(function (option) {
      if (optionText[option.value]) {
        option.textContent = optionText[option.value];
      }
    });
  }

  function translateTextNodes(root, targetLang) {
    const walker = document.createTreeWalker(root || document.body, NodeFilter.SHOW_TEXT, {
      acceptNode: function (node) {
        const parent = node.parentElement;
        if (!parent) return NodeFilter.FILTER_REJECT;
        if (parent.closest("script, style")) return NodeFilter.FILTER_REJECT;
        if (parent.closest("#laborMode")) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });

    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);

    nodes.forEach(function (node) {
      node.nodeValue = translateString(node.nodeValue, targetLang);
    });
  }

  function translateAttributes(targetLang) {
    document.querySelectorAll("[placeholder]").forEach(function (el) {
      el.setAttribute("placeholder", translateString(el.getAttribute("placeholder"), targetLang));
    });

    document.querySelectorAll("[aria-label]").forEach(function (el) {
      if (el.id === "langToggle") return;
      el.setAttribute("aria-label", translateString(el.getAttribute("aria-label"), targetLang));
    });

    if (document.title) {
      document.title = translateString(document.title, targetLang);
    }
  }

  function applyLanguage(targetLang) {
    if (isApplying) return;
    isApplying = true;

    currentLang = targetLang;
    localStorage.setItem(STORAGE_KEY, currentLang);
    document.documentElement.lang = currentLang === "zh" ? "zh-Hant" : "en";
    button.textContent = currentLang === "zh" ? "EN" : "中";

    translateTextNodes(document.body, currentLang);
    translateAttributes(currentLang);
    normalizeWorkingHourTexts();

    isApplying = false;
  }

  button.addEventListener("click", function (event) {
    event.preventDefault();
    event.stopPropagation();
    applyLanguage(currentLang === "en" ? "zh" : "en");
  });

  // Translate future text generated by existing page scripts.
  const observer = new MutationObserver(function (mutations) {
    if (isApplying || currentLang !== "zh") return;

    window.setTimeout(function () {
      mutations.forEach(function (mutation) {
        mutation.addedNodes.forEach(function (node) {
          if (node.nodeType === Node.TEXT_NODE) {
            node.nodeValue = translateString(node.nodeValue, currentLang);
          } else if (node.nodeType === Node.ELEMENT_NODE) {
            translateTextNodes(node, currentLang);
          }
        });

        if (mutation.type === "characterData" && mutation.target && mutation.target.nodeType === Node.TEXT_NODE) {
          mutation.target.nodeValue = translateString(mutation.target.nodeValue, currentLang);
        }
      });
      translateAttributes(currentLang);
      normalizeWorkingHourTexts();
    }, 0);
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true,
    characterData: true
  });

  // Initial state
  applyLanguage(currentLang);
})();
