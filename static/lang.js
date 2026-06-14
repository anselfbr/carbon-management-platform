/* Carbon Management Platform - Language Toggle
   Non-invasive implementation:
   - Does not modify existing platform functions
   - Does not overwrite onclick handlers
   - Only translates visible text nodes and placeholders
*/
(function () {
  const toggle = document.getElementById('langToggle');
  if (!toggle) return;

  let currentLang = 'en';

  const pairs = [
    ['Carbon Management Platform', '碳管理平台'],
    ['Manufacturing Data Layer', '製造資料層'],
    ['Modules', '模組'],
    ['Product Data Preparation', '產品資料準備'],
    ['Rule Management', '規則管理'],
    ['BOM Expansion', 'BOM 展開'],
    ['Carbon Emission Factor Selection', '碳排放係數選擇'],
    ['PCF Calculation', '產品碳足跡計算'],
    ['Production Environment', '正式環境'],
    ['Online | Version 1.0', '線上｜版本 1.0'],
    ['Rule Master Enabled', 'Rule Master 已啟用'],
    ['Multi-file Upload', '多檔上傳'],
    ['Version 1.0', '版本 1.0'],

    ['Prepare manufacturing data, expand BOM structures, select emission factors, and calculate product carbon footprints.', '準備製造資料、展開 BOM 結構、選擇碳排放係數，並計算產品碳足跡。'],
    ['Work order & batch preparation', '工單與批次資料準備'],
    ['Multi-level BOM explosion', '多階 BOM 展開'],
    ['Factor mapping and selection', '係數對應與選擇'],
    ['Product carbon footprint', '產品碳足跡'],

    ['Step 1 · Work Order Processing', '步驟 1 · 工單處理'],
    ['Upload one or multiple SAP production work order files.', '上傳一份或多份 SAP 生產工單檔案。'],
    ['Upload SAP Production Work Orders', '上傳 SAP 生產工單'],
    ['Reporting Year', '報告年度'],
    ['Run Consolidation & Classification', '執行合併與分類'],
    ['Ready for processing', '準備處理'],
    ['Upload SAP work orders and start the classification workflow.', '上傳 SAP 工單並開始分類流程。'],
    ['Idle', '待命'],
    ['Download Step 1 Output Excel', '下載 Step 1 輸出 Excel'],

    ['Step 2 · Batch Data Formatting', '步驟 2 · 批次資料格式化'],
    ['Convert Step 1 output and batch file template into a standardized batch file.', '將 Step 1 輸出與批次範本轉換為標準化批次檔。'],
    ['Step 1 Output', 'Step 1 輸出'],
    ['Annual output & classification result', '年度產量與分類結果'],
    ['Batch Template', '批次範本'],
    ['Required batch file format', '必要批次檔格式'],
    ['Formatted Batch File', '已格式化批次檔'],
    ['Ready for downstream processing', '可供後續處理'],
    ['Upload Step 1 Output File', '上傳 Step 1 輸出檔'],
    ['Upload Bulk Template File', '上傳 Bulk 範本檔'],
    ['Generate Formatted Bulk File', '產生格式化 Bulk 檔'],
    ['Ready for batch formatting', '準備批次格式化'],
    ['Upload Step 1 output and bulk template to generate formatted bulk file.', '上傳 Step 1 輸出與 Bulk 範本以產生格式化檔案。'],
    ['Download Formatted Bulk File', '下載格式化 Bulk 檔'],
    ['Step 2 extracts required fields from the Step 1 output and writes them into the bulk template.', 'Step 2 會從 Step 1 輸出擷取必要欄位並寫入 Bulk 範本。'],
    ['Activity Data and Products sheets will be populated automatically.', '系統會自動填入 Activity Data 與 Products 分頁。'],

    ['Execution Log', '執行紀錄'],
    ['Summary', '摘要'],
    ['Rules', '規則'],
    ['Ready. Upload SAP production work orders and start processing.', '準備完成。請上傳 SAP 生產工單並開始處理。'],
    ['Version 1.0 Decision Flow', '版本 1.0 判斷流程'],
    ['Product Series Engine', '產品系列引擎'],
    ['Rule Master', '規則主檔'],
    ['Product Series Master', '產品系列主檔'],
    ['Default WIP', '預設 WIP'],
    ['Required SAP Fields', '必要 SAP 欄位'],

    ['Upload Standard BOM', '上傳標準 BOM'],
    ['Upload Raw Material Bulk Template', '上傳原物料 Bulk 範本'],
    ['Process BOM Expansion', '執行 BOM 展開'],
    ['Ready for BOM Expansion', '準備 BOM 展開'],
    ['Upload standard BOM and raw material bulk template to start processing.', '上傳標準 BOM 與原物料 Bulk 範本後開始處理。'],
    ['Download Raw Material Bulk', '下載原物料 Bulk'],
    ['BOM Expansion Logic', 'BOM 展開邏輯'],
    ['Configure source column names for BOM Expansion.', '設定 BOM 展開來源欄位名稱。'],
    ['Use ', '使用'],
    [' for system settings or ', '套用系統設定，或使用'],
    [' to apply your input.', '套用輸入內容。'],
    ['Parent Material', '母件料號'],
    ['Component Description', '元件描述'],
    ['Component', '元件料號'],
    ['Quantity', '數量'],
    ['Unit', '單位'],
    ['Material Group', '物料群組'],
    ['Valid From', '有效起始日'],
    ['Default', '預設'],
    ['Confirm', '確認'],
    ['Current Setting:', '目前設定：'],
    ['Semi-finished Rule', '半成品判斷規則'],
    ['Component also exists as Parent Node', 'Component 同時存在於 Parent Node'],
    ['Quantity Roll-up', '數量展開邏輯'],
    ['Multiply quantities across all BOM levels', '跨 BOM 階層累乘數量'],
    ['Raw Material Bulk Output', '原物料 Bulk 輸出'],
    ['Input Sheet Activity Data and Input Sheet Raw Material will be populated.', '系統會填入 Input Sheet Activity Data 與 Input Sheet Raw Material。'],
    ['Optional fields are not written.', 'Optional 欄位不會寫入。'],
    ['BOM Expansion is ready.', 'BOM 展開已準備完成。'],

    ['Import and download classification rule masters for Product Data Preparation.', '匯入與下載產品資料準備使用的分類規則主檔。'],
    ['Rule Management is a sub-function under Product Data Preparation.', '規則管理是產品資料準備下的子功能。'],
    ['It controls product type classification, product series mapping, WIP judgment and customer mapping logic.', '用於控制產品類型分類、產品系列對應、WIP 判斷與客戶對應邏輯。'],
    ['Upload Rule Master', '上傳 Rule Master'],
    ['Import Rule Master', '匯入 Rule Master'],
    ['Download Rule Master', '下載 Rule Master'],
    ['Download Product Series Master', '下載 Product Series Master'],
    ['Rule Master Fields', 'Rule Master 欄位'],
    ['Default Rules', '預設規則'],
    ['Rule Management is ready.', '規則管理已準備完成。'],

    ['1. Product Data Preparation', '1. 產品資料準備'],
    ['Complete Step 1 Work Order Processing, Step 2 Batch Data Formatting and Rule Management.', '完成 Step 1 工單處理、Step 2 批次資料格式化與規則管理。'],
    ['2. BOM Expansion', '2. BOM 展開'],
    ['Reserved module. This area can be extended for multi-level BOM explosion.', '預留模組。此區可延伸為多階 BOM 展開。'],
    ['3. Carbon Emission Factor Selection', '3. 碳排放係數選擇'],
    ['Reserved module. This area can be extended for emission factor mapping and selection.', '預留模組。此區可延伸為排放係數對應與選擇。'],
    ['4. PCF Calculation', '4. 產品碳足跡計算'],
    ['Reserved module. This area can be extended for product carbon footprint calculation.', '預留模組。此區可延伸為產品碳足跡計算。']
  ];

  function sortedPairs(fromIndex) {
    return pairs.slice().sort((a, b) => b[fromIndex].length - a[fromIndex].length);
  }

  function translateTextNodes(targetLang) {
    const fromIndex = targetLang === 'zh' ? 0 : 1;
    const toIndex = targetLang === 'zh' ? 1 : 0;
    const activePairs = sortedPairs(fromIndex);

    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
      acceptNode(node) {
        const parent = node.parentElement;
        if (!parent) return NodeFilter.FILTER_REJECT;
        if (parent.closest('script, style')) return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });

    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);

    nodes.forEach(node => {
      let value = node.nodeValue;
      activePairs.forEach(pair => {
        value = value.split(pair[fromIndex]).join(pair[toIndex]);
      });
      node.nodeValue = value;
    });
  }

  function translatePlaceholders(targetLang) {
    document.querySelectorAll('[placeholder]').forEach(el => {
      const value = el.getAttribute('placeholder');
      if (targetLang === 'zh') {
        if (value === 'e.g. 2024；blank = all years') el.setAttribute('placeholder', '例如 2024；空白 = 全部年度');
        if (value === 'Enter BOM column name') el.setAttribute('placeholder', '輸入 BOM 欄位名稱');
      } else {
        if (value === '例如 2024；空白 = 全部年度') el.setAttribute('placeholder', 'e.g. 2024；blank = all years');
        if (value === '輸入 BOM 欄位名稱') el.setAttribute('placeholder', 'Enter BOM column name');
      }
    });
  }

  function applyLanguage(targetLang) {
    currentLang = targetLang;
    translateTextNodes(targetLang);
    translatePlaceholders(targetLang);
    document.documentElement.lang = targetLang === 'zh' ? 'zh-Hant' : 'en';
    toggle.textContent = targetLang === 'zh' ? 'EN' : '中';
  }

  toggle.addEventListener('click', function (event) {
    event.preventDefault();
    event.stopPropagation();
    applyLanguage(currentLang === 'en' ? 'zh' : 'en');
  });

  // Existing navigation functions rewrite some text in English.
  // After those original events complete, re-apply Chinese if needed.
  document.addEventListener('click', function (event) {
    if (currentLang !== 'zh') return;
    const shouldRefresh = event.target.closest('.nav-item, .sub-nav-item, [data-dashboard-target], #homeButton, #homeLogo');
    if (shouldRefresh) {
      setTimeout(() => applyLanguage('zh'), 40);
    }
  });
})();
