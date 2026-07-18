/* DIP i18n MODULE3 unified progress + M4 collapse 20260718 v15
   Key-based + dynamic text translation.
   This file does not overwrite existing platform click handlers.
*/
(function () {
  const button = document.getElementById("langToggle");
  if (!button) return;

  const STORAGE_KEY = "dip_lang";
  let currentLang = localStorage.getItem(STORAGE_KEY) || localStorage.getItem("cmp_lang") || "en";
  let isApplying = false;

  if (localStorage.getItem(STORAGE_KEY)) {
    localStorage.setItem("cmp_lang", localStorage.getItem(STORAGE_KEY));
  }

  const keyed = {
    "automatedDataConsolidation": {"en":"Automated Data Consolidation","zh":"全自動化資料彙整"},
    "oneClickPipelineDesc": {"en":"One-click data consolidation and Bulk generation","zh":"一鍵完成資料彙整與 Bulk 產出"},
    "automationPageLead": {"en":"Upload the required sources and run automatically through M3. After suppliers are created in the data center, upload the official Raw Material batch template and continue M3A.","zh":"上傳完整資料來源後，系統將自動執行至 M3；完成供應商建立後，上傳資料中心下載的正式原物料批次檔，再繼續執行 M3A。"},
    "oneClickProcess": {"en":"One-click Generation","zh":"一鍵產出"},
    "automationProductionFiles": {"en":"Production Quantity Work Orders","zh":"生產數量工單"},
    "automationLaborFiles": {"en":"Working Hour Work Orders","zh":"工時工單"},
    "automationBomFiles": {"en":"Standard BOM Files","zh":"標準 BOM"},
    "automationProductTemplate": {"en":"Product Activity Bulk Template","zh":"Product Activity Bulk Template"},
    "automationRawTemplate": {"en":"Raw Material Bulk Template","zh":"Raw Material Bulk Template"},
    "automationSupplierFiles": {"en":"Supplier Files","zh":"供應商資料"},
    "automationCclFile": {"en":"CCL Mapping File","zh":"CCL 係數組配表"},
    "automationUploadSectionTitle": {"en":"Uploads and execution settings","zh":"上傳資料與執行設定"},
    "automationUploadSectionHint": {"en":"Use the control at the upper right to collapse or expand the upload fields.","zh":"可使用右上角按鈕收合或展開上傳欄位。"},
    "collapseUploadFields": {"en":"Collapse","zh":"收合"},
    "expandUploadFields": {"en":"Expand","zh":"展開"},
    "startAutomatedConsolidation": {"en":"Start Automated Data Consolidation","zh":"開始全自動化資料彙整"},
    "automationReady": {"en":"Automated workflow ready","zh":"全自動化流程已就緒"},
    "automationReadyHint": {"en":"Select all required files to begin.","zh":"選擇所有必要檔案後即可開始。"},
    "automationProgressEyebrow": {"en":"STATUS","zh":"執行狀態"},
    "automationProgressTitle": {"en":"Execution Progress","zh":"執行進度"},
    "automationProgressHint": {"en":"Each completed stage is checked automatically before the next stage begins.","zh":"每完成一個階段，系統會自動標示勾選並繼續下一階段。"},
    "automationExecutionLog": {"en":"Execution Log","zh":"執行紀錄"},
    "automationRunningHint": {"en":"The automated workflow is running. Do not close this page.","zh":"全自動化流程執行中，請勿關閉此頁面。"},
    "automationDataReading": {"en":"Reading data","zh":"資料讀取"},
    "automationDataWriting": {"en":"Writing data","zh":"資料寫入"},
    "automationDataMapping": {"en":"Mapping data","zh":"資料對應"},
    "automationDataProcessing": {"en":"Processing data","zh":"資料處理"},
    "automationRowsUnit": {"en":"rows","zh":"筆"},
    "automationCompleted": {"en":"Automated consolidation completed","zh":"全自動化資料彙整完成"},
    "automationCompletedHint": {"en":"The Product Activity Bulk, Supplier Bulk, and CCL Raw Material Bulk files are ready for download.","zh":"Product Activity Bulk、供應商 Bulk 與 CCL Raw Material Bulk 檔已可下載。"},
    "automationFailed": {"en":"Automated consolidation failed","zh":"全自動化資料彙整失敗"},
    "downloadM1bResult": {"en":"Download Product Activity Bulk","zh":"下載 Product Activity Bulk"},
    "downloadSupplierBulkResult": {"en":"Download Supplier Bulk","zh":"下載供應商 Bulk"},
    "automationSupplierBulkMissing": {"en":"Supplier Bulk was not generated. Confirm that the built-in Supplier Bulk Template exists.","zh":"供應商 Bulk 未成功產出，請確認內建 Supplier Bulk Template 是否存在。"},
    "automationSupplierBulkReadyLog": {"en":"Supplier Bulk: generated.","zh":"供應商 Bulk：已產出。"},
    "downloadM3aResult": {"en":"Download CCL Raw Material Bulk","zh":"下載 CCL Raw Material Bulk"},
    "automationLaborBoth": {"en":"Labor HR.Act + FOH-Others.Act","zh":"人員＋設備工時"},
    "automationLaborHrOnly": {"en":"Labor HR.Act Only","zh":"僅人員工時"},
    "automationFohOnly": {"en":"FOH-Others.Act Only","zh":"僅設備工時"},
    "automationStepM1a": {"en":"Work Order Consolidation","zh":"工單資料彙整"},
    "automationStepM2a": {"en":"Standard BOM Total Usage","zh":"標準 BOM 總用量"},
    "automationStepM1b": {"en":"Product Activity Bulk","zh":"產品活動數據 Bulk"},
    "automationStepM2b": {"en":"Raw Material Bulk","zh":"原物料 Bulk 寫入"},
    "automationStepM2c": {"en":"Supplier Data Mapping","zh":"供應商資料對應"},
    "automationStepM3": {"en":"CCL Factor Mapping","zh":"CCL 係數對應"},
    "automationStepM3a": {"en":"Official Raw Material Template Application & Splitting","zh":"正式原物料 Template 套版與切檔"},
    "automationStageWaiting": {"en":"Waiting","zh":"待執行"},
    "automationStageRunning": {"en":"In progress","zh":"執行中"},
    "automationStageCompleted": {"en":"Completed","zh":"已完成"},
    "automationStageFailed": {"en":"Failed","zh":"失敗"},
    "automationGeneratingProductBulk": {"en":"Generating the Product Activity Bulk file.","zh":"正在產出產品活動數據 Bulk。"},
    "automationPipelineCompletedLog": {"en":"Automated data consolidation: all stages completed.","zh":"全自動化資料彙整：全部完成。"},
    "logoTitle": {
        "en": "Data Integration Platform (DIP)",
        "zh": "資料整合平台"
    },
    "logoSubtitle": {
        "en": "Manufacturing Data Layer",
        "zh": "製造資料層"
    },
    "eyebrowCarbonManagementPlatform": {
        "en": "Data Integration Platform (DIP)",
        "zh": "資料整合平台"
    },
    "modules": {
        "en": "Modules",
        "zh": "模組"
    },
    "module": {
        "en": "Module",
        "zh": "模組"
    },
    "productDataPreparation": {
        "en": "Product Data Preparation",
        "zh": "產品資料準備"
    },
    "ruleManagement": {
        "en": "Rule Management",
        "zh": "規則管理"
    },
    "moduleEntry": {
        "en": "Module Menu",
        "zh": "入口選單"
    },
    "module1Title": {
        "en": "Module 1 · Product Data Preparation",
        "zh": "模組 1 · 產品資料準備"
    },
    "module1Stage1Placeholder": {
        "en": "Module 1 Entry. Select Module 1A Work Order Processing, Module 1B Batch Data Formatting, or Rule Management to enter the corresponding workspace.",
        "zh": "模組1入口頁面。請選擇 Module 1A 工單處理、Module 1B 批次資料格式化或規則管理，進入各自專區。"
    },
    "module1PageLead": {
        "en": "Select Module 1A, Module 1B, or Rule Management to enter the corresponding workspace.",
        "zh": "請選擇 Module 1A、Module 1B 或規則管理，快速進入對應功能區。"
    },
    "module1aShortTitle": {
        "en": "Module 1A",
        "zh": "Module 1A"
    },
    "module1bShortTitle": {
        "en": "Module 1B",
        "zh": "Module 1B"
    },
    "module1aTitle": {
        "en": "Module 1A · Work Order Processing",
        "zh": "Module 1A · 工單處理"
    },
    "module1bTitle": {
        "en": "Module 1B · Batch Data Formatting",
        "zh": "Module 1B · 批次資料格式化"
    },
    "module1aStage1Desc": {
        "en": "Consolidate SAP production and working-hour work orders, then generate annual product output and classification results.",
        "zh": "整合 SAP 生產工單與工時工單，產出年度產品產量與分類結果。"
    },
    "module1bStage1Desc": {
        "en": "Use Module 1A annual output and classification results to fill the Product Activity Bulk Template.",
        "zh": "引用 Module 1A 年度產品產量與分類結果，寫入 Product Activity Bulk Template。"
    },
    "module1RuleStage1Desc": {
        "en": "Maintain product classification, product series, and Rule Master settings used by Module 1A / 1B.",
        "zh": "維護產品分類、產品系列與 Rule Master，供 Module 1A / 1B 使用。"
    },
    "module1IntegrationPolicy": {
        "en": "Module 1 Integration Policy",
        "zh": "模組 1 整合原則"
    },
    "module1IntegrationPolicyDesc": {
        "en": "Module 1A handles work order consolidation, annual output roll-up, and product classification. Module 1B converts Module 1A output into the third-party Bulk Template format.",
        "zh": "Module 1A 負責工單合併、年度產量彙總與產品分類；Module 1B 負責將 Module 1A 輸出轉換為第三方 Bulk Template 格式。"
    },
    "backToModule2Menu": {
        "en": "Back to Module 2 Menu",
        "zh": "返回模組 2 選單"
    },
    "module2Title": {
        "en": "Module 2 · BOM Expansion",
        "zh": "模組 2 · BOM 展開"
    },
    "module2Stage1Placeholder": {
        "en": "Module 2 Entry. Select Module 2A, 2B, or 2C to quickly enter the corresponding function area.",
        "zh": "模組2入口頁面。請選擇 Module 2A、2B 或 2C，快速進入對應功能區。"
    },
    "module2PageLead": {
        "en": "Select Module 2A, Module 2B, or Module 2C to enter the corresponding workspace.",
        "zh": "請選擇 Module 2A、Module 2B 或 Module 2C，快速進入對應功能區。"
    },
    "module2aShortTitle": {
        "en": "Module 2A",
        "zh": "Module 2A"
    },
    "module2bShortTitle": {
        "en": "Module 2B",
        "zh": "Module 2B"
    },
    "module2cShortTitle": {
        "en": "Module 2C",
        "zh": "Module 2C"
    },
    "module2aTitle": {
        "en": "Module 2A · Standard BOM Total Usage",
        "zh": "Module 2A · 標準BOM表總用量"
    },
    "module2bTitle": {
        "en": "Module 2B · Raw Material Bulk Writing",
        "zh": "Module 2B · Raw Material Bulk 寫入"
    },
    "module2cTitle": {
        "en": "Module 2C · Supplier Mapping Bulk",
        "zh": "Module 2C · Supplier Mapping Bulk"
    },
    "module2aStage1Desc": {
        "en": "Expand standard BOM, aggregate total raw material usage, and generate working_hour_rollup.",
        "zh": "展開標準 BOM、彙總原物料總用量，並產出 working_hour_rollup。"
    },
    "module2bStage1Desc": {
        "en": "Use Module 2A standard BOM total usage to fill the Raw Material Bulk Template.",
        "zh": "引用 Module 2A 標準BOM表總用量，寫入 Raw Material Bulk Template。"
    },
    "module2cStage1Desc": {
        "en": "Use Module 2B Raw Material Bulk ZIP, apply supplier data, and output the Module 2C Bulk.",
        "zh": "引用 Module 2B Raw Material Bulk ZIP，套用供應商資料並產出 Module 2C Bulk。"
    },
    "module2ProgressCheckingTitle": {
        "en": "Checking MODULE 2 progress",
        "zh": "MODULE 2 進度檢查中"
    },
    "module2ProgressCheckingText": {
        "en": "Checking the latest outputs from Module 2A, Module 2B, and Module 2C.",
        "zh": "系統正在確認 Module 2A、Module 2B 與 Module 2C 的最新產出狀態。"
    },
    "module2aOutputsTitle": {
        "en": "Module 2A Standard BOM Total Usage and Working-hour Rollup",
        "zh": "Module 2A 標準 BOM 總用量與工時彙整"
    },
    "module2bOutputTitle": {
        "en": "Module 2B Raw Material Bulk",
        "zh": "Module 2B Raw Material Bulk"
    },
    "module2cOutputsTitle": {
        "en": "Module 2C Supplier Mapping Outputs",
        "zh": "Module 2C 供應商對應結果"
    },
    "sourceChecking": {
        "en": "Checking source...",
        "zh": "來源檢查中..."
    },
    "module2IntegrationPolicy": {
        "en": "Module 2 Integration Policy",
        "zh": "模組 2 整合原則"
    },
    "module2IntegrationPolicyDesc": {
        "en": "Module 2A first generates standard BOM total usage and working_hour_rollup. Module 2B writes Raw Material Bulk. Module 2C then adds supplier data for Module 3.",
        "zh": "Module 2A 先產出標準 BOM 總用量與 working_hour_rollup；Module 2B 寫入 Raw Material Bulk；Module 2C 再加入供應商資料，供 Module 3 引用。"
    },
    "bomExpansion": {
        "en": "BOM Expansion",
        "zh": "BOM 展開"
    },
    "carbonEmissionFactorSelection": {
        "en": "Carbon Emission Factor Selection",
        "zh": "碳排放係數選擇"
    },

    "module3Title": {
        "en": "Module 3 · Carbon Emission Factor Selection",
        "zh": "模組 3 · 碳排放係數選擇"
    },
    "module3Stage1Placeholder": {
        "en": "Module 3 Entry. Select either CCL factor import or the Ecoinvent emission factor database to enter the corresponding workspace.",
        "zh": "模組3 入口頁面。請選擇以CCL 係數帶入或查詢Ecoinvent係數資料庫，進入各自專區。"
    },
    "stage1": {
        "en": "Stage 1",
        "zh": "第一階段"
    },
    "cclMapping": {
        "en": "CCLibrary Emission Factor Database",
        "zh": "CCL係數資料庫"
    },
    "cclMappingStage1Desc": {
        "en": "Section A provides the CCLibrary emission factor database and imports CCL factors by raw material number.",
        "zh": "A專區為CCLibrary係數資料庫，並以原物料料號帶入CCL係數。"
    },
    "factorLibrary": {
        "en": "Ecoinvent Emission Factor Database",
        "zh": "Ecoinvent係數資料庫"
    },
    "factorLibraryStage1Desc": {
        "en": "Section B provides the Ecoinvent emission factor database for keyword-based factor search.",
        "zh": "B專區為Ecoinvent係數資料庫，提供係數查詢功能。"
    },
    "enterWorkspace": {
        "en": "Enter Workspace",
        "zh": "進入專區"
    },
    "module3CclZoneTitle": {
        "en": "A. CCLibrary Emission Factor Database",
        "zh": "A. CCL係數資料庫"
    },
    "module3CclZoneLead": {
        "en": "Use Module 2C Raw Material Bulk and the Lite-On CCL mapping table to import CCL Items and emission factors.",
        "zh": "直接使用 Module 2C Raw Material Bulk，並依光寶 CCL 係數組配表帶入 CCL Item 與碳係數。"
    },
    "backToModule3Menu": {
        "en": "Back to Module 3 Menu",
        "zh": "返回模組 3 選單"
    },
    "backToModule1Menu": {
        "en": "Back to Module 1 Menu",
        "zh": "返回模組 1 選單"
    },
    "module3CclStep1": {
        "en": "Upload raw material bulk output from Module 2",
        "zh": "上傳 MODULE2 產出的 raw material bulk 檔"
    },
    "module3CclStep2": {
        "en": "Upload Lite-On CCL mapping table",
        "zh": "上傳光寶 CCL 係數組配表"
    },
    "module3CclStep3": {
        "en": "Map CCL Item, emission factor, and unit by Material",
        "zh": "依 Material 對應 CCL Item、碳係數與單位"
    },
    "module3LibraryZoneTitle": {
        "en": "B. Ecoinvent Emission Factor Database",
        "zh": "B. Ecoinvent係數資料庫"
    },
    "module3LibraryZoneLead": {
        "en": "Search the Ecoinvent factor database by keyword and review factor details.",
        "zh": "以關鍵字查詢 Ecoinvent 係數資料庫，並檢視係數詳細資訊。"
    },
    "keywordSearch": {
        "en": "Keyword Search",
        "zh": "關鍵字搜尋"
    },
    "factorSearchPlaceholder": {
        "en": "Enter keyword, e.g. solder",
        "zh": "請輸入關鍵字，例如 solder"
    },
    "module3LibraryNotConnected": {
        "en": "Search API is not enabled yet. The next stage will connect the APOS / Cut-off factor databases.",
        "zh": "搜尋 API 尚未啟用，下一階段接入 APOS / Cut-off 係數資料庫。"
    },
    "module3UploadRawBulk": {"en": "Upload raw material bulk file from Module 2", "zh": "上傳 MODULE2 產出的 raw material bulk 檔"},
    "module3UploadCclMapping": {"en": "Upload Lite-On CCL mapping table", "zh": "上傳光寶 CCL 係數組配表"},
    "module3RunCclMapping": {"en": "Run CCL Mapping", "zh": "執行 CCL 係數對應"},
    "module3CombinedReadyTitle": {"en": "Ready to start", "zh": "等待開始執行"},
    "module3CombinedRun": {"en": "Start", "zh": "開始執行"},
    "module3CclReadyTitle": {"en": "Ready for upload", "zh": "待上傳檔案"},
    "module3CclReadyText": {"en": "The system completes CCL mapping first, then automatically applies the official Raw Material batch template.", "zh": "系統會先完成 CCL 係數對應，再自動套用正式原物料批次檔。"},
    "module3DownloadFilledBulk": {"en": "Download M3 factor-filled intermediate", "zh": "下載 M3 係數對應完成檔"},
    "module3aSourceTitle": {"en": "Use M3 factor-filled output", "zh": "引用 M3 係數對應完成檔"},
    "module3aOfficialTemplateLabel": {"en": "Official Raw Material Batch Template", "zh": "正式原物料批次檔 Template"},
    "module3aOfficialTemplateHint": {"en": "After creating suppliers in the data center, download and upload the official Raw Material batch file.", "zh": "於資料中心建立供應商後，下載正式原物料批次檔並上傳"},
    "module3aRunOfficialTemplate": {"en": "Run M3A Official Template", "zh": "執行 M3A 正式套版"},
    "module3aReadyTitle": {"en": "Waiting to start", "zh": "等待開始執行"},
    "module3aReadyText": {"en": "M3A will start automatically after M3 is completed.", "zh": "M3 完成後將自動接續 M3A 正式套版。"},
    "module3aDownloadFinalBulk": {"en": "Download final Raw Material Bulk", "zh": "下載最終原物料 Bulk"},
    "automationContinueM3a": {"en": "Continue M3A", "zh": "繼續執行 M3A"},
    "module3ProcessingTitle": {"en": "Processing", "zh": "處理中"},
    "module3ProcessingText": {"en": "Writing CCL factor fields. Please wait.", "zh": "正在寫入 CCL 係數欄位，請稍候。"},
    "module3CclSuccessTitle": {"en": "CCL Mapping Completed", "zh": "CCL 係數對應完成"},
    "module3CclSuccessText": {"en": "Completed: {written} rows written, {unmatched} rows unmatched.", "zh": "完成：已寫入 {written} 筆，未對應 {unmatched} 筆。"},
    "module3CclFailed": {"en": "CCL Mapping failed", "zh": "CCL 係數對應失敗"},
    "module3LibraryReady": {"en": "Type keyword search or name search terms to search the Ecoinvent factor database automatically.", "zh": "請輸入關鍵字查詢或名稱查詢內容後將自動搜尋 Ecoinvent 係數資料庫。"},
    "module3SearchFactor": {"en": "Search", "zh": "搜尋"},
    "allItems": {"en": "All", "zh": "全部"},
    "processType": {"en": "Process Type", "zh": "製程類型"},
    "productionOnly": {"en": "Production only", "zh": "僅生產"},
    "productionWithTransport": {"en": "Production incl. transport", "zh": "生產(含運輸)"},
    "clearFilters": {"en": "Clear", "zh": "清除"},
    "factorSource": {"en": "Source", "zh": "來源"},
    "activityName": {"en": "Activity Name", "zh": "Activity Name"},
    "activityKeywordSearch": {"en": "Keyword search", "zh": "關鍵字查詢"},
    "geography": {"en": "Geography", "zh": "Geography"},
    "emissionFactorValue": {"en": "Emission Factor", "zh": "係數值"},
    "emissionFactorUnit": {"en": "Emission Factor Unit", "zh": "係數單位"},
    "referenceProductName": {"en": "Reference Product Name", "zh": "Reference Product Name"},
    "referenceNameSearch": {"en": "Name search", "zh": "名稱查詢"},
    "activityNameSearchPlaceholder": {"en": "Type keywords, e.g. wafer waste", "zh": "請輸入關鍵字，例如 wafer waste"},
    "referenceProductSearchPlaceholder": {"en": "Type name keywords, e.g. solder paste", "zh": "請輸入名稱，例如 solder paste"},
    "lciaIndicator": {"en": "LCIA Indicator", "zh": "LCIA 指標"},
    "clickForFactorDetail": {"en": "Click to view factor details", "zh": "點選查看係數詳細說明"},
    "unit": {"en": "Unit", "zh": "單位"},
    "rowsPerPage": {"en": "Rows per page", "zh": "每頁顯示"},
    "rowsUnit": {"en": "rows", "zh": "筆"},
    "previousPage": {"en": "Previous", "zh": "上一頁"},
    "nextPage": {"en": "Next", "zh": "下一頁"},
    "ipcc2021Gwp100": {"en": "IPCC 2021 GWP100", "zh": "IPCC 2021 GWP100"},
    "module3KeywordTooShort": {"en": "Please enter at least 2 characters in either Activity Name or Reference Product Name.", "zh": "請至少在其中一個欄位輸入 2 個字元。"},
    "copy": {"en": "Copy", "zh": "複製"},
    "copied": {"en": "Copied", "zh": "已複製"},
    "copyFailed": {"en": "Copy failed", "zh": "複製失敗"},
    "module3Searching": {"en": "Searching, please wait...", "zh": "查詢中，請稍候..."},
    "module3SearchFailed": {"en": "Factor search failed", "zh": "係數搜尋失敗"},
    "module3SearchSuccessText": {"en": "Found {count} records. Results are ordered APOS first, then Cut-off.", "zh": "找到 {count} 筆資料；顯示順序為 APOS 優先，再 Cut-off。"},
    "module3SearchCompletedText": {"en": "Search completed. Results below: {count} records.", "zh": "已完成查詢，結果如下：共 {count} 筆資料。"},
    "module3NoFactorResult": {"en": "No matching records found.", "zh": "查無符合資料。"},
    "module3IntegrationPolicy": {
        "en": "Module 3 Integration Policy",
        "zh": "模組 3 整合原則"
    },
    "module3IntegrationPolicyDesc": {
        "en": "The CCLibrary workspace maps raw material numbers to CCL Items and automatically imports emission factors. The Ecoinvent workspace is primarily designed for factor searching and will support recommended factor selection in future releases.",
        "zh": "CCL係數資料庫專區以原物料料號對應，帶入CCL Item和碳係數；Ecoinvent係數資料庫專區主要為查詢功能，未來將導入建議係數功能。"
    },
    "module3PageLead": {
        "en": "Section A provides the CCLibrary emission factor database, while Section B provides the Ecoinvent emission factor database.",
        "zh": "A專區為CCLibrary係數資料庫，B專區為Ecoinvent係數資料庫。"
    },
    "module3NoticeTitle": {
        "en": "3. Carbon Emission Factor Selection",
        "zh": "3. 碳排放係數選擇"
    },
    "module3NoticeText": {
        "en": "Section A provides the CCLibrary emission factor database, while Section B provides the Ecoinvent emission factor database.",
        "zh": "A專區為CCLibrary係數資料庫，B專區為Ecoinvent係數資料庫。"
    },
    "dashboardLead": {"en": "Integrate manufacturing data, expand BOM structures, select emission factors, and support product carbon footprint workflows.", "zh": "整合製造資料、展開 BOM 結構、選擇碳排放係數，並支援產品碳足跡流程。"},
    "productDataLead": {"en": "Prepare production output and batch data for product carbon footprint workflows.", "zh": "準備產品碳足跡流程所需的生產產出與批次資料。"},
    "ruleManagementLead": {"en": "Maintain Product Data Preparation Rules, including Rule Master and Product Series Master.", "zh": "維護產品資料準備規則，包含 Rule Master 與 Product Series Master。"},
    "modulePrepNoticeText": {"en": "Complete Module 1A Work Order Processing, Module 1B Batch Data Formatting and Rule Management.", "zh": "完成 Module 1A 工單處理、Module 1B 批次資料格式化與規則管理。"},
    "modulePcfNoticeText": {"en": "Reserved module. This area can be extended for product carbon footprint calculation.", "zh": "預留模組。此區可延伸為產品碳足跡計算。"},
    "bomExpansionPageLead": {"en": "Expand multi-level BOM structures, aggregate total raw material demand for finished products, and generate Raw Material Bulk files.", "zh": "展開多階 BOM 結構、彙總成品需求原物料總數量，並產生原物料 Bulk 檔。"},
    "pcfCalculation": {
        "en": "PCF Calculation",
        "zh": "產品碳足跡計算"
    },
    "productionEnvironment": {
        "en": "Production Environment",
        "zh": "正式環境"
    },
    "onlineVersion": {
        "en": "Online | Version 1.0",
        "zh": "線上｜版本 1.0"
    },
    "ruleMasterEnabled": {
        "en": "Rule Master Enabled",
        "zh": "Rule Master 已啟用"
    },
    "multiFileUpload": {
        "en": "Multi-file Upload",
        "zh": "多檔上傳"
    },
    "version10": {
        "en": "Version 1.0",
        "zh": "版本 1.0"
    },
    "uploadSupplierFiles": {
        "en": "Upload Supplier Files",
        "zh": "上傳供應商檔案"
    },
    "supplierFilesHint": {
        "en": "Optional. Supports multiple Excel files. Uses Raw Material Code + Vendor number to write Supplier Name. Transportation Origin uses Supplier Address or Country + City + Street. Also exports supplier_bulk_create from the built-in template.",
        "zh": "選填，可多檔上傳。系統會依原物料代碼與 Vendor number 寫入 Supplier Name；Transportation Origin 使用 Supplier Address，若無則使用 Country + City + Street，並由內建範本輸出 supplier_bulk_create。"
    },
    "workOrderBatchPreparation": {
        "en": "Work order & batch preparation",
        "zh": "工單與批次資料準備"
    },
    "multiLevelBomExplosion": {
        "en": "Multi-level BOM explosion",
        "zh": "多階 BOM 展開"
    },
    "factorMappingSelection": {
        "en": "Factor mapping and selection",
        "zh": "係數對應與選擇"
    },
    "productCarbonFootprint": {
        "en": "Product carbon footprint",
        "zh": "產品碳足跡"
    },
    "businessUnitRuleSet": {
        "en": "Business Unit Rule Set",
        "zh": "BU 規則組"
    },
    "productionQuantityWorkOrders": {
        "en": "Production Quantity Work Orders",
        "zh": "生產數量工單"
    },
    "uploadSapWorkingHourOrders": {
        "en": "Upload SAP Working Hour Orders",
        "zh": "生產工時工單"
    },
    "laborAllocationSource": {
        "en": "Labor Allocation Source",
        "zh": "工時來源"
    },
    "reportingYear": {
        "en": "Reporting Year",
        "zh": "報告年度"
    },
    "yearPlaceholder": {
        "en": "e.g. 2024; blank = all years",
        "zh": "例如：2024；空白＝全部年度"
    },
    "runConsolidationClassification": {
        "en": "Run Consolidation & Classification",
        "zh": "執行合併與分類"
    },
    "downloadStep1OutputExcel": {
        "en": "Download Module 1A Output Excel",
        "zh": "下載 Module 1A 輸出 Excel"
    },
    "step1WorkOrderProcessing": {
        "en": "Work Order Processing",
        "zh": "工單處理"
    },
    "uploadSapProductionWorkOrderFiles": {
        "en": "Upload one or multiple SAP production work order files.",
        "zh": "上傳一份或多份 SAP 生產工單檔案。"
    },
    "step1": {
        "en": "Module 1A",
        "zh": "Module 1A"
    },
    "step2BatchDataFormatting": {
        "en": "Batch Data Formatting",
        "zh": "批次資料格式化"
    },
    "convertStep1OutputBulkTemplate": {
        "en": "Convert Module 1A output and batch file template into a standardized batch file.",
        "zh": "將 Module 1A 輸出與批次範本轉換為標準化批次檔。"
    },
    "step2": {
        "en": "Module 1B",
        "zh": "Module 1B"
    },
    "step1Output": {
        "en": "Module 1A Output",
        "zh": "Module 1A 輸出"
    },
    "annualOutputClassificationResult": {
        "en": "Annual output & classification result",
        "zh": "年度產量與分類結果"
    },
    "batchTemplate": {
        "en": "Batch Template",
        "zh": "批次範本"
    },
    "requiredBatchFileFormat": {
        "en": "Required batch file format",
        "zh": "必要批次檔格式"
    },
    "formattedBatchFile": {
        "en": "Formatted Batch File",
        "zh": "已格式化批次檔"
    },
    "readyForDownstreamProcessing": {
        "en": "Ready for downstream processing",
        "zh": "可供後續處理"
    },
    "uploadStep1OutputFile": {
        "en": "Upload Module 1A Output File",
        "zh": "上傳 Module 1A 輸出檔"
    },
    "uploadBulkTemplateFile": {
        "en": "Upload Bulk Template File",
        "zh": "上傳 Bulk 範本檔"
    },
    "workingHourSource": {
        "en": "Working Hour Source",
        "zh": "工時來源"
    },
    "directWorkingHour": {
        "en": "Direct Working Hour",
        "zh": "僅成品工時"
    },
    "includeSemiWorkingHour": {
        "en": "Include Semi-finished Working Hour",
        "zh": "包含半品工時"
    },
    "semiHourHint": {
        "en": "This option requires Module 2A working_hour_rollup. Please complete Module 2A first.",
        "zh": "此選項需先完成 Module 2A working_hour_rollup。請先完成 Module 2A。"
    },
    "generateFormattedBulkFile": {
        "en": "Generate Formatted Bulk File",
        "zh": "產生格式化 Bulk 檔"
    },
    "downloadFormattedBulkFile": {
        "en": "Download Formatted Bulk File",
        "zh": "下載格式化 Bulk 檔"
    },
    "downloadBulkZipFile": {
        "en": "Download Bulk ZIP File",
        "zh": "下載 Bulk ZIP 檔"
    },
    "step2Hint": {
        "en": "Module 1B extracts required fields from the Module 1A output and writes them into the bulk template. Activity Data and Products sheets will be populated automatically.",
        "zh": "Module 1B 會從 Module 1A 輸出擷取必要欄位並寫入 Bulk 範本。系統會自動填入 Activity Data 與 Products 分頁。"
    },
    "executionLog": {
        "en": "Execution Log",
        "zh": "執行紀錄"
    },
    "summary": {
        "en": "Summary",
        "zh": "摘要"
    },
    "rules": {
        "en": "Rules",
        "zh": "規則"
    },
    "version10DecisionFlow": {
        "en": "Version 1.0 Decision Flow",
        "zh": "版本 1.0 判斷流程"
    },
    "requiredSapFields": {
        "en": "Required SAP Fields",
        "zh": "必要 SAP 欄位"
    },
    "uploadStandardBom": {
        "en": "Upload Standard BOM",
        "zh": "上傳標準 BOM"
    },
    "uploadStandardBomFiles": {
        "en": "Upload Standard BOM Files",
        "zh": "上傳標準 BOM 檔案"
    },
    "uploadStep1OutputForRollup": {
        "en": "Use Module 1A Annual Product Output & Classification Result",
        "zh": "引用 Module 1A 年度產品產量與分類結果"
    },
    "downloadWorkingHourRollup": {
        "en": "Download Working Hour Roll-up",
        "zh": "下載工時 Roll-up"
    },
    "downloadSupplierBulk": {
        "en": "Download Supplier Bulk Create",
        "zh": "下載供應商建立 Bulk"
    },
    "processingBomExpansion": {
        "en": "Processing BOM Expansion...",
        "zh": "正在處理 BOM 展開..."
    },
    "uploadRawMaterialBulkTemplate": {
        "en": "Upload Raw Material Bulk Template",
        "zh": "上傳原物料 Bulk 範本"
    },
    "processBomExpansion": {
        "en": "Process BOM Expansion",
        "zh": "執行 BOM 展開"
    },
    "downloadRawMaterialBulk": {
        "en": "Download Raw Material Bulk",
        "zh": "下載原物料 Bulk"
    },
    "bomExpansionLogic": {
        "en": "BOM Expansion Logic",
        "zh": "BOM 展開邏輯"
    },
    "bomColumnMappingHint": {
        "en": "Configure source column names for BOM Expansion. Use Default for system settings or Confirm to apply your input.",
        "zh": "設定 BOM 展開來源欄位名稱。使用「預設」套用系統設定，或使用「確認」套用輸入內容。"
    },
    "parentMaterial": {
        "en": "Parent Material",
        "zh": "母件料號"
    },
    "component": {
        "en": "Component",
        "zh": "元件料號"
    },
    "quantity": {
        "en": "Quantity",
        "zh": "數量"
    },
    "unit": {
        "en": "Unit",
        "zh": "單位"
    },
    "componentDescription": {
        "en": "Component Description",
        "zh": "元件描述"
    },
    "materialGroup": {
        "en": "Material Group",
        "zh": "物料群組"
    },
    "validFrom": {
        "en": "Valid From",
        "zh": "有效起始日"
    },
    "default": {
        "en": "Default",
        "zh": "預設"
    },
    "confirm": {
        "en": "Confirm",
        "zh": "確認"
    },
    "enterBomColumnName": {
        "en": "Enter BOM column name",
        "zh": "輸入 BOM 欄位名稱"
    },
    "semiFinishedRule": {
        "en": "Semi-finished Rule",
        "zh": "半成品判斷規則"
    },
    "semiFinishedRuleDesc": {
        "en": "Component also exists as Parent Node",
        "zh": "Component 同時存在於 Parent Node"
    },
    "quantityRollUp": {
        "en": "Quantity Roll-up",
        "zh": "數量展開邏輯"
    },
    "quantityRollUpDesc": {
        "en": "Multiply quantities across all BOM levels",
        "zh": "跨 BOM 階層累乘數量"
    },
    "rawMaterialBulkOutput": {
        "en": "Raw Material Bulk Output",
        "zh": "原物料 Bulk 輸出"
    },
    "rawMaterialBulkOutputHint": {
        "en": "Input Sheet Activity Data and Input Sheet Raw Material will be populated. Optional fields are not written.",
        "zh": "系統會填入 Input Sheet Activity Data 與 Input Sheet Raw Material。Optional 欄位不會寫入。"
    },
    "ruleMaster": {
        "en": "Rule Master",
        "zh": "Rule Master"
    },
    "importDownloadClassificationRules": {
        "en": "Import and download classification rule masters for Product Data Preparation.",
        "zh": "匯入與下載產品資料準備使用的分類規則主檔。"
    },
    "ruleManagementHint": {
        "en": "Rule Management is a sub-function under Product Data Preparation. It controls product type classification, product series mapping, WIP judgment and customer mapping logic.",
        "zh": "規則管理是產品資料準備下的子功能，用於控制產品類型分類、產品系列對應、WIP 判斷與客戶對應邏輯。"
    },
    "ruleMasterReady": {
        "en": "Rule Master ready",
        "zh": "Rule Master 已就緒"
    },
    "selectBuAndImportRuleMaster": {
        "en": "Select a Business Unit Rule Set and import a Rule Master file.",
        "zh": "請選擇 BU 規則組並匯入 Rule Master 檔案。"
    },
    "currentRuleSet": {
        "en": "Current Rule Set",
        "zh": "目前規則組"
    },
    "currentRules": {
        "en": "Current Rules",
        "zh": "目前規則數"
    },
    "lastImport": {
        "en": "Last Import",
        "zh": "最近匯入"
    },
    "uploadRuleMaster": {
        "en": "Upload Rule Master",
        "zh": "上傳 Rule Master"
    },
    "importRuleMaster": {
        "en": "Import Rule Master",
        "zh": "匯入 Rule Master"
    },
    "downloadRuleMaster": {
        "en": "Download Rule Master",
        "zh": "下載 Rule Master"
    },
    "downloadProductSeriesMaster": {
        "en": "Download Product Series Master",
        "zh": "下載 Product Series Master"
    },
    "ruleMasterFields": {
        "en": "Rule Master Fields",
        "zh": "Rule Master 欄位"
    },
    "ruleTypeGuide": {
        "en": "Rule Type Guide",
        "zh": "Rule Type 說明"
    },
    "ruleTypePlantExact": {
        "en": "Plant Exact",
        "zh": "Plant Exact"
    },
    "ruleTypePlantExactDesc": {
        "en": "Plant exact match, for example: 3760 → Shijie Plant-IPS",
        "zh": "Plant 完全符合，例如：3760 → 石碣廠-IPS"
    },
    "ruleTypeMaterialNumberExact": {
        "en": "Material Number Exact",
        "zh": "Material Number Exact"
    },
    "ruleTypeMaterialNumberExactDesc": {
        "en": "Exact material number match, for example: SG-96000-00A",
        "zh": "完整料號比對，例如：SG-96000-00A"
    },
    "ruleTypeMaterialNumberPrefix": {
        "en": "Material Number Prefix",
        "zh": "Material Number Prefix"
    },
    "ruleTypeMaterialNumberPrefixDesc": {
        "en": "Material number prefix match, for example: 850-, 851-, 852-",
        "zh": "料號前綴比對，例如：850-、851-、852-"
    },
    "ruleTypeSeriesExact": {
        "en": "Series Exact",
        "zh": "Series Exact"
    },
    "ruleTypeSeriesExactDesc": {
        "en": "Exact product series match, for example: SN3103B02",
        "zh": "完整產品系列比對，例如：SN3103B02"
    },
    "ruleTypeSeriesPrefix": {
        "en": "Series Prefix",
        "zh": "Series Prefix"
    },
    "ruleTypeSeriesPrefixDesc": {
        "en": "Product series prefix match, for example: SN, SP, FU, SCMC",
        "zh": "產品系列前綴比對，例如：SN、SP、FU、SCMC"
    },
    "ruleTypeDescriptionContains": {
        "en": "Description Contains",
        "zh": "Description Contains"
    },
    "ruleTypeDescriptionContainsDesc": {
        "en": "Material description contains keywords, for example: TOUCHPAD MODULE, ASSY",
        "zh": "Material description 包含關鍵字，例如：TOUCHPAD MODULE、ASSY"
    },
    "maintenancePrinciples": {
        "en": "Maintenance Principles",
        "zh": "維護原則"
    },
    "maintenancePriority": {
        "en": "The smaller the Priority number, the higher the priority.",
        "zh": "Priority 數字越小，優先度越高。"
    },
    "maintenanceRuleMasterControlsSite": {
        "en": "Product Line and Production Site are controlled only by Rule Master.",
        "zh": "Product Line 與 Production Site 僅由 Rule Master 控制。"
    },
    "maintenanceChangeRuleMasterOnly": {
        "en": "When adding or adjusting classification rules, only modify the corresponding BU rule_master.csv; main.py does not need to be modified.",
        "zh": "新增或調整分類規則時，只需修改對應 BU 的 rule_master.csv，不需修改 main.py。"
    },
    "maintenanceBlankKeptBlank": {
        "en": "If Product Line or Production Site is blank, the platform keeps it blank and does not infer automatically.",
        "zh": "若 Product Line 或 Production Site 空白，平台會保持空白，不自動推論。"
    },
    "noFileSelected": {
        "en": "No file selected",
        "zh": "尚未選擇檔案"
    },
    "pleaseSelectRuleMasterBeforeImport": {
        "en": "Please select a Rule Master file before importing.",
        "zh": "請先選擇 Rule Master 檔案後再匯入。"
    },
    "pleaseSelectRuleFileFirst": {
        "en": "Please select a rule file first.",
        "zh": "請先選擇規則檔。"
    },
    "importFailed": {
        "en": "Import failed",
        "zh": "匯入失敗"
    },
    "importingRuleMaster": {
        "en": "Importing Rule Master...",
        "zh": "正在匯入 Rule Master..."
    },
    "uploadingRuleFileTo": {
        "en": "Uploading rule file to {ruleSet}.",
        "zh": "正在上傳規則檔至 {ruleSet}。"
    },
    "ruleMasterImportedSuccessfully": {
        "en": "Rule Master imported successfully.",
        "zh": "Rule Master 匯入成功。"
    },
    "ruleImportSuccessDetail": {
        "en": "Rule Set: {ruleSet} · Total Rules: {count}",
        "zh": "規則組：{ruleSet} · 規則總數：{count}"
    },
    "importSuccess": {
        "en": "Import Success",
        "zh": "匯入成功"
    },
    "ruleSetLabel": {
        "en": "Rule Set",
        "zh": "規則組"
    },
    "totalRulesLabel": {
        "en": "Total rules",
        "zh": "規則總數"
    },
    "importedAtLabel": {
        "en": "Imported at",
        "zh": "匯入時間"
    },
    "ruleMasterImportFailed": {
        "en": "Rule Master import failed",
        "zh": "Rule Master 匯入失敗"
    },
    "pleaseCheckFileFormat": {
        "en": "Please check the file format.",
        "zh": "請確認檔案格式。"
    },
    "errorPrefix": {
        "en": "Error: ",
        "zh": "錯誤："
    },
    "readyForProcessing": {
        "en": "Ready for processing",
        "zh": "準備處理"
    },
    "readyForBatchFormatting": {
        "en": "Ready for batch formatting",
        "zh": "準備批次格式化"
    },
    "readyForBomExpansion": {
        "en": "Ready for BOM Expansion",
        "zh": "準備 BOM 展開"
    },
    "idle": {
        "en": "Idle",
        "zh": "待命"
    },
    "completed": {
        "en": "Completed",
        "zh": "已完成"
    },
    "processingInProgress": {
        "en": "Processing in progress",
        "zh": "處理中"
    },
    "processingCompleted": {
        "en": "Processing completed",
        "zh": "處理完成"
    },
    "processingFailed": {
        "en": "Processing failed",
        "zh": "處理失敗"
    }
};

  const phraseZh = {
    "Carbon Management Platform": "資料整合平台",
    "Manufacturing Data Layer": "製造資料層",
    "Modules": "模組",
    "Module": "模組",
    "Product Data Preparation": "產品資料準備",
    "Rule Management": "規則管理",
    "BOM Expansion": "BOM 展開",
    "Upload Supplier Files": "上傳供應商檔案",
    "Optional. Supports multiple Excel files. Uses Raw Material Code + Vendor number to write Supplier Name. Transportation Origin uses Supplier Address or Country + City + Street. Also exports supplier_bulk_create from the built-in template.": "選填，可多檔上傳。系統會依原物料代碼與 Vendor number 寫入 Supplier Name；Transportation Origin 使用 Supplier Address，若無則使用 Country + City + Street，並由內建範本輸出 supplier_bulk_create。",
  "Download Supplier Bulk Create": "下載供應商建立 Bulk",
    "Matching supplier files and dropdown options...": "正在比對供應商檔案與下拉選項...",
    "Carbon Emission Factor Selection": "碳排放係數選擇",
    "PCF Calculation": "產品碳足跡計算",
    "Production Environment": "正式環境",
    "Online | Version 1.0": "線上｜版本 1.0",
    "Rule Master Enabled": "Rule Master 已啟用",
    "Multi-file Upload": "多檔上傳",
    "Version 1.0": "版本 1.0",
    "Work order & batch preparation": "工單與批次資料準備",
    "Multi-level BOM explosion": "多階 BOM 展開",
    "Factor mapping and selection": "係數對應與選擇",
    "Product carbon footprint": "產品碳足跡",
    "Business Unit Rule Set": "BU 規則組",
    "Production Quantity Work Orders": "生產數量工單",
    "Upload SAP Working Hour Orders": "生產工時工單",
    "Labor Allocation Source": "工時來源",
    "Reporting Year": "報告年度",
    "e.g. 2024; blank = all years": "例如：2024；空白＝全部年度",
    "Run Consolidation & Classification": "執行合併與分類",
    "Download Module 1A Output Excel": "下載 Module 1A 輸出 Excel",
    "Module 1A · Work Order Processing": "Module 1A · 工單處理",
    "Upload one or multiple SAP production work order files.": "上傳一份或多份 SAP 生產工單檔案。",
    "Module 1A": "Module 1A",
    "Module 1B · Batch Data Formatting": "Module 1B · 批次資料格式化",
    "Convert Module 1A output and batch file template into a standardized batch file.": "將 Module 1A 輸出與批次範本轉換為標準化批次檔。",
    "Module 1B": "Module 1B",
    "Module 1A Output": "Module 1A 輸出",
    "Annual output & classification result": "年度產量與分類結果",
    "Batch Template": "批次範本",
    "Required batch file format": "必要批次檔格式",
    "Formatted Batch File": "已格式化批次檔",
    "Ready for downstream processing": "可供後續處理",
    "Upload Module 1A Output File": "上傳 Module 1A 輸出檔",
    "Upload Bulk Template File": "上傳 Bulk 範本檔",
    "Working Hour Source": "工時來源",
    "Direct Working Hour": "僅成品工時",
    "Include Semi-finished Working Hour": "包含半品工時",
    "This option requires Module 2A working_hour_rollup. Please complete Module 2A first.": "此選項需先完成 Module 2A working_hour_rollup。請先完成 Module 2A。",
    "Generate Formatted Bulk File": "產生格式化 Bulk 檔",
    "Download Formatted Bulk File": "下載格式化 Bulk 檔",
    "Download Bulk ZIP File": "下載 Bulk ZIP 檔",
    "Module 1B extracts required fields from the Module 1A output and writes them into the bulk template. Activity Data and Products sheets will be populated automatically.": "Module 1B 會從 Module 1A 輸出擷取必要欄位並寫入 Bulk 範本。系統會自動填入 Activity Data 與 Products 分頁。",
    "Execution Log": "執行紀錄",
    "Summary": "摘要",
    "Rules": "規則",
    "Version 1.0 Decision Flow": "版本 1.0 判斷流程",
    "Required SAP Fields": "必要 SAP 欄位",
    "Upload Standard BOM": "上傳標準 BOM",
    "Duplicates Removed": "移除重複筆數",
    "BOM Rows After Dedup": "去重後 BOM 筆數",
    "BOM Rows Before Dedup": "去重前 BOM 筆數",
    "BOM Files": "BOM 檔案數",
    "Reading and merging BOM file(s)...": "讀取並合併 BOM 檔案...",
    "Reading standard BOM file(s) and validating raw material bulk template.": "讀取標準 BOM 檔案並檢查原物料 Bulk 範本。",
    "Upload Standard BOM Files": "上傳標準 BOM 檔案",
    "Upload Raw Material Bulk Template": "上傳原物料 Bulk 範本",
    "Process BOM Expansion": "執行 BOM 展開",
    "Download Raw Material Bulk": "下載原物料 Bulk",
    "Download Working Hour Roll-up": "下載工時 Roll-up",
    "Upload Annual Product Output & Classification Result": "上傳年度產品產量與分類結果",
    "Working Hour Roll-up Rows": "工時 Roll-up 筆數",
    "Semi Working Hours": "半品工時",
    "BOM Expansion Logic": "BOM 展開邏輯",
    "Configure source column names for BOM Expansion. Use Default for system settings or Confirm to apply your input.": "設定 BOM 展開來源欄位名稱。使用「預設」套用系統設定，或使用「確認」套用輸入內容。",
    "Parent Material": "母件料號",
    "Component": "元件料號",
    "Quantity": "數量",
    "Unit": "單位",
    "Component Description": "元件描述",
    "Material Group": "物料群組",
    "Valid From": "有效起始日",
    "Default": "預設",
    "Confirm": "確認",
    "Enter BOM column name": "輸入 BOM 欄位名稱",
    "Semi-finished Rule": "半成品判斷規則",
    "Component also exists as Parent Node": "Component 同時存在於 Parent Node",
    "Quantity Roll-up": "數量展開邏輯",
    "Multiply quantities across all BOM levels": "跨 BOM 階層累乘數量",
    "Raw Material Bulk Output": "原物料 Bulk 輸出",
    "Input Sheet Activity Data and Input Sheet Raw Material will be populated. Optional fields are not written.": "系統會填入 Input Sheet Activity Data 與 Input Sheet Raw Material。Optional 欄位不會寫入。",
    "Rule Master": "Rule Master",
    "Import and download classification rule masters for Product Data Preparation.": "匯入與下載產品資料準備使用的分類規則主檔。",
    "Rule Management is a sub-function under Product Data Preparation. It controls product type classification, product series mapping, WIP judgment and customer mapping logic.": "規則管理是產品資料準備下的子功能，用於控制產品類型分類、產品系列對應、WIP 判斷與客戶對應邏輯。",
    "Rule Master ready": "Rule Master 已就緒",
    "Select a Business Unit Rule Set and import a Rule Master file.": "請選擇 BU 規則組並匯入 Rule Master 檔案。",
    "Current Rule Set": "目前規則組",
    "Current Rules": "目前規則數",
    "Last Import": "最近匯入",
    "Upload Rule Master": "上傳 Rule Master",
    "Import Rule Master": "匯入 Rule Master",
    "Download Rule Master": "下載 Rule Master",
    "Download Product Series Master": "下載 Product Series Master",
    "Rule Master Fields": "Rule Master 欄位",
    "Rule Type Guide": "Rule Type 說明",
    "Plant Exact": "Plant Exact",
    "Plant exact match, for example: 3760 → Shijie Plant-IPS": "Plant 完全符合，例如：3760 → 石碣廠-IPS",
    "Material Number Exact": "Material Number Exact",
    "Exact material number match, for example: SG-96000-00A": "完整料號比對，例如：SG-96000-00A",
    "Material Number Prefix": "Material Number Prefix",
    "Material number prefix match, for example: 850-, 851-, 852-": "料號前綴比對，例如：850-、851-、852-",
    "Series Exact": "Series Exact",
    "Exact product series match, for example: SN3103B02": "完整產品系列比對，例如：SN3103B02",
    "Series Prefix": "Series Prefix",
    "Product series prefix match, for example: SN, SP, FU, SCMC": "產品系列前綴比對，例如：SN、SP、FU、SCMC",
    "Description Contains": "Description Contains",
    "Material description contains keywords, for example: TOUCHPAD MODULE, ASSY": "Material description 包含關鍵字，例如：TOUCHPAD MODULE、ASSY",
    "Maintenance Principles": "維護原則",
    "The smaller the Priority number, the higher the priority.": "Priority 數字越小，優先度越高。",
    "Product Line and Production Site are controlled only by Rule Master.": "Product Line 與 Production Site 僅由 Rule Master 控制。",
    "When adding or adjusting classification rules, only modify the corresponding BU rule_master.csv; main.py does not need to be modified.": "新增或調整分類規則時，只需修改對應 BU 的 rule_master.csv，不需修改 main.py。",
    "If Product Line or Production Site is blank, the platform keeps it blank and does not infer automatically.": "若 Product Line 或 Production Site 空白，平台會保持空白，不自動推論。",
    "No file selected": "尚未選擇檔案",
    "Please select a Rule Master file before importing.": "請先選擇 Rule Master 檔案後再匯入。",
    "Please select a rule file first.": "請先選擇規則檔。",
    "Import failed": "匯入失敗",
    "Importing Rule Master...": "正在匯入 Rule Master...",
    "Uploading rule file to {ruleSet}.": "正在上傳規則檔至 {ruleSet}。",
    "Rule Master imported successfully.": "Rule Master 匯入成功。",
    "Rule Set: {ruleSet} · Total Rules: {count}": "規則組：{ruleSet} · 規則總數：{count}",
    "Import Success": "匯入成功",
    "Rule Set": "規則組",
    "Total rules": "規則總數",
    "Imported at": "匯入時間",
    "Rule Master import failed": "Rule Master 匯入失敗",
    "Please check the file format.": "請確認檔案格式。",
    "Error: ": "錯誤：",
    "Ready for processing": "準備處理",
    "Ready for batch formatting": "準備批次格式化",
    "Ready for BOM Expansion": "準備 BOM 展開",
    "Idle": "待命",
    "Completed": "已完成",
    "Processing in progress": "處理中",
    "Processing completed": "處理完成",
    "Processing failed": "處理失敗",
    "Integrate manufacturing data, expand BOM structures, select emission factors, and support product carbon footprint workflows.": "整合製造資料、展開 BOM 結構、選擇碳排放係數，並支援產品碳足跡流程。",
    "Prepare production output and batch data for product carbon footprint workflows.": "準備產品碳足跡流程所需的生產產出與批次資料。",
    "Maintain Product Data Preparation Rules, including Rule Master and Product Series Master.": "維護 Product Data Preparation Rules，包含 Rule Master 與 Product Series Master。",
    "Expand multi-level BOM structures, aggregate total raw material demand for finished products, and generate Raw Material Bulk files.": "展開多階 BOM 結構、彙總成品需求原物料總數量，並產生原物料 Bulk 檔。",
    "Complete Module 1A Work Order Processing, Module 1B Batch Data Formatting and Rule Management.": "完成 Module 1A 工單處理、Module 1B 批次資料格式化與規則管理。",
    "Reserved module. This area can be extended for multi-level BOM explosion.": "預留模組。此區可延伸為多階 BOM 展開。",
    "Reserved module. This area can be extended for emission factor mapping and selection.": "預留模組。此區可延伸為排放係數對應與選擇。",
    "Reserved module. This area can be extended for product carbon footprint calculation.": "預留模組。此區可延伸為產品碳足跡計算。",
    "Upload production quantity work orders and optional labor work orders to start the classification workflow.": "上傳生產數量工單與選填生產工時工單後開始分類流程。",
    "Ready. Upload production quantity work orders and optional labor work orders to start processing.": "準備完成。請上傳生產數量工單與選填生產工時工單後開始處理。",
    "Reading Excel files and applying classification rules. Please keep this page open.": "正在讀取 Excel 並套用分類規則，請保持頁面開啟。",
    "Reading uploaded Excel files...": "讀取上傳的 Excel 檔案...",
    "Merging SAP production work orders...": "合併 SAP 生產工單...",
    "Filtering reporting year...": "篩選報告年度...",
    "Extracting product series...": "解析產品系列...",
    "Applying Rule Master and Product Series Master...": "套用 Rule Master 與 Product Series Master...",
    "Generating output Excel...": "產生輸出 Excel...",
    "Classification result is ready. You can download the Excel output.": "分類結果已完成，可下載 Excel 輸出檔。",
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
    "BOM Expansion is ready.": "BOM 展開已準備完成。",
    "Batch formatting in progress": "批次格式化處理中",
    "Reading Module 1A output and validating the bulk template.": "讀取 Module 1A 輸出並檢查 Bulk 範本。",
    "Reading Module 1A output file...": "讀取 Module 1A 輸出檔...",
    "Copying original bulk template...": "複製原始 Bulk 範本...",
    "Writing Input Sheet Activity Data...": "寫入 Input Sheet Activity Data...",
    "Writing Input Sheet Products...": "寫入 Input Sheet Products...",
    "Preserving template formatting and validation...": "保留範本格式與資料驗證...",
    "Generating formatted bulk file...": "產生格式化 Bulk 檔...",
    "Batch formatting completed": "批次格式化完成",
    "Batch formatting failed": "批次格式化失敗",
    "Please review the input files and template format.": "請檢查輸入檔案與範本格式。",
    "Checking latest BOM Expansion result for semi-finished working hours.": "正在檢查最新 BOM Expansion 結果以納入半品工時。",
    "Direct Working Hour Enabled.": "已啟用僅成品工時。",
    "Processing...": "處理中...",
    "Completed.": "已完成。",
    "Error:": "錯誤：",
    "Year": "年度",
    "Uploaded files": "上傳檔案數",
    "Uploaded Files": "上傳檔案數",
    "Labor files": "工時檔案數",
    "Labor Files": "工時檔案數",
    "Work order rows": "工單筆數",
    "Work Order Rows": "工單筆數",
    "Summary rows": "彙總筆數",
    "Summary Rows": "彙總筆數",
    "Total output": "總產量",
    "Total Output": "總產量",
    "Total hours": "總工時",
    "Total Hours": "總工時",
    "WIP rows": "WIP 筆數",
    "WIP Rows": "WIP 筆數",
    "Files in ZIP": "ZIP 檔案數",
    "Activity Data": "Activity Data",
    "Products": "Products",
    "Excluded WIP": "排除 WIP",
    "Working Hour": "工時來源",
    "Semi-finished": "半品",
    "Semi-finished Components": "半品元件數",
    "Raw material rows": "原物料筆數",
    "Raw Materials": "原物料數",
    "Max BOM Level": "最大 BOM 階層",
    "Current Setting:": "目前設定：",
    "Rule Management is ready.": "規則管理已準備完成。",
    "Labor HR.Act + FOH-Others.Act": "人員+設備工時",
    "Labor HR.Act Only": "人員工時",
    "FOH-Others.Act Only": "設備工時",
    "direct": "僅成品",
    "include_semi": "含半品"
};


  Object.assign(phraseZh, {
    "Reading Excel files and applying classification rules. Please keep this page open.": "正在讀取 Excel 並套用分類規則，請保持頁面開啟。",
    "Reading uploaded Excel files...": "讀取上傳的 Excel 檔案...",
    "Merging SAP production work orders...": "合併 SAP 生產工單...",
    "Filtering reporting year...": "篩選報告年度...",
    "Extracting product series...": "解析產品系列...",
    "Applying Rule Master and Product Series Master...": "套用 Rule Master 與 Product Series Master...",
    "Generating output Excel...": "產生輸出 Excel...",
    "Processing in progress": "處理中",
    "Processing completed": "處理完成",
    "Processing failed": "處理失敗",
    "Classification result is ready. You can download the Excel output.": "分類結果已完成，可下載 Excel 輸出檔。",
    "Please review the error message and input file format.": "請確認錯誤訊息與輸入檔格式。",
    "Ready for processing": "準備處理",
    "Idle": "待命",
    "Completed": "已完成",
    "Error": "錯誤",

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

    "Batch formatting in progress": "批次格式化處理中",
    "Reading Module 1A output and validating the bulk template.": "讀取 Module 1A 輸出並檢查 Bulk 範本。",
    "Reading Module 1A output file...": "讀取 Module 1A 輸出檔...",
    "Copying original bulk template...": "複製原始 Bulk 範本...",
    "Writing Input Sheet Activity Data...": "寫入 Input Sheet Activity Data...",
    "Writing Input Sheet Products...": "寫入 Input Sheet Products...",
    "Preserving template formatting and validation...": "保留範本格式與資料驗證...",
    "Generating formatted bulk file...": "產生格式化 Bulk 檔...",
    "Batch formatting completed": "批次格式化完成",
    "Batch formatting failed": "批次格式化失敗",
    "Please review the input files and template format.": "請檢查輸入檔案與範本格式。",
    "Direct Working Hour Enabled.": "已啟用僅成品工時。",
    "Checking latest BOM Expansion result for semi-finished working hours.": "正在檢查最新 BOM Expansion 結果以納入半品工時。"
  });


  Object.assign(phraseZh, {
    "CCL Mapping": "CCL係數資料庫",
    "Factor Library": "Ecoinvent係數資料庫",
    "Stage 2 entry page. Choose CCL Mapping or Factor Library to enter each workspace.": "模組3 入口頁面。請選擇以CCL 係數帶入或查詢Ecoinvent係數資料庫，進入各自專區。",
    "Stage 2: A/B workspaces are available for CCL Mapping and Factor Library implementation.": "A專區為CCLibrary係數資料庫，B專區為Ecoinvent係數資料庫。",
    "This version only creates the A/B workspace entry points and does not call any backend API yet. Module 1 and Module 2 routes, forms, Excel logic, and outputs remain unchanged.": "CCL係數資料庫專區以原物料料號對應，帶入CCL Item和碳係數；Ecoinvent係數資料庫專區主要為查詢功能，未來將導入建議係數功能。",
    "Use Module 1A Annual Product Output & Classification Result": "引用 Module 1A 年度產品產量與分類結果",
    "Expand multi-level BOM structures, roll up raw material quantities, and generate raw material bulk files.": "展開多階 BOM 結構、彙總成品需求原物料總數量，並產生原物料 Bulk 檔。"
  });

  const phraseEn = {};

  const preserveExact = new Set([
    "DIP", "SAP", "BOM", "PCF", "WIP", "NB", "TP", "SCMC", "SN", "SP", "SM", "SK",
    "IPS", "AE", "PC&CE", "PC_CE", "-",
    "Order、Plant、Material Number、Material description、Delivered quantity (GMEIN)、Actual finish date",
    "Priority、Rule Type、Key、Product Type、Product Line、Production Site、Customer、Customer Code Logic、Is_WIP、Enabled",
    "Parent Node", "CS03 Qty", "CS03 UoM", "Material group", "BOM Valid From"
  ]);

  function t(key, fallback) {
    if (keyed[key] && keyed[key][currentLang]) return keyed[key][currentLang];
    if (fallback && currentLang === "zh" && phraseZh[fallback]) return phraseZh[fallback];
    if (fallback && currentLang === "en" && phraseEn[fallback]) return phraseEn[fallback];
    return fallback || key;
  }

  function format(key, fallback, values) {
    let text = t(key, fallback);
    Object.keys(values || {}).forEach(function (name) {
      text = text.split("{" + name + "}").join(values[name]);
    });
    return text;
  }

  function isSkippableText(text) {
    const value = String(text || "").trim();
    if (!value) return true;
    if (preserveExact.has(value)) return true;
    if (/^~?\d+s remaining$/.test(value)) return true;
    if (/^約\s*\d+\s*秒$/.test(value)) return true;
    if (/^~?\d+秒$/.test(value)) return true;
    if (/^\d+$/.test(value)) return true;
    return false;
  }

  function translateString(input, targetLang) {
    return input;
  }

  function translateKeyedElements(targetLang) {
    document.querySelectorAll("[data-i18n]").forEach(function (el) {
      const key = el.getAttribute("data-i18n");
      if (keyed[key] && keyed[key][targetLang]) {
        el.textContent = keyed[key][targetLang];
      }
    });

    document.querySelectorAll("[data-i18n-dynamic]").forEach(function (el) {
      const key = el.getAttribute("data-i18n-dynamic");
      if (keyed[key] && keyed[key][targetLang]) {
        el.textContent = keyed[key][targetLang];
      }
    });

    document.querySelectorAll("[data-i18n-placeholder]").forEach(function (el) {
      const key = el.getAttribute("data-i18n-placeholder");
      if (keyed[key] && keyed[key][targetLang]) {
        el.setAttribute("placeholder", keyed[key][targetLang]);
      }
    });
  }

  function normalizeSelectOptions() {
    const labor = document.getElementById("laborMode");
    if (labor) {
      const opts = currentLang === "zh"
        ? { both: "人員+設備工時", labor_hr: "人員工時", foh: "設備工時" }
        : { both: "Labor HR.Act + FOH-Others.Act", labor_hr: "Labor HR.Act Only", foh: "FOH-Others.Act Only" };
      Array.from(labor.options).forEach(function (option) {
        if (opts[option.value]) option.textContent = opts[option.value];
      });
    }

    const bulk = document.getElementById("bulkWorkingHourSource");
    if (bulk) {
      const opts = currentLang === "zh"
        ? { direct: "僅成品工時", include_semi: "包含半品工時" }
        : { direct: "Direct Working Hour", include_semi: "Include Semi-finished Working Hour" };
      Array.from(bulk.options).forEach(function (option) {
        if (opts[option.value]) option.textContent = opts[option.value];
      });
    }
  }

  function translateTextNodes(root, targetLang) {
    return;
  }

  function translateAttributes(targetLang) {
    document.querySelectorAll("[data-i18n-aria-label]").forEach(function (el) {
      const key = el.getAttribute("data-i18n-aria-label");
      if (keyed[key] && keyed[key][targetLang]) {
        el.setAttribute("aria-label", keyed[key][targetLang]);
      }
    });

    if (document.title) {
      document.title = targetLang === "zh" ? "資料整合平台" : "Data Integration Platform (DIP)";
    }
  }

  function applyLanguage(targetLang) {
    if (isApplying) return;
    isApplying = true;

    currentLang = targetLang || currentLang;
    localStorage.setItem(STORAGE_KEY, currentLang);
    localStorage.setItem("cmp_lang", currentLang);
    document.documentElement.lang = currentLang === "zh" ? "zh-Hant" : "en";
    button.textContent = currentLang === "zh" ? "EN" : "中";

    translateKeyedElements(currentLang);
    translateTextNodes(document.body, currentLang);
    translateAttributes(currentLang);
    normalizeSelectOptions();

    isApplying = false;
  }

  window.CMPI18N = {
    t: function (key, fallback) { return t(key, fallback); },
    format: function (key, fallback, values) { return format(key, fallback, values); },
    apply: function (lang) { applyLanguage(lang || currentLang); },
    current: function () { return currentLang; }
  };

  button.addEventListener("click", function (event) {
    event.preventDefault();
    event.stopPropagation();
    try {
      applyLanguage(currentLang === "en" ? "zh" : "en");
    } catch (err) {
      console.error("DIP i18n toggle failed:", err);
    }
  });

  const observer = new MutationObserver(function (mutations) {
    if (isApplying) return;
    window.setTimeout(function () {
      translateKeyedElements(currentLang);
      translateAttributes(currentLang);
      normalizeSelectOptions();
    }, 0);
  });

  observer.observe(document.body, {
    childList: true,
    subtree: true
  });

  applyLanguage(currentLang);
})();


// MODULE3 Stage 2 v15 i18n patch
try {
  if (typeof translations !== 'undefined') {
    if (translations.zh) {
      translations.zh.factorDetailKicker = translations.zh.factorDetailKicker || '係數詳情';
      translations.zh.factorDetailTitle = translations.zh.factorDetailTitle || '係數詳細說明';
      translations.zh.referenceProductName = translations.zh.referenceProductName || 'Reference Product Name';
      translations.zh.activityKeywordSearch = translations.zh.activityKeywordSearch || '關鍵字查詢';
      translations.zh.referenceNameSearch = translations.zh.referenceNameSearch || '名稱查詢';
      translations.zh.copy = translations.zh.copy || '複製';
      translations.zh.copied = translations.zh.copied || '已複製';
      translations.zh.lciaIndicator = translations.zh.lciaIndicator || 'LCIA 指標';
      translations.zh.clickForFactorDetail = translations.zh.clickForFactorDetail || '點選查看係數詳細說明';
      translations.zh.close = translations.zh.close || '關閉';
    }
    if (translations.en) {
      translations.en.factorDetailKicker = translations.en.factorDetailKicker || 'Factor Detail';
      translations.en.factorDetailTitle = translations.en.factorDetailTitle || 'Factor Detail';
      translations.en.referenceProductName = translations.en.referenceProductName || 'Reference Product Name';
      translations.en.activityKeywordSearch = translations.en.activityKeywordSearch || 'Keyword search';
      translations.en.referenceNameSearch = translations.en.referenceNameSearch || 'Name search';
      translations.en.copy = translations.en.copy || 'Copy';
      translations.en.copied = translations.en.copied || 'Copied';
      translations.en.lciaIndicator = translations.en.lciaIndicator || 'LCIA Indicator';
      translations.en.clickForFactorDetail = translations.en.clickForFactorDetail || 'Click to view factor detail';
      translations.en.close = translations.en.close || 'Close';
    }
  }
} catch (e) { console.warn('MODULE3 v15 i18n patch skipped', e); }
