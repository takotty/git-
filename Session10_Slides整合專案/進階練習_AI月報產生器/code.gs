// ============================================================
// 進階練習：AI 月報簡報自動產生器
// 對應：Session 10（Slides 整合、跨服務操作）
// 進階挑戰：串接 ChatGPT API 自動產生文字分析
// ============================================================

// ============================================================
// ⚙️  全域設定區（請依需求修改）
// ============================================================

var AI_CONFIG = {
  // ① 填入你的 OpenAI API Key（建議用 PropertiesService 儲存，見下方說明）
  //   → 取得網址：https://platform.openai.com/api-keys
  //   → 安全做法：執行「設定 API Key」函式，Key 不會出現在程式碼中
  OPENAI_API_KEY: "",          // 直接填入，或留空使用 PropertiesService

  // ② 使用的模型（預設 gpt-4o-mini，便宜且快速）
  MODEL: "gpt-4o-mini",

  // ③ 回應最大 Token 數（分析文字約 400~800 token 即可）
  MAX_TOKENS: 800,

  // ④ 溫度（0=保守/穩定，1=創意）
  TEMPERATURE: 0.7,

  // ⑤ 逾時秒數
  TIMEOUT_MS: 30000
};

// ============================================================
// 🔑  API Key 安全管理（建議使用，避免 Key 寫死在程式碼）
// ============================================================

/**
 * 設定 OpenAI API Key（只需執行一次）
 * Key 存入 PropertiesService，不會出現在程式碼中
 */
function 設定API_Key() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    "🔑 設定 OpenAI API Key",
    "請貼上你的 API Key（格式：sk-...）\n\n" +
    "Key 將安全存入 Script Properties，不會外洩。",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() === ui.Button.OK) {
    var key = result.getResponseText().trim();
    if (key.startsWith("sk-")) {
      PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", key);
      ui.alert("✅ API Key 已安全儲存！\n\n現在可以執行「產生 AI 月報簡報」。");
    } else {
      ui.alert("❌ 格式不正確！\nOpenAI API Key 應以 sk- 開頭。");
    }
  }
}

/**
 * 取得 API Key（優先使用 PropertiesService，其次用 CONFIG 中的值）
 * @returns {string} API Key
 */
function 取得API_Key_() {
  var fromProps = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  return fromProps || AI_CONFIG.OPENAI_API_KEY;
}

/**
 * 清除已儲存的 API Key
 */
function 清除API_Key() {
  PropertiesService.getScriptProperties().deleteProperty("OPENAI_API_KEY");
  SpreadsheetApp.getUi().alert("🗑️ API Key 已清除。");
}

// ============================================================
// 🤖  ChatGPT API 串接核心函式
// ============================================================

/**
 * 呼叫 ChatGPT API 取得文字分析
 *
 * @param {string} systemPrompt - 系統角色指示（設定 AI 的角色與回應風格）
 * @param {string} userPrompt   - 用戶問題/資料（送給 AI 分析的內容）
 * @returns {string} AI 回應的文字；若失敗則回傳 null
 */
function 呼叫ChatGPT_(systemPrompt, userPrompt) {
  var apiKey = 取得API_Key_();

  if (!apiKey) {
    Logger.log("⚠️ 未設定 API Key，跳過 AI 分析。");
    return null;
  }

  var payload = {
    model: AI_CONFIG.MODEL,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user",   content: userPrompt   }
    ],
    max_tokens:  AI_CONFIG.MAX_TOKENS,
    temperature: AI_CONFIG.TEMPERATURE
  };

  var options = {
    method:          "post",
    contentType:     "application/json",
    headers:         { Authorization: "Bearer " + apiKey },
    payload:         JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var 回應 = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
    var 狀態碼 = 回應.getResponseCode();

    if (狀態碼 === 200) {
      var json = JSON.parse(回應.getContentText());
      var 回應文字 = json.choices[0].message.content.trim();
      Logger.log("✅ ChatGPT 回應成功（" + 回應文字.length + " 字）");
      return 回應文字;
    } else {
      var 錯誤詳情 = JSON.parse(回應.getContentText());
      Logger.log("❌ API 錯誤 " + 狀態碼 + "：" + JSON.stringify(錯誤詳情));
      return null;
    }
  } catch (e) {
    Logger.log("❌ 網路錯誤：" + e.message);
    return null;
  }
}

/**
 * 產生月報 AI 分析文字
 *
 * @param {Object} 統計 - 包含 總營收、總成本、利潤、利潤率、新客戶、專案完成
 * @param {Array}  部門資料 - 二維陣列，各部門詳細數字
 * @param {string} 月份 - 報告月份字串
 * @returns {Object} { 執行摘要, KPI分析, 風險建議, 下月目標 }
 */
function 產生月報AI分析_(統計, 部門資料, 月份) {

  // ── 組合給 AI 的資料摘要 ──────────────────────────────────
  var 部門文字 = 部門資料.slice(1).map(function(row) {
    return row[0] + "：營收 NT$" + (row[1] || 0).toLocaleString() +
           "，成本 NT$" + (row[2] || 0).toLocaleString() +
           "，新客戶 " + (row[3] || 0) + " 家" +
           "，完成專案 " + (row[4] || 0) + " 個";
  }).join("\n");

  var 資料摘要 =
    "【" + 月份 + " 月度營運數據】\n" +
    "─────────────────────────\n" +
    "年度總營收：NT$ " + (統計.總營收 / 10000).toFixed(0) + " 萬\n" +
    "年度總成本：NT$ " + (統計.總成本 / 10000).toFixed(0) + " 萬\n" +
    "淨利潤：    NT$ " + (統計.利潤 / 10000).toFixed(0) + " 萬\n" +
    "利潤率：    " + 統計.利潤率 + "%\n" +
    "新增客戶：  " + 統計.新客戶 + " 家\n" +
    "完成專案：  " + 統計.專案完成 + " 個\n\n" +
    "【各部門詳情】\n" + 部門文字;

  // ── System Prompt：設定 AI 角色 ───────────────────────────
  var systemPrompt =
    "你是一位資深企業財務分析師，專精於月度營運報告撰寫。" +
    "請使用繁體中文，語氣專業且簡潔。" +
    "回應須條列式，每點不超過 40 字，避免過度修飾語言。";

  // ── 依分析維度分別呼叫 API（或一次呼叫取得完整分析）──────
  // 採用「一次呼叫，結構化輸出」策略，節省 API 費用
  var userPrompt =
    "請根據以下數據，產生月報分析，格式必須完全按照以下結構輸出：\n\n" +
    資料摘要 + "\n\n" +
    "請依以下格式回應（每區塊之間空一行）：\n\n" +
    "【執行摘要】\n" +
    "（2~3 句話總結本月整體表現）\n\n" +
    "【KPI 分析】\n" +
    "• （利潤率分析）\n" +
    "• （客戶成長分析）\n" +
    "• （專案完成率分析）\n\n" +
    "【風險與建議】\n" +
    "• （風險點 1）\n" +
    "• （風險點 2）\n" +
    "• （改善建議）\n\n" +
    "【下月目標】\n" +
    "• （具體目標 1，含數字）\n" +
    "• （具體目標 2，含數字）\n" +
    "• （具體目標 3，含數字）";

  var ai回應 = 呼叫ChatGPT_(systemPrompt, userPrompt);

  if (!ai回應) {
    // 降級：用規則式產生基礎分析（無 API 也能執行）
    return 產生規則式分析_(統計);
  }

  // ── 解析 AI 回應，分割成各區塊 ───────────────────────────
  var 結果 = {
    執行摘要: "",
    KPI分析:  "",
    風險建議: "",
    下月目標: "",
    使用AI: true
  };

  var 區塊對應 = {
    "【執行摘要】": "執行摘要",
    "【KPI 分析】": "KPI分析",
    "【風險與建議】": "風險建議",
    "【下月目標】": "下月目標"
  };

  var 當前區塊 = null;
  var 行列表 = ai回應.split("\n");

  行列表.forEach(function(行) {
    var 找到區塊 = false;
    for (var 標籤 in 區塊對應) {
      if (行.indexOf(標籤) !== -1) {
        當前區塊 = 區塊對應[標籤];
        找到區塊 = true;
        break;
      }
    }
    if (!找到區塊 && 當前區塊 && 行.trim()) {
      結果[當前區塊] += (結果[當前區塊] ? "\n" : "") + 行.trim();
    }
  });

  return 結果;
}

/**
 * 規則式分析（當 API 不可用時的降級方案）
 * @param {Object} 統計
 * @returns {Object} 分析結果
 */
function 產生規則式分析_(統計) {
  var 利率 = Number(統計.利潤率);
  var 利潤評語 = 利率 >= 30 ? "利潤率優秀（≥30%），策略執行有效" :
                  利率 >= 20 ? "利潤率良好（20~30%），可進一步優化成本" :
                               "利潤率偏低（<20%），建議檢視成本結構";
  var 客戶評語 = 統計.新客戶 >= 20 ? "客戶拓展快速，須同步強化服務品質" :
                  統計.新客戶 >= 10 ? "客戶成長穩定，維持現有行銷策略" :
                                     "客戶拓展緩慢，建議加強推廣力道";

  return {
    執行摘要: "本月整體營收達 NT$" + (統計.總營收 / 10000).toFixed(0) + " 萬，" +
              "利潤率 " + 統計.利潤率 + "%，" +
              "新增客戶 " + 統計.新客戶 + " 家，完成 " + 統計.專案完成 + " 個專案。",
    KPI分析:  "• " + 利潤評語 + "\n" +
              "• " + 客戶評語 + "\n" +
              "• 專案完成率良好，本月共交付 " + 統計.專案完成 + " 個專案",
    風險建議: "• 持續監控成本支出，避免侵蝕利潤空間\n" +
              "• 強化客戶留存策略，降低流失率\n" +
              "• 建議定期檢視各部門 KPI 達成狀況",
    下月目標: "• 營收目標：NT$" + ((統計.總營收 * 1.1) / 10000).toFixed(0) + " 萬（成長 10%）\n" +
              "• 新增客戶：目標 " + Math.ceil(統計.新客戶 * 1.15) + " 家（成長 15%）\n" +
              "• 利潤率：維持在 " + Math.max(利率, 25).toFixed(1) + "% 以上",
    使用AI: false
  };
}

// ============================================================
// 🚀  一鍵產生 AI 月報簡報（主函式）
// ============================================================

/**
 * 一鍵產生 AI 風格月報簡報
 * 整合 Sheets 數據 + ChatGPT 文字分析 + 圖表 + Slides 排版
 */
function 產生AI月報簡報() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 開始 = new Date();
    var ui = SpreadsheetApp.getUi();

    var apiKey = 取得API_Key_();
    var ai模式提示 = apiKey
      ? "🤖 AI 分析模式（ChatGPT " + AI_CONFIG.MODEL + "）"
      : "📋 規則分析模式（未設定 API Key）";

    ui.alert("⏳ 正在生成 AI 月報簡報...\n" + ai模式提示 + "\n\n請稍候約 20~40 秒。");

    // ========== Step 1：讀取資料 ==========
    var 業績表 = ss.getSheetByName("月報資料");
    if (!業績表) {
      ui.alert("❌ 請先執行「初始化月報資料」！");
      return;
    }
    var 資料 = 業績表.getDataRange().getValues();

    // 統計計算
    var 統計 = { 總營收: 0, 總成本: 0, 新客戶: 0, 專案完成: 0 };
    for (var i = 1; i < 資料.length; i++) {
      統計.總營收   += 資料[i][1] || 0;
      統計.總成本   += 資料[i][2] || 0;
      統計.新客戶   += 資料[i][3] || 0;
      統計.專案完成 += 資料[i][4] || 0;
    }
    統計.利潤   = 統計.總營收 - 統計.總成本;
    統計.利潤率 = ((統計.利潤 / 統計.總營收) * 100).toFixed(1);

    var 月份 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月");

    // ========== Step 2：呼叫 ChatGPT API ==========
    Logger.log("Step 2：呼叫 ChatGPT API...");
    var ai分析 = 產生月報AI分析_(統計, 資料, 月份);
    Logger.log("AI 分析完成。使用AI=" + ai分析.使用AI);

    // ========== Step 3：建立簡報 ==========
    Logger.log("Step 3：建立簡報...");
    var ppt = SlidesApp.create("📊 " + 月份 + " 營運月報");

    // ─────────────────────────────────────────────
    // P1：封面
    // ─────────────────────────────────────────────
    var p1 = ppt.getSlides()[0];
    p1.getBackground().setSolidFill("#0a1628");
    p1.getPlaceholders().forEach(function(p) { p.remove(); });

    // AI 徽章
    var badge = p1.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 240, 28, 240, 35);
    badge.getFill().setSolidFill(ai分析.使用AI ? "#1a73e8" : "#546e7a");
    badge.getBorder().setTransparent();
    var badgeText = p1.insertTextBox(
      ai分析.使用AI ? "🤖 ChatGPT AI 自動分析" : "📋 規則式分析",
      250, 31, 220, 28
    );
    badgeText.getText().getTextStyle().setFontSize(11).setForegroundColor("#ffffff");
    badgeText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // 主標題
    var 標題 = p1.insertTextBox("📊 " + 月份 + "\n營運月報", 40, 110, 640, 130);
    標題.getText().getTextStyle().setFontSize(40).setBold(true).setForegroundColor("#ffffff");
    標題.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // 副標題
    var 副標 = p1.insertTextBox(
      Utilities.formatDate(new Date(), "Asia/Taipei", "報告日期：yyyy/MM/dd") +
      " ｜ " + ss.getName(),
      40, 265, 640, 30
    );
    副標.getText().getTextStyle().setFontSize(12).setForegroundColor("#64b5f6");
    副標.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // 裝飾線
    var 裝飾線 = p1.insertLine(SlidesApp.LineCategory.STRAIGHT, 160, 258, 560, 258);
    裝飾線.getLineFill().setSolidFill("#1a73e8");
    裝飾線.setWeight(2);

    // ─────────────────────────────────────────────
    // P2：KPI 摘要（含 AI 執行摘要）
    // ─────────────────────────────────────────────
    var p2 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    p2.getBackground().setSolidFill("#fafafa");

    var p2Title = p2.insertTextBox("🎯 關鍵營運指標 (KPI)", 40, 12, 640, 40);
    p2Title.getText().getTextStyle().setFontSize(26).setBold(true).setForegroundColor("#0a1628");

    // KPI 卡片
    var kpis = [
      { label: "總營收",  value: "NT$ " + (統計.總營收 / 10000).toFixed(0) + " 萬", color: "#1a73e8", icon: "💰" },
      { label: "淨利潤",  value: "NT$ " + (統計.利潤 / 10000).toFixed(0) + " 萬",   color: "#34a853", icon: "📈" },
      { label: "利潤率",  value: 統計.利潤率 + "%",                                   color: "#fbbc04", icon: "📊" },
      { label: "新客戶",  value: 統計.新客戶 + " 家",                                  color: "#ea4335", icon: "🤝" }
    ];

    kpis.forEach(function(kpi, idx) {
      var x = 30 + idx * 170;
      var card = p2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 68, 155, 140);
      card.getFill().setSolidFill(kpi.color);
      card.getBorder().setTransparent();

      var icon = p2.insertTextBox(kpi.icon, x + 10, 78, 135, 35);
      icon.getText().getTextStyle().setFontSize(28);
      icon.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      var val = p2.insertTextBox(kpi.value, x + 10, 113, 135, 35);
      val.getText().getTextStyle().setFontSize(17).setBold(true).setForegroundColor("#ffffff");
      val.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      var lbl = p2.insertTextBox(kpi.label, x + 10, 153, 135, 25);
      lbl.getText().getTextStyle().setFontSize(11).setForegroundColor("#e0e0e0");
      lbl.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    });

    // AI 執行摘要區塊
    var 摘要標籤 = p2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 228, 148, 26);
    摘要標籤.getFill().setSolidFill(ai分析.使用AI ? "#1a73e8" : "#546e7a");
    摘要標籤.getBorder().setTransparent();
    var 摘要標籤文字 = p2.insertTextBox(
      ai分析.使用AI ? "🤖 AI 執行摘要" : "📋 執行摘要",
      44, 230, 140, 22
    );
    摘要標籤文字.getText().getTextStyle().setFontSize(10).setBold(true).setForegroundColor("#ffffff");

    var 摘要框 = p2.insertTextBox(ai分析.執行摘要 || "（無摘要）", 40, 260, 640, 90);
    摘要框.getText().getTextStyle().setFontSize(13).setForegroundColor("#37474f");
    // 加入背景底色框
    var 摘要底框 = p2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 36, 255, 648, 100);
    摘要底框.getFill().setSolidFill("#e8f4fd");
    摘要底框.getBorder().setTransparent();
    // 重疊順序：先底框再文字（GAS Slides 無 z-index，以插入順序決定）
    摘要底框.sendToBack();

    // ─────────────────────────────────────────────
    // P3：🤖 AI KPI 分析頁（新增）
    // ─────────────────────────────────────────────
    var p3ai = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    p3ai.getBackground().setSolidFill("#0d1b2a");

    var p3aiTitle = p3ai.insertTextBox(
      (ai分析.使用AI ? "🤖 ChatGPT " : "📋 ") + "AI KPI 深度分析",
      40, 12, 640, 42
    );
    p3aiTitle.getText().getTextStyle().setFontSize(26).setBold(true).setForegroundColor("#64b5f6");

    // 分隔線
    var 分隔1 = p3ai.insertLine(SlidesApp.LineCategory.STRAIGHT, 40, 58, 680, 58);
    分隔1.getLineFill().setSolidFill("#1a73e8");
    分隔1.setWeight(1);

    // KPI 分析文字
    var kpiAnalysisBox = p3ai.insertTextBox(ai分析.KPI分析 || "（無分析）", 50, 70, 620, 160);
    kpiAnalysisBox.getText().getTextStyle().setFontSize(13).setForegroundColor("#e0e0e0");

    // 風險與建議標題
    var 風險標題 = p3ai.insertTextBox("⚠️ 風險與改善建議", 40, 240, 640, 35);
    風險標題.getText().getTextStyle().setFontSize(18).setBold(true).setForegroundColor("#fbbc04");

    var 風險框 = p3ai.insertTextBox(ai分析.風險建議 || "（無建議）", 50, 278, 620, 120);
    風險框.getText().getTextStyle().setFontSize(13).setForegroundColor("#cfd8dc");

    // ─────────────────────────────────────────────
    // P4：🎯 AI 下月目標頁（新增）
    // ─────────────────────────────────────────────
    var p4ai = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    p4ai.getBackground().setSolidFill("#fafafa");

    var p4aiTitle = p4ai.insertTextBox("🎯 下月目標設定", 40, 12, 640, 42);
    p4aiTitle.getText().getTextStyle().setFontSize(26).setBold(true).setForegroundColor("#0a1628");

    // 目標卡片背景
    var 目標底框 = p4ai.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 36, 62, 648, 160);
    目標底框.getFill().setSolidFill("#e8f5e9");
    目標底框.getBorder().setTransparent();

    var 目標框 = p4ai.insertTextBox(ai分析.下月目標 || "（無目標）", 50, 72, 620, 145);
    目標框.getText().getTextStyle().setFontSize(14).setForegroundColor("#1b5e20");

    // AI 標示
    var ai標注 = p4ai.insertTextBox(
      ai分析.使用AI
        ? "🤖 以上目標由 ChatGPT（" + AI_CONFIG.MODEL + "）基於本月數據自動生成"
        : "📋 以上目標由規則式邏輯自動計算",
      40, 232, 640, 25
    );
    ai標注.getText().getTextStyle().setFontSize(10).setForegroundColor("#9e9e9e");

    // 下半部：本月重點摘要數字
    var 數字小標 = p4ai.insertTextBox("📊 本月績效數字回顧", 40, 268, 640, 30);
    數字小標.getText().getTextStyle().setFontSize(16).setBold(true).setForegroundColor("#0a1628");

    var 績效數字 = [
      { 標籤: "完成專案", 值: 統計.專案完成 + " 個", 色: "#1a73e8" },
      { 標籤: "利潤率",   值: 統計.利潤率 + " %",    色: "#34a853" },
      { 標籤: "新增客戶", 值: 統計.新客戶 + " 家",    色: "#ea4335" }
    ];
    績效數字.forEach(function(item, idx) {
      var px = 40 + idx * 215;
      var 小卡 = p4ai.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, px, 305, 200, 75);
      小卡.getFill().setSolidFill(item.色);
      小卡.getBorder().setTransparent();
      var 小值 = p4ai.insertTextBox(item.值, px + 5, 315, 190, 35);
      小值.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#ffffff");
      小值.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
      var 小標 = p4ai.insertTextBox(item.標籤, px + 5, 352, 190, 22);
      小標.getText().getTextStyle().setFontSize(10).setForegroundColor("#e0e0e0");
      小標.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    });

    // ─────────────────────────────────────────────
    // P5：圖表頁
    // ─────────────────────────────────────────────
    var 圖表列表 = 業績表.getCharts();
    if (圖表列表.length > 0) {
      var p5 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      p5.getBackground().setSolidFill("#ffffff");
      var ct = p5.insertTextBox("📈 營運趨勢圖表", 40, 10, 640, 35);
      ct.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#0a1628");

      var blob = 圖表列表[0].getBlob();
      var img = p5.insertImage(blob);
      img.setLeft(40).setTop(52).setWidth(640).setHeight(355);
    }

    // ─────────────────────────────────────────────
    // P6：部門表格
    // ─────────────────────────────────────────────
    var p6 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    p6.getBackground().setSolidFill("#fafafa");
    var ttl = p6.insertTextBox("🏢 部門業績明細", 40, 10, 640, 35);
    ttl.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#0a1628");

    var rows = Math.min(資料.length, 8);
    var cols = Math.min(資料[0].length, 5);
    var table = p6.insertTable(rows, cols);
    for (var ri = 0; ri < rows; ri++) {
      for (var ci = 0; ci < cols; ci++) {
        var v = 資料[ri][ci];
        if (typeof v === "number" && ri > 0) v = "NT$ " + v.toLocaleString();
        table.getCell(ri, ci).getText().setText(String(v));
        table.getCell(ri, ci).getText().getTextStyle().setFontSize(10);
        if (ri === 0) {
          table.getCell(ri, ci).getFill().setSolidFill("#0a1628");
          table.getCell(ri, ci).getText().getTextStyle().setForegroundColor("#ffffff").setBold(true);
        } else if (ri % 2 === 0) {
          table.getCell(ri, ci).getFill().setSolidFill("#e3f2fd");
        }
      }
    }

    // ─────────────────────────────────────────────
    // P7：結尾
    // ─────────────────────────────────────────────
    var pEnd = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    pEnd.getBackground().setSolidFill("#0a1628");
    var end = pEnd.insertTextBox(
      "謝謝聆聽\n\n" + (ai分析.使用AI ? "🤖 本報告由 ChatGPT AI 自動分析產生" : "📋 本報告由規則式邏輯自動產生"),
      40, 130, 640, 130
    );
    end.getText().getTextStyle().setFontSize(28).setBold(true).setForegroundColor("#ffffff");
    end.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var endSub = pEnd.insertTextBox(
      "模型：" + (ai分析.使用AI ? AI_CONFIG.MODEL : "N/A（規則式）") +
      "  ｜  資料來源：" + ss.getName(),
      40, 278, 640, 25
    );
    endSub.getText().getTextStyle().setFontSize(10).setForegroundColor("#64b5f6");
    endSub.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // ========== 完成 ==========
    var 耗時 = Math.round((new Date() - 開始) / 1000);
    Logger.log("✅ 月報簡報完成！耗時 " + 耗時 + " 秒，共 " + ppt.getSlides().length + " 頁");

    ui.alert(
      "🎉 AI 月報簡報已完成！\n\n" +
      "📄 共 " + ppt.getSlides().length + " 頁  ｜  ⏱️ 耗時 " + 耗時 + " 秒\n" +
      (ai分析.使用AI
        ? "🤖 AI 分析：ChatGPT（" + AI_CONFIG.MODEL + "）\n"
        : "⚠️ 規則分析模式（未使用 ChatGPT）\n   → 執行「設定 API Key」可啟用 AI 分析\n") +
      "\n🔗 " + ppt.getUrl()
    );

  } catch (錯誤) {
    Logger.log("❌ " + 錯誤.message + "\n" + 錯誤.stack);
    SpreadsheetApp.getUi().alert("❌ " + 錯誤.message);
  }
}

// ============================================================
// 🧪  測試函式
// ============================================================

/**
 * 單獨測試 ChatGPT API 是否正常連線
 * 執行後查看 Logs（Ctrl+Enter）確認結果
 */
function 測試ChatGPT連線() {
  var ui = SpreadsheetApp.getUi();
  var apiKey = 取得API_Key_();

  if (!apiKey) {
    ui.alert("❌ 尚未設定 API Key！\n請先執行「設定 API Key」。");
    return;
  }

  ui.alert("⏳ 正在測試 ChatGPT 連線...");

  var 回應 = 呼叫ChatGPT_(
    "你是一位助理，請用繁體中文回應。",
    "請用一句話確認連線成功，並說明你使用的模型名稱。"
  );

  if (回應) {
    ui.alert("✅ ChatGPT 連線成功！\n\n回應：\n" + 回應);
  } else {
    ui.alert("❌ ChatGPT 連線失敗！\n請確認 API Key 是否正確，以及帳號是否有餘額。\n\n詳細錯誤請查看 Apps Script Logs。");
  }
}

// ============================================================
// 📦  初始化資料
// ============================================================

function 初始化月報資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("月報資料");
  if (!sheet) sheet = ss.insertSheet("月報資料"); else sheet.clear();

  var 標題 = [["部門", "營收", "成本", "新客戶", "完成專案"]];
  var 資料 = [
    ["業務部", 4800000, 2400000, 12, 8],
    ["行銷部", 2100000, 1200000, 8, 5],
    ["研發部", 0, 3500000, 0, 12],
    ["客服部", 950000, 600000, 5, 15],
    ["人資部", 0, 780000, 0, 3],
    ["財務部", 0, 420000, 0, 6]
  ];

  sheet.getRange(1, 1, 1, 5).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 5).setValues(資料);
  sheet.getRange("A1:E1").setBackground("#0a1628").setFontColor("#ffffff").setFontWeight("bold");
  sheet.getRange("B2:C7").setNumberFormat("#,##0");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 5; c++) sheet.autoResizeColumn(c);

  // 建立簡單柱狀圖
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("A1:C7"))
    .setPosition(9, 1, 0, 0)
    .setOption("title", "部門營收 vs 成本")
    .setOption("width", 600).setOption("height", 350)
    .setOption("colors", ["#1a73e8", "#ea4335"])
    .setOption("vAxis", { format: "#,##0" })
    .build();
  sheet.insertChart(chart);

  SpreadsheetApp.getUi().alert("✅ 月報資料已建立！\n\n下一步：\n1. 執行「設定 API Key」（可選）\n2. 執行「產生 AI 月報簡報」");
}

// ============================================================
// 📋  自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 AI 月報產生器")
    .addItem("📦 初始化月報資料",           "初始化月報資料")
    .addSeparator()
    .addItem("🔑 設定 OpenAI API Key",       "設定API_Key")
    .addItem("🧪 測試 ChatGPT 連線",         "測試ChatGPT連線")
    .addItem("🗑️ 清除 API Key",              "清除API_Key")
    .addSeparator()
    .addItem("🚀 產生 AI 月報簡報",          "產生AI月報簡報")
    .addToUi();
}
