// ============================================================
// 進階練習：自動化營運 Dashboard
// 對應：Session 9（柱狀圖、折線圖、圓餅圖）
// ============================================================

// ============================================================
// 輔助函式：在 Dashboard 工作表插入顏色圖例說明列
// ============================================================

/**
 * 在指定列繪製一排「色塊 + 文字」圖例
 * @param {Sheet}    sheet    - 目標工作表
 * @param {number}   row      - 起始列號（1-indexed）
 * @param {number}   startCol - 起始欄號（1-indexed）
 * @param {string}   標題     - 圖例區塊標題（寫在最前一格）
 * @param {Array}    項目     - [{色碼, 名稱}, ...]
 */
function 寫入圖例列(sheet, row, startCol, 標題, 項目) {
  sheet.setRowHeight(row, 22);

  // 標題格（灰底粗體）
  sheet.getRange(row, startCol)
    .setValue(標題)
    .setBackground("#eeeeee")
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("right")
    .setVerticalAlignment("middle");

  // 每個圖例：色塊格 + 文字格交錯排列
  var col = startCol + 1;
  for (var i = 0; i < 項目.length; i++) {
    // 色塊格
    sheet.getRange(row, col)
      .setValue("  ")
      .setBackground(項目[i].色碼);
    // 文字格
    sheet.getRange(row, col + 1)
      .setValue(項目[i].名稱)
      .setFontSize(9)
      .setVerticalAlignment("middle")
      .setBackground("#ffffff");
    col += 2;
  }
}

// ============================================================
// 主函式：一鍵建立完整營運儀表板（多種圖表組合）
// ============================================================

/**
 * 建立 Dashboard 工作表，包含 4 個圖表及各自的顏色圖例說明列
 */
function 建立營運Dashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 資料表 = ss.getSheetByName("營運資料");
    if (!資料表) { SpreadsheetApp.getUi().alert("❌ 請先初始化"); return; }

    // 建立（或清空）Dashboard 工作表
    var sheet = ss.getSheetByName("Dashboard");
    if (sheet) {
      sheet.getCharts().forEach(function(c) { sheet.removeChart(c); });
      sheet.clear();
    } else {
      sheet = ss.insertSheet("Dashboard", 0);
    }

    // ── 標題橫幅 ─────────────────────────────────────────────
    sheet.getRange("A1:L1").merge();
    sheet.getRange("A1")
      .setValue("📊 2026 年營運 Dashboard")
      .setFontSize(22).setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#0d47a1").setFontColor("#fff");
    sheet.setRowHeight(1, 55);

    // ── 圖表 1：月度營收折線圖（位置：列3, 欄A） ─────────────
    //   顏色：藍=營收、紅=成本、綠=利潤
    var 月度範圍 = 資料表.getRange("A1:D13");
    var 折線圖 = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(月度範圍)
      .setPosition(3, 1, 0, 0)
      .setOption("title", "📈 月度營收 vs 成本 vs 利潤趨勢")
      .setOption("titleTextStyle", { fontSize: 14, bold: true })
      .setOption("width", 600).setOption("height", 350)
      .setOption("curveType", "function")
      .setOption("pointSize", 5).setOption("lineWidth", 3)
      .setOption("colors", ["#1a73e8", "#ea4335", "#34a853"])
      .setOption("legend", { position: "bottom" })          // ✅ 圖表內圖例
      .setOption("vAxis", { format: "#,##0" })
      .build();
    sheet.insertChart(折線圖);

    // 折線圖圖例說明列（圖表下方約第 20 列）
    寫入圖例列(sheet, 20, 1, "📈 圖例：", [
      { 色碼: "#1a73e8", 名稱: "營收" },
      { 色碼: "#ea4335", 名稱: "成本" },
      { 色碼: "#34a853", 名稱: "利潤" }
    ]);

    // ── 圖表 2：部門目標 vs 實績柱狀圖（位置：列3, 欄G） ─────
    //   顏色：淺藍=目標、深藍=實績
    var 部門表 = ss.getSheetByName("部門績效");
    if (部門表) {
      var 部門範圍 = 部門表.getRange("A1:C7");
      var 柱狀圖 = sheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(部門範圍)
        .setPosition(3, 7, 0, 0)
        .setOption("title", "📊 部門目標 vs 實績")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("width", 500).setOption("height", 350)
        .setOption("colors", ["#90caf9", "#1565c0"])
        .setOption("legend", { position: "bottom" })        // ✅ 圖表內圖例
        .setOption("vAxis", { format: "#,##0" })
        .build();
      sheet.insertChart(柱狀圖);

      // 柱狀圖圖例說明列
      寫入圖例列(sheet, 20, 7, "📊 圖例：", [
        { 色碼: "#90caf9", 名稱: "目標" },
        { 色碼: "#1565c0", 名稱: "實績" }
      ]);
    }

    // ── 圖表 3：客戶產業分佈圓餅圖（位置：列23, 欄A） ────────
    //   顏色：各產業對應色碼
    var 產業表 = ss.getSheetByName("客戶產業");
    if (產業表) {
      var 產業資料 = 產業表.getRange("A2:A6").getValues(); // 讀取產業名稱
      var 產業範圍 = 產業表.getRange("A1:B6");
      var 圓餅圖 = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(產業範圍)
        .setPosition(23, 1, 0, 0)
        .setOption("title", "🥧 客戶產業分佈")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("width", 500).setOption("height", 350)
        .setOption("pieHole", 0.4)
        .setOption("colors", ["#1a73e8", "#34a853", "#fbbc04", "#ea4335", "#9c27b0"])
        .setOption("pieSliceText", "percentage")
        .setOption("legend", { position: "right" })         // ✅ 圖表內圖例（右側）
        .build();
      sheet.insertChart(圓餅圖);

      // 圓餅圖圖例說明列（以實際資料名稱對應色碼）
      var 色碼清單 = ["#1a73e8", "#34a853", "#fbbc04", "#ea4335", "#9c27b0"];
      var 產業項目 = [];
      for (var p = 0; p < 產業資料.length; p++) {
        產業項目.push({ 色碼: 色碼清單[p], 名稱: 產業資料[p][0] });
      }
      寫入圖例列(sheet, 38, 1, "🥧 圖例：", 產業項目);
    }

    // ── 圖表 4：客戶數量面積圖（位置：列23, 欄G） ────────────
    //   顏色：藍=新客戶、綠=續約客戶
    var 客戶趨勢 = ss.getSheetByName("客戶趨勢");
    if (客戶趨勢) {
      var 趨勢範圍 = 客戶趨勢.getRange("A1:C13");
      var 面積圖 = sheet.newChart()
        .setChartType(Charts.ChartType.AREA)
        .addRange(趨勢範圍)
        .setPosition(23, 7, 0, 0)
        .setOption("title", "📈 客戶數量趨勢")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("width", 500).setOption("height", 350)
        .setOption("colors", ["#1a73e8", "#34a853"])
        .setOption("isStacked", true)
        .setOption("legend", { position: "bottom" })        // ✅ 圖表內圖例
        .build();
      sheet.insertChart(面積圖);

      // 面積圖圖例說明列
      寫入圖例列(sheet, 38, 7, "📈 圖例：", [
        { 色碼: "#1a73e8", 名稱: "新客戶" },
        { 色碼: "#34a853", 名稱: "續約客戶" }
      ]);
    }

    // ── 隱藏格線，更像儀表板 ─────────────────────────────────
    sheet.setHiddenGridlines(true);

    SpreadsheetApp.getUi().alert(
      "✅ 營運 Dashboard 已建立！共 4 個圖表。\n" +
      "　每個圖表下方均有顏色圖例說明列。"
    );

  } catch (錯誤) {
    Logger.log("❌ " + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ " + 錯誤.message);
  }
}

// ============================================================
// 初始化範例資料（4 個工作表）
// ============================================================

function 初始化營運資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── 營運資料（月度）──────────────────────────────────────
  var s1 = ss.getSheetByName("營運資料");
  if (!s1) s1 = ss.insertSheet("營運資料"); else s1.clear();
  s1.getRange(1, 1, 1, 4).setValues([["月份", "營收", "成本", "利潤"]]);
  var 月度 = [];
  for (var m = 1; m <= 12; m++) {
    var 營收 = 3000000 + m * 200000 + Math.floor(Math.random() * 500000);
    var 成本 = Math.floor(營收 * (0.55 + Math.random() * 0.1));
    月度.push([m + "月", 營收, 成本, 營收 - 成本]);
  }
  s1.getRange(2, 1, 12, 4).setValues(月度);
  s1.getRange("A1:D1").setBackground("#1565c0").setFontColor("#fff").setFontWeight("bold");
  s1.getRange("B2:D13").setNumberFormat("#,##0");
  // 圖例備註（A15 起）
  s1.getRange("A15").setValue("【顏色說明】").setFontWeight("bold");
  s1.getRange("A16").setBackground("#1a73e8").setValue("  ").setNote("藍色 = 營收");
  s1.getRange("B16").setValue("營收");
  s1.getRange("A17").setBackground("#ea4335").setValue("  ").setNote("紅色 = 成本");
  s1.getRange("B17").setValue("成本");
  s1.getRange("A18").setBackground("#34a853").setValue("  ").setNote("綠色 = 利潤");
  s1.getRange("B18").setValue("利潤");

  // ── 部門績效 ──────────────────────────────────────────────
  var s2 = ss.getSheetByName("部門績效");
  if (!s2) s2 = ss.insertSheet("部門績效"); else s2.clear();
  s2.getRange(1, 1, 1, 3).setValues([["部門", "目標", "實績"]]);
  s2.getRange(2, 1, 6, 3).setValues([
    ["業務部", 5000000, 4800000], ["行銷部", 2000000, 2100000],
    ["研發部", 3000000, 2900000], ["客服部", 1000000, 950000],
    ["人資部", 800000, 780000],   ["財務部", 600000, 620000]
  ]);
  s2.getRange("A1:C1").setBackground("#2e7d32").setFontColor("#fff").setFontWeight("bold");
  s2.getRange("B2:C7").setNumberFormat("#,##0");
  // 圖例備註
  s2.getRange("A9").setValue("【顏色說明】").setFontWeight("bold");
  s2.getRange("A10").setBackground("#90caf9").setValue("  ");
  s2.getRange("B10").setValue("目標（淺藍）");
  s2.getRange("A11").setBackground("#1565c0").setValue("  ");
  s2.getRange("B11").setValue("實績（深藍）");

  // ── 客戶產業 ──────────────────────────────────────────────
  var s3 = ss.getSheetByName("客戶產業");
  if (!s3) s3 = ss.insertSheet("客戶產業"); else s3.clear();
  s3.getRange(1, 1, 1, 2).setValues([["產業", "客戶數"]]);
  var 產業清單 = [
    ["科技業", 45], ["製造業", 30], ["服務業", 25], ["金融業", 18], ["其他", 12]
  ];
  s3.getRange(2, 1, 5, 2).setValues(產業清單);
  s3.getRange("A1:B1").setBackground("#6a1b9a").setFontColor("#fff").setFontWeight("bold");
  // 圖例備註（依色碼順序對應產業）
  var 產業色碼 = ["#1a73e8", "#34a853", "#fbbc04", "#ea4335", "#9c27b0"];
  s3.getRange("A8").setValue("【顏色說明】").setFontWeight("bold");
  for (var pi = 0; pi < 產業清單.length; pi++) {
    s3.getRange(9 + pi, 1).setBackground(產業色碼[pi]).setValue("  ");
    s3.getRange(9 + pi, 2).setValue(產業清單[pi][0]);
  }

  // ── 客戶趨勢 ──────────────────────────────────────────────
  var s4 = ss.getSheetByName("客戶趨勢");
  if (!s4) s4 = ss.insertSheet("客戶趨勢"); else s4.clear();
  s4.getRange(1, 1, 1, 3).setValues([["月份", "新客戶", "續約客戶"]]);
  var 趨勢 = [];
  for (var n = 1; n <= 12; n++) {
    趨勢.push([n + "月", Math.floor(Math.random() * 10) + 5, Math.floor(Math.random() * 15) + 20]);
  }
  s4.getRange(2, 1, 12, 3).setValues(趨勢);
  s4.getRange("A1:C1").setBackground("#e65100").setFontColor("#fff").setFontWeight("bold");
  // 圖例備註
  s4.getRange("A15").setValue("【顏色說明】").setFontWeight("bold");
  s4.getRange("A16").setBackground("#1a73e8").setValue("  ");
  s4.getRange("B16").setValue("新客戶（藍）");
  s4.getRange("A17").setBackground("#34a853").setValue("  ");
  s4.getRange("B17").setValue("續約客戶（綠）");

  SpreadsheetApp.getUi().alert(
    "✅ 營運資料已建立！（4 個資料表）\n" +
    "　每張資料表下方均有【顏色說明】區塊。\n" +
    "　請執行「建立營運 Dashboard」。"
  );
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 營運 Dashboard")
    .addItem("📦 初始化營運資料", "初始化營運資料")
    .addItem("📊 建立營運 Dashboard", "建立營運Dashboard")
    .addToUi();
}
