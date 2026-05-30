// ============================================================
// Session 9：動態圖表生成
// 日期：115/05/30（六）09:00~12:00
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. Apps Script 介面複習
//   2. 建立與更新圖表（柱狀圖、折線圖、圓餅圖）
//   3. 組合圖表（柱狀圖 + 折線圖）
//   4. 儀表板（單一工作表多圖表排列）
//   5. 觸發器應用（靠近更新 → 圖表更新）
//   6. 實作：動態圖表生成
// ============================================================

// ============================================================
// 第一部分：建立柱狀圖
// ============================================================

/**
 * 建立部門業績柱狀圖
 */
function 建立柱狀圖() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("圖表資料");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化圖表資料」");
      return;
    }

    // 移除該表上既有的圖表（避免重複）
    var 舊圖表 = sheet.getCharts();
    for (var i = 0; i < 舊圖表.length; i++) {
      if (舊圖表[i].getOptions().get("title") === "部門季度業績") {
        sheet.removeChart(舊圖表[i]);
      }
    }

    // 資料範圍：A1:E7（部門 + Q1~Q4）
    var 資料範圍 = sheet.getRange("A1:E7");

    // 建立柱狀圖
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)        // 柱狀圖
      .addRange(資料範圍)                              // 資料來源
      .setPosition(9, 1, 0, 0)                        // 位置：第 9 列、第 1 欄
      .setOption("title", "部門季度業績")              // 標題
      .setOption("titleTextStyle", {
        fontSize: 16,
        bold: true,
        color: "#1a237e"
      })
      .setOption("width", 700)                         // 寬度
      .setOption("height", 400)                        // 高度
      .setOption("legend", { position: "bottom" })     // 圖例位置
      .setOption("hAxis", {
        title: "部門",
        titleTextStyle: { fontSize: 12 }
      })
      .setOption("vAxis", {
        title: "業績 (NT$)",
        titleTextStyle: { fontSize: 12 },
        format: "#,##0"
      })
      .setOption("colors", ["#1a73e8", "#34a853", "#fbbc04", "#ea4335"])  // 自訂顏色
      .setOption("isStacked", false)                   // 非堆疊
      .build();

    sheet.insertChart(chart);

    Logger.log("✅ 柱狀圖已建立！");
    SpreadsheetApp.getUi().alert("✅ 部門季度業績柱狀圖已生成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 建立堆疊柱狀圖
 */
function 建立堆疊柱狀圖() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("圖表資料");
    if (!sheet) return;

    var 資料範圍 = sheet.getRange("A1:E7");

    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(資料範圍)
      .setPosition(9, 6, 0, 0)
      .setOption("title", "部門季度業績（堆疊）")
      .setOption("titleTextStyle", { fontSize: 14, bold: true })
      .setOption("width", 600)
      .setOption("height", 400)
      .setOption("isStacked", true)                    // ⬅️ 堆疊模式
      .setOption("legend", { position: "bottom" })
      .setOption("colors", ["#1a73e8", "#34a853", "#fbbc04", "#ea4335"])
      .setOption("vAxis", { format: "#,##0" })
      .build();

    sheet.insertChart(chart);
    Logger.log("✅ 堆疊柱狀圖已建立！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第二部分：建立折線圖
// ============================================================

/**
 * 建立月度趨勢折線圖
 */
function 建立折線圖() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("月度趨勢");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化圖表資料」");
      return;
    }

    // 清除舊圖表
    sheet.getCharts().forEach(function(c) { sheet.removeChart(c); });

    var 資料範圍 = sheet.getRange("A1:D13"); // 月份 + 3 個系列

    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)             // 折線圖
      .addRange(資料範圍)
      .setPosition(15, 1, 0, 0)
      .setOption("title", "2026 年月度營收趨勢")
      .setOption("titleTextStyle", { fontSize: 16, bold: true, color: "#1a237e" })
      .setOption("width", 800)
      .setOption("height", 450)
      .setOption("curveType", "function")              // 平滑曲線
      .setOption("pointSize", 6)                        // 資料點大小
      .setOption("lineWidth", 3)                        // 線條寬度
      .setOption("legend", { position: "bottom" })
      .setOption("hAxis", {
        title: "月份",
        slantedText: true
      })
      .setOption("vAxis", {
        title: "金額 (NT$)",
        format: "#,##0",
        gridlines: { count: 6 }
      })
      .setOption("colors", ["#1a73e8", "#ea4335", "#34a853"])
      .build();

    sheet.insertChart(chart);

    Logger.log("✅ 折線圖已建立！");
    SpreadsheetApp.getUi().alert("✅ 月度趨勢折線圖已生成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第二點五部分：建立組合圖表（柱狀圖 + 折線圖）
// ============================================================

/**
 * 建立月度營收組合圖表
 * 說明：
 *   - 營收、成本 → 柱狀圖（BARS）
 *   - 利潤       → 折線圖（LINE），右側副 Y 軸
 * 使用「月度趨勢」工作表的 A1:D13 資料
 */
function 建立組合圖表() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("月度趨勢");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化圖表資料」");
      return;
    }

    // 移除同名舊圖表（避免重複建立）
    sheet.getCharts().forEach(function(c) {
      if (c.getOptions().get("title") === "月度營收 vs 利潤（組合圖）") {
        sheet.removeChart(c);
      }
    });

    // 資料範圍：A1:D13（月份 ＋ 營收 ＋ 成本 ＋ 利潤）
    var 資料範圍 = sheet.getRange("A1:D13");

    // ----------------------------------------------------------------
    // 組合圖表關鍵設定：
    //   setChartType → COMBO（組合類型）
    //   series.0 / series.1 → type: "bars"（柱狀）
    //   series.2            → type: "line"（折線）＋ targetAxisIndex: 1（右軸）
    // ----------------------------------------------------------------
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.COMBO)              // ⬅️ 組合圖表
      .addRange(資料範圍)
      .setPosition(30, 1, 0, 0)                           // 位置：第 30 列、第 1 欄
      .setOption("title", "月度營收 vs 利潤（組合圖）")
      .setOption("titleTextStyle", {
        fontSize: 16,
        bold: true,
        color: "#1a237e"
      })
      .setOption("width", 900)
      .setOption("height", 480)

      // --- 系列設定 ---
      // 系列 0：營收 → 柱狀
      .setOption("series", {
        0: { type: "bars",  color: "#1a73e8", targetAxisIndex: 0 },
        1: { type: "bars",  color: "#ea4335", targetAxisIndex: 0 },
        2: { type: "line",  color: "#34a853", targetAxisIndex: 1,
             lineWidth: 3, pointSize: 7, curveType: "function" }  // 平滑折線
      })

      // --- 軸設定 ---
      .setOption("hAxis", {
        title: "月份",
        slantedText: true,
        slantedTextAngle: 30
      })
      .setOption("vAxes", {
        0: {                                               // 左軸：營收 / 成本
          title: "金額 (NT$)",
          format: "#,##0",
          titleTextStyle: { color: "#1a73e8" },
          gridlines: { count: 6 }
        },
        1: {                                               // 右軸：利潤
          title: "利潤 (NT$)",
          format: "#,##0",
          titleTextStyle: { color: "#34a853" }
        }
      })

      // --- 其他樣式 ---
      .setOption("legend", { position: "bottom" })
      .setOption("bar", { groupWidth: "60%" })             // 柱寬比例
      .setOption("backgroundColor", { fill: "#fafafa" })  // 背景色
      .build();

    sheet.insertChart(chart);

    Logger.log("✅ 組合圖表已建立！");
    SpreadsheetApp.getUi().alert(
      "✅ 月度營收組合圖表已生成！\n" +
      "　🔵 藍柱 = 營收\n" +
      "　🔴 紅柱 = 成本\n" +
      "　🟢 綠線 = 利潤（右軸）"
    );

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第三部分：建立圓餅圖
// ============================================================

/**
 * 建立部門預算佔比圓餅圖
 */
function 建立圓餅圖() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("圖表資料");
    if (!sheet) return;

    // 計算年度總額（G 欄）
    var 資料 = sheet.getRange("A2:E7").getValues();
    var 圓餅資料 = [["部門", "年度總額"]];

    for (var i = 0; i < 資料.length; i++) {
      var 年度總額 = 資料[i][1] + 資料[i][2] + 資料[i][3] + 資料[i][4];
      圓餅資料.push([資料[i][0], 年度總額]);
    }

    // 寫入圓餅圖資料
    sheet.getRange("H1:I1").setValues([["部門", "年度總額"]]);
    sheet.getRange("H1:I1").setFontWeight("bold").setBackground("#e0e0e0");
    for (var j = 0; j < 圓餅資料.length - 1; j++) {
      sheet.getRange(j + 2, 8, 1, 2).setValues([圓餅資料[j + 1]]);
    }
    sheet.getRange("I2:I7").setNumberFormat("#,##0");

    // 建立圓餅圖
    var 圓餅範圍 = sheet.getRange("H1:I7");
    var chart = sheet.newChart()
      .setChartType(Charts.ChartType.PIE)              // 圓餅圖
      .addRange(圓餅範圍)
      .setPosition(25, 1, 0, 0)
      .setOption("title", "部門年度預算佔比")
      .setOption("titleTextStyle", { fontSize: 16, bold: true })
      .setOption("width", 600)
      .setOption("height", 400)
      .setOption("legend", { position: "right" })
      .setOption("pieHole", 0.4)                        // 環形圖效果（0~1）
      .setOption("colors", ["#1a73e8", "#34a853", "#fbbc04", "#ea4335", "#9c27b0", "#00bcd4"])
      .setOption("pieSliceText", "percentage")           // 顯示百分比
      .build();

    sheet.insertChart(chart);

    Logger.log("✅ 圓餅圖已建立！");
    SpreadsheetApp.getUi().alert("✅ 部門預算圓餅圖已生成！");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第四部分：動態更新圖表
// ============================================================

/**
 * 更新圖表資料
 * 說明：修改資料後，圖表會自動更新（因為圖表綁定了 Range）
 * 這裡示範如何手動觸發資料更新
 */
function 模擬資料更新() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("圖表資料");
    if (!sheet) return;

    // 模擬 Q4 業績更新（隨機增減 10%）
    for (var i = 2; i <= 7; i++) {
      var 原值 = sheet.getRange(i, 5).getValue(); // E 欄 = Q4
      var 變動 = 原值 * (0.9 + Math.random() * 0.2); // ±10%
      sheet.getRange(i, 5).setValue(Math.round(變動));
    }

    // 圖表會自動根據範圍資料更新！
    Logger.log("✅ Q4 業績資料已更新，圖表將自動刷新！");
    SpreadsheetApp.getUi().alert("✅ Q4 資料已更新！\n圖表會自動反映新數據。");

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 一鍵生成全部圖表
 */
function 一鍵生成圖表() {
  建立柱狀圖();
  建立堆疊柱狀圖();
  建立折線圖();
  建立圓餅圖();
  建立組合圖表();            // ⬅️ 新增組合圖表

  SpreadsheetApp.getUi().alert("✅ 全部圖表已生成！\n包含：柱狀圖、堆疊圖、折線圖、圓餅圖、組合圖");
}

// ============================================================
// 第五部分：儀表板
// ============================================================

/**
 * 建立儀表板
 * 說明：在「📊 儀表板」工作表中建立：
 *   ① 標題橫幅列
 *   ② KPI 摘要卡（總營收、平均利潤率、最高業績部門、Q4 成長率）
 *   ③ 2×2 排列的四個小型圖表
 *      左上：部門業績柱狀圖   右上：月度趨勢折線圖
 *      左下：預算佔比環形圖   右下：Q4 成長組合圖
 */
function 建立儀表板() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── 取得資料來源 ─────────────────────────────────────────
    var 部門sheet = ss.getSheetByName("圖表資料");
    var 月度sheet  = ss.getSheetByName("月度趨勢");
    if (!部門sheet || !月度sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化圖表資料」");
      return;
    }

    // ── 建立（或清空）儀表板工作表 ───────────────────────────
    var db = ss.getSheetByName("📊 儀表板");
    if (!db) {
      db = ss.insertSheet("📊 儀表板");
      // 移到最前面
      ss.setActiveSheet(db);
      ss.moveActiveSheet(1);
    } else {
      db.clear();
      db.getCharts().forEach(function(c) { db.removeChart(c); });
    }

    // ── ① 標題橫幅 ────────────────────────────────────────────
    db.setColumnWidth(1, 20);          // A 欄：左邊距
    for (var col = 2; col <= 13; col++) db.setColumnWidth(col, 80);
    db.setColumnWidth(14, 20);         // N 欄：右邊距

    db.setRowHeight(1, 8);             // 頂部留白
    db.setRowHeight(2, 50);            // 標題列
    db.setRowHeight(3, 8);             // 標題下間距

    var 標題儲存格 = db.getRange("B2:M2");
    標題儲存格.merge()
      .setValue("📊  2026 年度營運儀表板")
      .setBackground("#1a237e")
      .setFontColor("#ffffff")
      .setFontSize(20)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // ── ② KPI 摘要卡 ─────────────────────────────────────────
    // 計算 KPI 數值
    var 月度值 = 月度sheet.getRange("B2:D13").getValues();
    var 部門值 = 部門sheet.getRange("A2:E7").getValues();

    var 總營收 = 0, 總利潤 = 0;
    for (var r = 0; r < 月度值.length; r++) {
      總營收 += 月度值[r][0];
      總利潤 += 月度值[r][2];  // 利潤 = D 欄
    }
    var 平均利潤率 = Math.round((總利潤 / 總營收) * 100);

    var 最高業績部門 = "", 最高業績 = 0;
    for (var d = 0; d < 部門值.length; d++) {
      var 部門年合計 = 部門值[d][1] + 部門值[d][2] + 部門值[d][3] + 部門值[d][4];
      if (部門年合計 > 最高業績) { 最高業績 = 部門年合計; 最高業績部門 = 部門值[d][0]; }
    }

    // Q4 vs Q1 成長率（各部門平均）
    var Q1總 = 0, Q4總 = 0;
    for (var q = 0; q < 部門值.length; q++) {
      Q1總 += 部門值[q][1];
      Q4總 += 部門值[q][4];
    }
    var Q4成長率 = Math.round(((Q4總 - Q1總) / Q1總) * 100);

    // KPI 卡定義：[標題, 數值, 背景色, 文字色]
    var KPI卡 = [
      ["💰 年度總營收",  "NT$ " + (總營收 / 1000000).toFixed(1) + "M",  "#e3f2fd", "#0d47a1"],
      ["📈 平均利潤率",  平均利潤率 + " %",                              "#e8f5e9", "#1b5e20"],
      ["🏆 最高業績部門", 最高業績部門,                                   "#fff8e1", "#e65100"],
      ["🚀 Q4 成長率",   "+" + Q4成長率 + " %",                         "#fce4ec", "#880e4f"]
    ];

    db.setRowHeight(4, 18);  // KPI 標題列高
    db.setRowHeight(5, 36);  // KPI 數值列高
    db.setRowHeight(6, 12);  // KPI 下間距

    // 欄位對應：B~D, E~G, H~J, K~M（每卡佔 3 欄）
    var kpiCols = [[2,4],[5,7],[8,10],[11,13]];
    for (var k = 0; k < KPI卡.length; k++) {
      var sc = kpiCols[k][0], ec = kpiCols[k][1];
      var 標題格 = db.getRange(4, sc, 1, ec - sc + 1);
      var 數值格 = db.getRange(5, sc, 1, ec - sc + 1);

      標題格.merge()
        .setValue(KPI卡[k][0])
        .setBackground(KPI卡[k][2])
        .setFontColor(KPI卡[k][3])
        .setFontSize(10)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");

      數值格.merge()
        .setValue(KPI卡[k][1])
        .setBackground(KPI卡[k][2])
        .setFontColor(KPI卡[k][3])
        .setFontSize(18)
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
    }

    // ── ③ 圖表區（2×2 網格）────────────────────────────────
    //
    // 版面配置（小型圖表，各約 420×260px）：
    //   列 7~26  ： 上排圖表
    //   列 27~46 ： 下排圖表
    //   欄 B~G（左），欄 H~M（右）
    //
    // setPosition(row, col, offsetX, offsetY) 對應儲存格位置
    // ──────────────────────────────────────────────────────────

    var 小圖寬 = 430;
    var 小圖高 = 260;

    // ── 左上：部門季度業績（柱狀圖）───────────────────────────
    var chart1 = db.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(部門sheet.getRange("A1:E7"))
      .setPosition(7, 2, 0, 0)
      .setOption("title", "📊 部門季度業績")
      .setOption("titleTextStyle", { fontSize: 12, bold: true, color: "#1a237e" })
      .setOption("width",  小圖寬)
      .setOption("height", 小圖高)
      .setOption("legend", { position: "bottom", textStyle: { fontSize: 9 } })
      .setOption("hAxis",  { title: "", textStyle: { fontSize: 9 } })
      .setOption("vAxis",  { format: "#,##0", textStyle: { fontSize: 9 } })
      .setOption("colors", ["#1a73e8", "#34a853", "#fbbc04", "#ea4335"])
      .setOption("chartArea", { width: "80%", height: "65%" })
      .setOption("backgroundColor", { fill: "#f8f9fa" })
      .build();
    db.insertChart(chart1);

    // ── 右上：月度營收趨勢（折線圖）────────────────────────────
    var chart2 = db.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(月度sheet.getRange("A1:D13"))
      .setPosition(7, 8, 0, 0)
      .setOption("title", "📈 月度營收趨勢")
      .setOption("titleTextStyle", { fontSize: 12, bold: true, color: "#1a237e" })
      .setOption("width",  小圖寬)
      .setOption("height", 小圖高)
      .setOption("curveType", "function")
      .setOption("pointSize", 4)
      .setOption("lineWidth", 2)
      .setOption("legend", { position: "bottom", textStyle: { fontSize: 9 } })
      .setOption("hAxis",  { slantedText: true, slantedTextAngle: 45, textStyle: { fontSize: 8 } })
      .setOption("vAxis",  { format: "#,##0", textStyle: { fontSize: 9 } })
      .setOption("colors", ["#1a73e8", "#ea4335", "#34a853"])
      .setOption("chartArea", { width: "80%", height: "65%" })
      .setOption("backgroundColor", { fill: "#f8f9fa" })
      .build();
    db.insertChart(chart2);

    // ── 左下：部門年度預算佔比（環形圖）────────────────────────
    // 準備圓餅圖資料（從圖表資料工作表讀取）
    var 圓餅資料 = [["部門", "年度總額"]];
    for (var pd = 0; pd < 部門值.length; pd++) {
      var 年合 = 部門值[pd][1] + 部門值[pd][2] + 部門值[pd][3] + 部門值[pd][4];
      圓餅資料.push([部門值[pd][0], 年合]);
    }
    // 寫入儀表板工作表暫存區（P 欄起，不影響顯示）
    db.getRange(1, 16, 圓餅資料.length, 2).setValues(圓餅資料);
    db.getRange(1, 16, 1, 2).setFontWeight("bold");

    var chart3 = db.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(db.getRange(1, 16, 7, 2))
      .setPosition(27, 2, 0, 0)
      .setOption("title", "🥧 年度預算佔比")
      .setOption("titleTextStyle", { fontSize: 12, bold: true, color: "#1a237e" })
      .setOption("width",  小圖寬)
      .setOption("height", 小圖高)
      .setOption("pieHole", 0.45)
      .setOption("legend", { position: "right", textStyle: { fontSize: 9 } })
      .setOption("pieSliceText", "percentage")
      .setOption("colors", ["#1a73e8", "#34a853", "#fbbc04", "#ea4335", "#9c27b0", "#00bcd4"])
      .setOption("chartArea", { width: "70%", height: "75%" })
      .setOption("backgroundColor", { fill: "#f8f9fa" })
      .build();
    db.insertChart(chart3);

    // ── 右下：月度成本 vs 利潤（組合圖）───────────────────────
    var chart4 = db.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(月度sheet.getRange("A1:D13"))
      .setPosition(27, 8, 0, 0)
      .setOption("title", "📊📈 成本 vs 利潤（組合）")
      .setOption("titleTextStyle", { fontSize: 12, bold: true, color: "#1a237e" })
      .setOption("width",  小圖寬)
      .setOption("height", 小圖高)
      .setOption("series", {
        0: { type: "bars",  color: "#1a73e8", targetAxisIndex: 0 },
        1: { type: "bars",  color: "#ea4335", targetAxisIndex: 0 },
        2: { type: "line",  color: "#34a853", targetAxisIndex: 1,
             lineWidth: 2, pointSize: 4, curveType: "function" }
      })
      .setOption("vAxes", {
        0: { format: "#,##0", textStyle: { fontSize: 8 } },
        1: { format: "#,##0", textStyle: { fontSize: 8 } }
      })
      .setOption("hAxis",  { slantedText: true, slantedTextAngle: 45, textStyle: { fontSize: 8 } })
      .setOption("legend", { position: "bottom", textStyle: { fontSize: 9 } })
      .setOption("bar",    { groupWidth: "55%" })
      .setOption("chartArea", { width: "80%", height: "62%" })
      .setOption("backgroundColor", { fill: "#f8f9fa" })
      .build();
    db.insertChart(chart4);

    // ── ④ 儀表板美化：凍結標題列 / 隱藏格線 ─────────────────
    db.setFrozenRows(6);                     // 凍結 KPI 區域
    db.setHiddenGridlines(true);             // 隱藏格線，更像真正的儀表板
    db.getRange("A1:N50")
      .setBackground("#ffffff");

    // 隱藏暫存欄（P 欄之後）
    db.hideColumns(15, 5);

    // 跳到儀表板
    ss.setActiveSheet(db);

    Logger.log("✅ 儀表板已建立！");
    SpreadsheetApp.getUi().alert(
      "✅ 儀表板已建立完成！\n" +
      "　📍 工作表：📊 儀表板\n" +
      "　📌 左上：部門季度業績\n" +
      "　📌 右上：月度趨勢\n" +
      "　📌 左下：預算佔比\n" +
      "　📌 右下：成本 vs 利潤組合圖"
    );

  } catch (錯誤) {
    Logger.log("❌ 儀表板錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 清除儀表板（刪除工作表）
 */
function 清除儀表板() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var db = ss.getSheetByName("📊 儀表板");
  if (db) {
    ss.deleteSheet(db);
    SpreadsheetApp.getUi().alert("🗑️ 儀表板已清除！");
  } else {
    SpreadsheetApp.getUi().alert("⚠️ 找不到儀表板工作表。");
  }
}

// ============================================================
// 初始化範例資料
// ============================================================

function 初始化圖表資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 圖表資料（部門季度）---
  var sheet1 = ss.getSheetByName("圖表資料");
  if (!sheet1) sheet1 = ss.insertSheet("圖表資料"); else sheet1.clear();

  var 標題1 = [["部門", "Q1", "Q2", "Q3", "Q4"]];
  var 資料1 = [
    ["業務部", 1500000, 1800000, 2000000, 2200000],
    ["行銷部", 800000, 900000, 1000000, 1100000],
    ["研發部", 2500000, 2600000, 2700000, 2800000],
    ["人資部", 500000, 520000, 540000, 560000],
    ["財務部", 400000, 420000, 430000, 450000],
    ["客服部", 600000, 650000, 700000, 750000]
  ];

  sheet1.getRange(1, 1, 1, 5).setValues(標題1);
  sheet1.getRange(2, 1, 資料1.length, 5).setValues(資料1);
  sheet1.getRange("A1:E1").setBackground("#1a237e").setFontColor("#fff").setFontWeight("bold");
  sheet1.getRange("B2:E7").setNumberFormat("#,##0");
  sheet1.setFrozenRows(1);
  for (var c = 1; c <= 5; c++) sheet1.autoResizeColumn(c);

  // --- 月度趨勢資料 ---
  var sheet2 = ss.getSheetByName("月度趨勢");
  if (!sheet2) sheet2 = ss.insertSheet("月度趨勢"); else sheet2.clear();

  var 標題2 = [["月份", "營收", "成本", "利潤"]];
  var 月度資料 = [];
  var 基礎營收 = 3000000;
  for (var m = 1; m <= 12; m++) {
    var 營收 = 基礎營收 + Math.floor(Math.random() * 1000000) + (m * 100000);
    var 成本 = Math.floor(營收 * (0.55 + Math.random() * 0.15));
    var 利潤 = 營收 - 成本;
    月度資料.push([m + "月", 營收, 成本, 利潤]);
  }

  sheet2.getRange(1, 1, 1, 4).setValues(標題2);
  sheet2.getRange(2, 1, 月度資料.length, 4).setValues(月度資料);
  sheet2.getRange("A1:D1").setBackground("#e65100").setFontColor("#fff").setFontWeight("bold");
  sheet2.getRange("B2:D13").setNumberFormat("#,##0");
  sheet2.setFrozenRows(1);
  for (var c2 = 1; c2 <= 4; c2++) sheet2.autoResizeColumn(c2);

  SpreadsheetApp.getUi().alert("✅ 圖表資料已建立！\n（圖表資料 + 月度趨勢）");
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📚 Session 9 工具")
    .addItem("📦 初始化圖表資料", "初始化圖表資料")
    .addSeparator()
    // ── 個別圖表 ──
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("📈 個別圖表")
        .addItem("📊 建立柱狀圖",        "建立柱狀圖")
        .addItem("📊 建立堆疊柱狀圖",    "建立堆疊柱狀圖")
        .addItem("📈 建立折線圖",        "建立折線圖")
        .addItem("🥧 建立圓餅圖",        "建立圓餅圖")
        .addItem("📊📈 建立組合圖表",    "建立組合圖表")
        .addSeparator()
        .addItem("🚀 一鍵生成全部圖表",  "一鍵生成圖表")
    )
    .addSeparator()
    // ── 儀表板 ──
    .addItem("🖥️ 建立儀表板",   "建立儀表板")
    .addItem("🗑️ 清除儀表板",   "清除儀表板")
    .addSeparator()
    .addItem("🔄 模擬資料更新", "模擬資料更新")
    .addToUi();
}
