// ============================================================
// Session 10：Google Slides 整合專案
// 日期：115/05/30（六）13:30~16:30
// 講師：林冠廷
// ============================================================
// 本課程涵蓋：
//   1. Google Slides 整合（簡報建立、插入圖表與文字）
//   2. 執行限制（跨服務操作限制）
//   3. 專案實作：一鍵產生含圖表的報表簡報
//   4. 延伸練習 1：PDF 匯出並存到 Google Drive
//   5. 延伸練習 2：Email 通知主管（含簡報連結）
//   6. 延伸練習 3：模板套用（填入動態資料）
// ============================================================

// ============================================================
// ⚙️  全域設定區（請依需求修改）
// ============================================================

var CONFIG = {
  // 延伸練習 2：主管 Email（可填多個，用逗號分隔）
  主管Email: "supervisor@example.com",
  // 延伸練習 3：報告模板簡報的 ID（建立模板後填入）
  //   → 開啟模板簡報，URL 中 /d/xxxxx/edit 的 xxxxx 即為 ID
  //   → 若此 ID 為空，程式會先自動建立一份模板
  模板簡報ID: "",
  // PDF 存放資料夾名稱（會在 Drive 根目錄建立）
  Drive資料夾名: "GAS 自動報告",
};

// ============================================================
// 第一部分：Google Slides 基本操作
// ============================================================

/**
 * 建立新的 Google Slides 簡報
 * 說明：示範如何用程式碼建立和操作簡報
 */
function 建立基本簡報() {
  try {
    // 建立新簡報
    var presentation = SlidesApp.create("GAS 自動產生簡報 - " +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd"));

    Logger.log("簡報已建立！");
    Logger.log("簡報 ID：" + presentation.getId());
    Logger.log("簡報 URL：" + presentation.getUrl());

    // --- 編輯第一張投影片（標題頁）---
    var 標題頁 = presentation.getSlides()[0];

    // 取得標題佔位符
    var 標題 = 標題頁.getPlaceholder(SlidesApp.PlaceholderType.CENTER_TITLE);
    if (標題) {
      標題.asShape().getText().setText("📊 2026 年度營運報告");
      var 文字樣式 = 標題.asShape().getText().getTextStyle();
      文字樣式.setFontSize(36);
      文字樣式.setBold(true);
      文字樣式.setForegroundColor("#1a237e");
    }

    // 取得副標題佔位符
    var 副標題 = 標題頁.getPlaceholder(SlidesApp.PlaceholderType.SUBTITLE);
    if (副標題) {
      副標題.asShape().getText().setText(
        "自動產生日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月 dd 日") +
        "\n由 Google Apps Script 自動生成"
      );
    }

    // --- 設定背景色 ---
    標題頁.getBackground().setSolidFill("#e8eaf6");

    Logger.log("✅ 標題頁已設定");

    // 開啟簡報
    SpreadsheetApp.getUi().alert(
      "✅ 簡報已建立！\n\n" +
      "點選連結開啟：\n" + presentation.getUrl()
    );

    return presentation;

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第二部分：在簡報中插入內容
// ============================================================

/**
 * 建立含有多頁內容的簡報
 */
function 建立多頁簡報() {
  try {
    var presentation = SlidesApp.create("多頁簡報範例 - " +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmm"));

    // === 第 1 頁：標題頁 ===
    var 第1頁 = presentation.getSlides()[0];
    第1頁.getBackground().setSolidFill("#1a237e");

    // 清除預設佈局，改用自訂元素
    第1頁.getPlaceholders().forEach(function(p) { p.remove(); });

    // 加入標題文字方塊
    var 標題框 = 第1頁.insertTextBox("📊 季度營運分析報告", 50, 150, 620, 80);
    var 標題樣式 = 標題框.getText().getTextStyle();
    標題樣式.setFontSize(32).setBold(true).setForegroundColor("#ffffff");
    標題框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // 加入日期
    var 日期框 = 第1頁.insertTextBox(
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月 dd 日"),
      50, 250, 620, 40
    );
    日期框.getText().getTextStyle().setFontSize(18).setForegroundColor("#b0bec5");
    日期框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // === 第 2 頁：摘要數據 ===
    var 第2頁 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    第2頁.getBackground().setSolidFill("#ffffff");

    // 頁面標題
    var 頁標題 = 第2頁.insertTextBox("📈 關鍵績效指標", 40, 20, 640, 50);
    頁標題.getText().getTextStyle().setFontSize(28).setBold(true).setForegroundColor("#1a237e");

    // 從試算表讀取資料
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("圖表資料");

    if (sheet) {
      var 資料 = sheet.getDataRange().getValues();

      // 建立表格
      var 表格 = 第2頁.insertTable(資料.length, 資料[0].length > 5 ? 5 : 資料[0].length);

      for (var i = 0; i < Math.min(資料.length, 7); i++) {
        for (var j = 0; j < Math.min(資料[0].length, 5); j++) {
          var 儲存格 = 表格.getCell(i, j);
          var 值 = 資料[i][j];

          // 格式化數字
          if (typeof 值 === "number" && i > 0) {
            值 = "NT$ " + 值.toLocaleString();
          }

          儲存格.getText().setText(String(值));
          儲存格.getText().getTextStyle().setFontSize(11);

          // 標題列格式
          if (i === 0) {
            儲存格.getFill().setSolidFill("#1a237e");
            儲存格.getText().getTextStyle()
              .setFontSize(12)
              .setBold(true)
              .setForegroundColor("#ffffff");
          } else if (i % 2 === 0) {
            儲存格.getFill().setSolidFill("#e8eaf6");
          }
        }
      }
    }

    // === 第 3 頁：重點摘要 ===
    var 第3頁 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    第3頁.getBackground().setSolidFill("#f5f5f5");

    var 摘要標題 = 第3頁.insertTextBox("💡 本季重點摘要", 40, 20, 640, 50);
    摘要標題.getText().getTextStyle().setFontSize(28).setBold(true).setForegroundColor("#1a237e");

    var 摘要內容 = 第3頁.insertTextBox(
      "✅ 研發部持續領先，Q4 預算達 280 萬\n\n" +
      "📈 業務部穩定成長，年增 46.7%\n\n" +
      "⚠️ 財務部預算最低，需評估人力配置\n\n" +
      "🎯 下季目標：整體營收成長 15%",
      60, 90, 600, 300
    );
    摘要內容.getText().getTextStyle().setFontSize(18).setForegroundColor("#37474f");
    摘要內容.getText().getListStyle().applyListPreset(SlidesApp.ListPreset.DISC_CIRCLE_SQUARE);

    // === 第 4 頁：結尾頁 ===
    var 第4頁 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    第4頁.getBackground().setSolidFill("#1a237e");

    var 結尾框 = 第4頁.insertTextBox("謝謝聆聽", 50, 180, 620, 60);
    結尾框.getText().getTextStyle().setFontSize(40).setBold(true).setForegroundColor("#ffffff");
    結尾框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var 聯絡框 = 第4頁.insertTextBox("📧 contact@company.com", 50, 260, 620, 40);
    聯絡框.getText().getTextStyle().setFontSize(16).setForegroundColor("#90caf9");
    聯絡框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    Logger.log("✅ 多頁簡報已建立！共 " + presentation.getSlides().length + " 頁");
    SpreadsheetApp.getUi().alert(
      "✅ 多頁簡報已建立！\n\n" +
      "共 " + presentation.getSlides().length + " 頁\n" +
      "點選連結開啟：\n" + presentation.getUrl()
    );

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第三部分：圖表插入簡報
// ============================================================

/**
 * 將試算表中的圖表插入到簡報
 */
function 插入圖表到簡報() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("圖表資料");
    if (!sheet) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化圖表資料」並建立圖表");
      return;
    }

    // 取得所有圖表
    var 圖表列表 = sheet.getCharts();
    if (圖表列表.length === 0) {
      SpreadsheetApp.getUi().alert("❌ 找不到圖表！\n請先在 Session 9 中建立圖表。");
      return;
    }

    // 建立簡報
    var presentation = SlidesApp.create("含圖表簡報 - " +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmm"));

    // 標題頁
    var 標題頁 = presentation.getSlides()[0];
    標題頁.getBackground().setSolidFill("#e8eaf6");
    var placeholders = 標題頁.getPlaceholders();
    if (placeholders.length > 0) {
      placeholders[0].asShape().getText().setText("📊 圖表分析報告");
      placeholders[0].asShape().getText().getTextStyle()
        .setFontSize(32).setBold(true).setForegroundColor("#1a237e");
    }

    // 為每個圖表建立一頁
    for (var i = 0; i < 圖表列表.length; i++) {
      var 新頁 = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

      // 從試算表圖表取得圖片
      var 圖表圖片 = 圖表列表[i].getBlob();

      // 插入圖表圖片到簡報
      var image = 新頁.insertImage(圖表圖片);

      // 調整位置和大小
      image.setLeft(40);
      image.setTop(60);
      image.setWidth(640);
      image.setHeight(380);

      // 加入頁面標題
      var 圖表標題 = 圖表列表[i].getOptions().get("title") || "圖表 " + (i + 1);
      var 標題框 = 新頁.insertTextBox("📊 " + 圖表標題, 40, 10, 640, 45);
      標題框.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#1a237e");
    }

    Logger.log("✅ 含圖表簡報已建立！共插入 " + 圖表列表.length + " 個圖表");
    SpreadsheetApp.getUi().alert(
      "✅ 含圖表簡報已建立！\n\n" +
      "共 " + (圖表列表.length + 1) + " 頁（包含 " + 圖表列表.length + " 個圖表）\n\n" +
      "點選連結開啟：\n" + presentation.getUrl()
    );

  } catch (錯誤) {
    Logger.log("❌ 錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 第四部分：🔧 專案實作 — 一鍵產生含圖表的報表簡報
// ============================================================

/**
 * 🚀 一鍵產生完整報表簡報
 * 說明：整合所有學過的技能
 *   1. 讀取試算表資料
 *   2. 計算統計資訊
 *   3. 建立圖表
 *   4. 產生簡報（含表格、圖表、摘要）
 */
function 一鍵產生報表簡報() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 開始時間 = new Date();

    SpreadsheetApp.getUi().alert("⏳ 開始生成報表簡報...\n\n這可能需要 10~30 秒，請稍候。");

    // ========== Step 1：讀取資料 ==========
    Logger.log("Step 1：讀取資料...");

    var 部門表 = ss.getSheetByName("圖表資料");
    var 趨勢表 = ss.getSheetByName("月度趨勢");

    if (!部門表 || !趨勢表) {
      SpreadsheetApp.getUi().alert("❌ 請先執行「初始化圖表資料」！");
      return;
    }

    var 部門資料 = 部門表.getDataRange().getValues();
    var 趨勢資料 = 趨勢表.getDataRange().getValues();

    // ========== Step 2：統計計算 ==========
    Logger.log("Step 2：計算統計...");

    var 部門統計 = [];
    var 年度總營收 = 0;
    for (var i = 1; i < 部門資料.length; i++) {
      var 年度合計 = 部門資料[i][1] + 部門資料[i][2] + 部門資料[i][3] + 部門資料[i][4];
      年度總營收 += 年度合計;
      部門統計.push({
        部門: 部門資料[i][0],
        Q1: 部門資料[i][1],
        Q4: 部門資料[i][4],
        年度: 年度合計,
        成長率: ((部門資料[i][4] - 部門資料[i][1]) / 部門資料[i][1] * 100).toFixed(1)
      });
    }

    // 排序：年度業績由高到低
    部門統計.sort(function(a, b) { return b.年度 - a.年度; });

    var 月度統計 = [];
    var 最高月營收 = 0, 最高月 = "";
    for (var j = 1; j < 趨勢資料.length; j++) {
      var 營收 = 趨勢資料[j][1];
      if (營收 > 最高月營收) {
        最高月營收 = 營收;
        最高月 = 趨勢資料[j][0];
      }
    }

    // ========== Step 3：確保圖表存在 ==========
    Logger.log("Step 3：檢查圖表...");

    var 圖表列表 = 部門表.getCharts();
    var 趨勢圖表列表 = 趨勢表.getCharts();

    // ========== Step 4：建立簡報 ==========
    Logger.log("Step 4：建立簡報...");

    var 簡報名 = "📊 營運分析報告 - " +
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy.MM.dd");
    var ppt = SlidesApp.create(簡報名);

    // --- 第 1 頁：封面 ---
    var 封面 = ppt.getSlides()[0];
    封面.getBackground().setSolidFill("#0d47a1");
    封面.getPlaceholders().forEach(function(p) { p.remove(); });

    var t1 = 封面.insertTextBox("📊 2026 年度營運分析報告", 40, 130, 640, 80);
    t1.getText().getTextStyle().setFontSize(36).setBold(true).setForegroundColor("#ffffff");
    t1.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var t2 = 封面.insertTextBox(
      Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月 dd 日") +
      " ｜ 自動產生報告",
      40, 230, 640, 40
    );
    t2.getText().getTextStyle().setFontSize(16).setForegroundColor("#90caf9");
    t2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // 加入裝飾線
    var 線 = 封面.insertLine(
      SlidesApp.LineCategory.STRAIGHT, 200, 220, 520, 220
    );
    線.getLineFill().setSolidFill("#64b5f6");
    線.setWeight(2);

    // --- 第 2 頁：KPI 摘要 ---
    var KPI頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    KPI頁.getBackground().setSolidFill("#fafafa");

    var kpiTitle = KPI頁.insertTextBox("🎯 關鍵績效指標 (KPI)", 40, 15, 640, 45);
    kpiTitle.getText().getTextStyle().setFontSize(26).setBold(true).setForegroundColor("#1a237e");

    // KPI 卡片
    var kpiData = [
      { 標籤: "年度總營收", 值: "NT$ " + (年度總營收 / 10000).toFixed(0) + " 萬", 色: "#1a73e8" },
      { 標籤: "部門數", 值: 部門統計.length + " 個", 色: "#34a853" },
      { 標籤: "最高月營收", 值: 最高月, 色: "#fbbc04" },
      { 標籤: "營收冠軍", 值: 部門統計[0].部門, 色: "#ea4335" }
    ];

    for (var k = 0; k < kpiData.length; k++) {
      var x = 40 + (k * 165);
      // 卡片背景
      var card = KPI頁.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, 80, 150, 120);
      card.getFill().setSolidFill(kpiData[k].色);
      card.getBorder().setWeight(0);

      // 卡片文字
      var 值框 = KPI頁.insertTextBox(kpiData[k].值, x + 10, 100, 130, 40);
      值框.getText().getTextStyle().setFontSize(18).setBold(true).setForegroundColor("#ffffff");
      值框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      var 標籤框 = KPI頁.insertTextBox(kpiData[k].標籤, x + 10, 150, 130, 30);
      標籤框.getText().getTextStyle().setFontSize(11).setForegroundColor("#e0e0e0");
      標籤框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }

    // 部門排行表格
    var 排行標題 = KPI頁.insertTextBox("🏆 部門年度業績排行", 40, 230, 640, 35);
    排行標題.getText().getTextStyle().setFontSize(18).setBold(true).setForegroundColor("#1a237e");

    var 表格列數 = Math.min(部門統計.length + 1, 7);
    var table = KPI頁.insertTable(表格列數, 4);

    // 表格標題
    var thData = ["排名", "部門", "年度業績", "Q1→Q4 成長率"];
    for (var th = 0; th < 4; th++) {
      table.getCell(0, th).getText().setText(thData[th]);
      table.getCell(0, th).getText().getTextStyle().setFontSize(10).setBold(true).setForegroundColor("#ffffff");
      table.getCell(0, th).getFill().setSolidFill("#1a237e");
    }

    for (var r = 0; r < Math.min(部門統計.length, 6); r++) {
      var d = 部門統計[r];
      var rowData = [String(r + 1), d.部門, "NT$ " + (d.年度 / 10000).toFixed(0) + " 萬", d.成長率 + "%"];
      for (var c = 0; c < 4; c++) {
        table.getCell(r + 1, c).getText().setText(rowData[c]);
        table.getCell(r + 1, c).getText().getTextStyle().setFontSize(10);
        if (r % 2 === 1) {
          table.getCell(r + 1, c).getFill().setSolidFill("#e8eaf6");
        }
      }
    }

    // --- 第 3-N 頁：圖表頁 ---
    if (圖表列表.length > 0) {
      for (var ci = 0; ci < 圖表列表.length; ci++) {
        var 圖表頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        圖表頁.getBackground().setSolidFill("#ffffff");

        var chartTitle = 圖表列表[ci].getOptions().get("title") || "圖表 " + (ci + 1);
        var ct = 圖表頁.insertTextBox("📊 " + chartTitle, 40, 10, 640, 40);
        ct.getText().getTextStyle().setFontSize(24).setBold(true).setForegroundColor("#1a237e");

        var blob = 圖表列表[ci].getBlob();
        var img = 圖表頁.insertImage(blob);
        img.setLeft(40);
        img.setTop(60);
        img.setWidth(640);
        img.setHeight(380);
      }
    }

    if (趨勢圖表列表.length > 0) {
      for (var ti = 0; ti < 趨勢圖表列表.length; ti++) {
        var 趨勢頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        趨勢頁.getBackground().setSolidFill("#ffffff");

        var trendTitle = 趨勢圖表列表[ti].getOptions().get("title") || "趨勢圖 " + (ti + 1);
        var tt = 趨勢頁.insertTextBox("📈 " + trendTitle, 40, 10, 640, 40);
        tt.getText().getTextStyle().setFontSize(24).setBold(true).setForegroundColor("#1a237e");

        var tBlob = 趨勢圖表列表[ti].getBlob();
        var tImg = 趨勢頁.insertImage(tBlob);
        tImg.setLeft(40);
        tImg.setTop(60);
        tImg.setWidth(640);
        tImg.setHeight(380);
      }
    }

    // --- 最後一頁：結語 ---
    var 結語頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    結語頁.getBackground().setSolidFill("#0d47a1");

    var endText = 結語頁.insertTextBox("感謝聆聽", 40, 150, 640, 60);
    endText.getText().getTextStyle().setFontSize(40).setBold(true).setForegroundColor("#ffffff");
    endText.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var endSub = 結語頁.insertTextBox(
      "本報告由 Google Apps Script 自動產生\n" +
      "資料來源：" + ss.getName(),
      40, 230, 640, 50
    );
    endSub.getText().getTextStyle().setFontSize(14).setForegroundColor("#90caf9");
    endSub.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // ========== 完成 ==========
    var 結束時間 = new Date();
    var 耗時 = Math.round((結束時間 - 開始時間) / 1000);

    Logger.log("✅ 報表簡報已產生！耗時 " + 耗時 + " 秒");
    Logger.log("簡報 URL：" + ppt.getUrl());

    SpreadsheetApp.getUi().alert(
      "🎉 報表簡報已成功產生！\n\n" +
      "📄 共 " + ppt.getSlides().length + " 頁\n" +
      "⏱️ 耗時 " + 耗時 + " 秒\n\n" +
      "📎 簡報名稱：" + 簡報名 + "\n\n" +
      "🔗 點選連結開啟：\n" + ppt.getUrl()
    );

    return ppt;

  } catch (錯誤) {
    Logger.log("❌ 報表簡報錯誤：" + 錯誤.message);
    Logger.log("錯誤堆疊：" + 錯誤.stack);
    SpreadsheetApp.getUi().alert("❌ 發生錯誤：" + 錯誤.message);
    return null;
  }
}

// ============================================================
// 延伸練習 1：PDF 匯出並存到 Google Drive
// ============================================================

/**
 * 取得（或建立）Drive 資料夾
 * @param {string} folderName - 資料夾名稱
 * @returns {Folder} Drive 資料夾物件
 */
function 取得Drive資料夾_(folderName) {
  var 根目錄 = DriveApp.getRootFolder();
  var 搜尋結果 = 根目錄.getFoldersByName(folderName);
  if (搜尋結果.hasNext()) {
    return 搜尋結果.next();
  } else {
    Logger.log("建立新資料夾：" + folderName);
    return 根目錄.createFolder(folderName);
  }
}

/**
 * 延伸練習 1：將簡報匯出為 PDF 並存到 Google Drive
 * 說明：
 *   1. 呼叫「一鍵產生報表簡報」取得簡報物件
 *   2. 透過 Drive API 的 export 連結下載 PDF blob
 *   3. 將 PDF blob 存入指定的 Drive 資料夾
 */
function 匯出PDF到Drive() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert("⏳ 正在產生簡報並匯出 PDF...\n\n請稍候約 20~40 秒。");

    // Step 1：產生簡報（重用既有函式）
    var ppt = 一鍵產生報表簡報();
    if (!ppt) {
      ui.alert("❌ 簡報產生失敗，無法匯出 PDF。");
      return;
    }

    // Step 2：確保簡報已完全寫入（等待 Google 伺服器同步）
    SpreadsheetApp.flush();
    Utilities.sleep(3000);  // 等待 3 秒

    // Step 3：透過 export URL 取得 PDF Blob
    //   格式：https://docs.google.com/presentation/d/{ID}/export/pdf
    var pptId = ppt.getId();
    var exportUrl = "https://docs.google.com/presentation/d/" + pptId + "/export/pdf";

    var 回應 = UrlFetchApp.fetch(exportUrl, {
      headers: {
        Authorization: "Bearer " + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });

    if (回應.getResponseCode() !== 200) {
      ui.alert("❌ PDF 匯出失敗！\nHTTP 狀態：" + 回應.getResponseCode());
      return;
    }

    var pdfBlob = 回應.getBlob();

    // Step 4：組合 PDF 檔名
    var 日期字串 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmm");
    var pdf檔名 = "營運分析報告_" + 日期字串 + ".pdf";
    pdfBlob.setName(pdf檔名);

    // Step 5：存入 Drive 資料夾
    var 資料夾 = 取得Drive資料夾_(CONFIG.Drive資料夾名);
    var pdfFile = 資料夾.createFile(pdfBlob);

    // Step 6：記錄與提示
    Logger.log("✅ PDF 已存入 Drive：" + pdfFile.getUrl());

    ui.alert(
      "🎉 PDF 匯出成功！\n\n" +
      "📁 存放資料夾：" + CONFIG.Drive資料夾名 + "\n" +
      "📄 檔名：" + pdf檔名 + "\n\n" +
      "🔗 點選連結開啟 PDF：\n" + pdfFile.getUrl()
    );

  } catch (錯誤) {
    Logger.log("❌ PDF 匯出錯誤：" + 錯誤.message);
    Logger.log("錯誤堆疊：" + 錯誤.stack);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 延伸練習 2：Email 通知主管（含簡報連結與 PDF 附件）
// ============================================================

/**
 * 延伸練習 2：產生簡報後，自動寄 Email 給主管
 * 說明：
 *   - 寄送人：執行者的 Google 帳號
 *   - 收件人：CONFIG.主管Email
 *   - 內容：包含簡報連結、摘要統計數據
 *   - 附件：PDF（選用，如需附加 PDF 請設定 附帶PDF = true）
 */
function 寄送報告Email() {
  try {
    var ui = SpreadsheetApp.getUi();

    // ── 確認主管 Email ─────────────────────────────────────
    if (!CONFIG.主管Email || CONFIG.主管Email === "supervisor@example.com") {
      ui.alert(
        "⚠️ 請先設定主管 Email！\n\n" +
        "請修改程式碼頂部的 CONFIG.主管Email\n" +
        "（例如：supervisor@yourcompany.com）"
      );
      return;
    }

    ui.alert("⏳ 正在產生簡報並準備寄送 Email...\n\n請稍候約 20~40 秒。");

    // ── Step 1：讀取試算表資料（用於 Email 摘要）────────────
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 部門表 = ss.getSheetByName("圖表資料");
    var 趨勢表 = ss.getSheetByName("月度趨勢");

    if (!部門表 || !趨勢表) {
      ui.alert("❌ 請先執行「初始化圖表資料」！");
      return;
    }

    // 計算摘要統計
    var 部門資料 = 部門表.getRange("A2:E7").getValues();
    var 趨勢資料 = 趨勢表.getRange("B2:D13").getValues();
    var 年度總營收 = 0;
    var 冠軍部門 = "", 冠軍業績 = 0;
    for (var i = 0; i < 部門資料.length; i++) {
      var 合計 = 部門資料[i][1] + 部門資料[i][2] + 部門資料[i][3] + 部門資料[i][4];
      年度總營收 += 合計;
      if (合計 > 冠軍業績) { 冠軍業績 = 合計; 冠軍部門 = 部門資料[i][0]; }
    }
    var 總利潤 = 0;
    for (var j = 0; j < 趨勢資料.length; j++) { 總利潤 += 趨勢資料[j][2]; }
    var 利潤率 = Math.round((總利潤 / 年度總營收) * 100);

    // ── Step 2：產生簡報 ───────────────────────────────────
    var ppt = 一鍵產生報表簡報();
    if (!ppt) {
      ui.alert("❌ 簡報產生失敗，無法寄送 Email。");
      return;
    }
    var pptUrl = ppt.getUrl();
    var pptId  = ppt.getId();

    // ── Step 3：下載 PDF 作為附件 ──────────────────────────
    SpreadsheetApp.flush();
    Utilities.sleep(3000);

    var exportUrl = "https://docs.google.com/presentation/d/" + pptId + "/export/pdf";
    var 回應 = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    var pdfBlob = null;
    if (回應.getResponseCode() === 200) {
      var 日期字串 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd");
      pdfBlob = 回應.getBlob().setName("營運分析報告_" + 日期字串 + ".pdf");
    }

    // ── Step 4：組合 Email ─────────────────────────────────
    var 今日 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月 dd 日");
    var 寄件人名稱 = Session.getActiveUser().getEmail().split("@")[0];

    // HTML 內文
    var htmlBody =
      "<div style='font-family:Arial,sans-serif;max-width:600px;margin:auto;'>" +
      "  <div style='background:#0d47a1;padding:24px;border-radius:8px 8px 0 0;'>" +
      "    <h1 style='color:#fff;margin:0;font-size:22px;'>📊 2026 年度營運分析報告</h1>" +
      "    <p style='color:#90caf9;margin:6px 0 0;'>自動產生日期：" + 今日 + "</p>" +
      "  </div>" +
      "  <div style='background:#fafafa;padding:24px;border:1px solid #e0e0e0;'>" +
      "    <p>您好，</p>" +
      "    <p>本月份營運分析報告已自動產生完畢，請參閱以下重點摘要：</p>" +
      "    <table style='width:100%;border-collapse:collapse;margin:16px 0;'>" +
      "      <tr style='background:#1a237e;color:#fff;'>" +
      "        <th style='padding:8px 12px;text-align:left;'>指標</th>" +
      "        <th style='padding:8px 12px;text-align:right;'>數值</th>" +
      "      </tr>" +
      "      <tr style='background:#e8eaf6;'>" +
      "        <td style='padding:8px 12px;'>年度總營收</td>" +
      "        <td style='padding:8px 12px;text-align:right;font-weight:bold;'>NT$ " + (年度總營收 / 10000).toFixed(0) + " 萬</td>" +
      "      </tr>" +
      "      <tr>" +
      "        <td style='padding:8px 12px;'>平均利潤率</td>" +
      "        <td style='padding:8px 12px;text-align:right;font-weight:bold;color:#2e7d32;'>" + 利潤率 + " %</td>" +
      "      </tr>" +
      "      <tr style='background:#e8eaf6;'>" +
      "        <td style='padding:8px 12px;'>年度業績冠軍部門</td>" +
      "        <td style='padding:8px 12px;text-align:right;font-weight:bold;'>🏆 " + 冠軍部門 + "</td>" +
      "      </tr>" +
      "    </table>" +
      "    <div style='text-align:center;margin:24px 0;'>" +
      "      <a href='" + pptUrl + "' style='background:#1a73e8;color:#fff;padding:12px 28px;" +
      "         border-radius:6px;text-decoration:none;font-size:15px;font-weight:bold;'>" +
      "        📑 開啟完整簡報報告" +
      "      </a>" +
      "    </div>" +
      "    <p style='color:#757575;font-size:12px;'>" +
      "      本郵件及附件由 Google Apps Script 自動產生。<br>" +
      "      資料來源：" + ss.getName() + "<br>" +
      "      寄件人：" + 寄件人名稱 +
      "    </p>" +
      "  </div>" +
      "</div>";

    // 純文字備用版本
    var plainBody =
      "【2026 年度營運分析報告】\n" +
      "產生日期：" + 今日 + "\n\n" +
      "📊 重點摘要：\n" +
      "  年度總營收：NT$ " + (年度總營收 / 10000).toFixed(0) + " 萬\n" +
      "  平均利潤率：" + 利潤率 + " %\n" +
      "  業績冠軍：" + 冠軍部門 + "\n\n" +
      "🔗 完整簡報連結：\n" + pptUrl + "\n\n" +
      "（本郵件由 Google Apps Script 自動發送）";

    // ── Step 5：寄送 Email ─────────────────────────────────
    var mailOptions = {
      name: "GAS 自動報告系統",
      htmlBody: htmlBody,
      replyTo: Session.getActiveUser().getEmail()
    };
    // 如果 PDF 成功下載，加入附件
    if (pdfBlob) {
      mailOptions.attachments = [pdfBlob];
    }

    MailApp.sendEmail(
      CONFIG.主管Email,
      "【自動通知】2026 年度營運分析報告 - " + 今日,
      plainBody,
      mailOptions
    );

    // ── Step 6：記錄與提示 ─────────────────────────────────
    Logger.log("✅ Email 已寄送至：" + CONFIG.主管Email);

    ui.alert(
      "🎉 Email 已成功寄出！\n\n" +
      "📧 收件人：" + CONFIG.主管Email + "\n" +
      (pdfBlob ? "📎 含 PDF 附件\n" : "") +
      "🔗 簡報連結已附於信中\n\n" +
      "請提醒主管查收信箱。"
    );

  } catch (錯誤) {
    Logger.log("❌ Email 寄送錯誤：" + 錯誤.message);
    Logger.log("錯誤堆疊：" + 錯誤.stack);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 延伸練習 3：模板套用（用程式填入動態資料，非從零建立）
// ============================================================

/**
 * 延伸練習 3 - Step A：建立報告模板簡報
 * 說明：
 *   建立一份含有「佔位標記」的模板簡報，標記格式為 {{欄位名}}。
 *   建立完成後，將簡報 ID 複製到 CONFIG.模板簡報ID，即可重複套用。
 *
 * 模板標記說明：
 *   {{年份}}       → 報告年份
 *   {{產生日期}}   → 自動產生日期
 *   {{年度總營收}} → 年度總營收金額
 *   {{利潤率}}     → 平均利潤率百分比
 *   {{冠軍部門}}   → 年度業績冠軍部門
 *   {{冠軍業績}}   → 冠軍部門年度業績
 *   {{來源試算表}} → 資料來源試算表名稱
 *   {{重點摘要}}   → 三點式重點摘要文字
 */
function 建立報告模板() {
  try {
    var ui = SpreadsheetApp.getUi();

    var 模板 = SlidesApp.create("【模板】年度營運分析報告");

    // ─────────────────────────────────────────────
    // 模板第 1 頁：封面
    // ─────────────────────────────────────────────
    var 封面 = 模板.getSlides()[0];
    封面.getBackground().setSolidFill("#0d47a1");
    封面.getPlaceholders().forEach(function(p) { p.remove(); });

    var 主標題 = 封面.insertTextBox("📊 {{年份}} 年度營運分析報告", 40, 120, 640, 80);
    主標題.getText().getTextStyle().setFontSize(34).setBold(true).setForegroundColor("#ffffff");
    主標題.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var 副標題 = 封面.insertTextBox("產生日期：{{產生日期}}\n由 Google Apps Script 自動生成", 40, 220, 640, 50);
    副標題.getText().getTextStyle().setFontSize(15).setForegroundColor("#90caf9");
    副標題.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var 封面線 = 封面.insertLine(SlidesApp.LineCategory.STRAIGHT, 180, 210, 540, 210);
    封面線.getLineFill().setSolidFill("#64b5f6");
    封面線.setWeight(2);

    // ─────────────────────────────────────────────
    // 模板第 2 頁：KPI 摘要
    // ─────────────────────────────────────────────
    var KPI頁 = 模板.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    KPI頁.getBackground().setSolidFill("#fafafa");

    var kpiPageTitle = KPI頁.insertTextBox("🎯 關鍵績效指標 (KPI)", 40, 15, 640, 45);
    kpiPageTitle.getText().getTextStyle().setFontSize(26).setBold(true).setForegroundColor("#1a237e");

    // KPI 卡片（使用標記作為內容）
    var kpiTemplate = [
      { 標記: "NT$ {{年度總營收}}", 標籤: "年度總營收", 色: "#1a73e8", x: 40 },
      { 標記: "{{利潤率}} %",       標籤: "平均利潤率", 色: "#34a853", x: 205 },
      { 標記: "{{冠軍部門}}",       標籤: "業績冠軍",   色: "#ea4335", x: 370 },
      { 標記: "{{冠軍業績}}",       標籤: "冠軍業績",   色: "#fbbc04", x: 535 }
    ];

    for (var k = 0; k < kpiTemplate.length; k++) {
      var card = KPI頁.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, kpiTemplate[k].x, 75, 150, 120);
      card.getFill().setSolidFill(kpiTemplate[k].色);
      card.getBorder().setWeight(0);

      var 值框 = KPI頁.insertTextBox(kpiTemplate[k].標記, kpiTemplate[k].x + 5, 95, 140, 45);
      值框.getText().getTextStyle().setFontSize(14).setBold(true).setForegroundColor("#ffffff");
      值框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      var 標籤框 = KPI頁.insertTextBox(kpiTemplate[k].標籤, kpiTemplate[k].x + 5, 150, 140, 28);
      標籤框.getText().getTextStyle().setFontSize(10).setForegroundColor("#f0f0f0");
      標籤框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }

    // ─────────────────────────────────────────────
    // 模板第 3 頁：重點摘要
    // ─────────────────────────────────────────────
    var 摘要頁 = 模板.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    摘要頁.getBackground().setSolidFill("#f5f5f5");

    var 摘要頁標題 = 摘要頁.insertTextBox("💡 本年度重點摘要", 40, 20, 640, 50);
    摘要頁標題.getText().getTextStyle().setFontSize(28).setBold(true).setForegroundColor("#1a237e");

    var 摘要框 = 摘要頁.insertTextBox("{{重點摘要}}", 60, 90, 600, 300);
    摘要框.getText().getTextStyle().setFontSize(16).setForegroundColor("#37474f");

    var 備註框 = 摘要頁.insertTextBox("資料來源：{{來源試算表}}", 40, 420, 640, 30);
    備註框.getText().getTextStyle().setFontSize(11).setForegroundColor("#9e9e9e");

    // ─────────────────────────────────────────────
    // 模板第 4 頁：結語
    // ─────────────────────────────────────────────
    var 結語頁 = 模板.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    結語頁.getBackground().setSolidFill("#0d47a1");

    var 結語框 = 結語頁.insertTextBox("感謝聆聽", 40, 150, 640, 60);
    結語框.getText().getTextStyle().setFontSize(40).setBold(true).setForegroundColor("#ffffff");
    結語框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var 結語副框 = 結語頁.insertTextBox(
      "本報告由 Google Apps Script 自動產生\n資料來源：{{來源試算表}}",
      40, 230, 640, 50
    );
    結語副框.getText().getTextStyle().setFontSize(14).setForegroundColor("#90caf9");
    結語副框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // ─────────────────────────────────────────────
    // 輸出模板 ID
    // ─────────────────────────────────────────────
    var 模板ID = 模板.getId();
    Logger.log("✅ 模板已建立！ID：" + 模板ID);

    ui.alert(
      "✅ 報告模板已建立！\n\n" +
      "🔑 模板 ID（請複製到 CONFIG.模板簡報ID）：\n" +
      模板ID + "\n\n" +
      "📝 模板內含以下佔位標記：\n" +
      "  {{年份}}、{{產生日期}}、{{年度總營收}}\n" +
      "  {{利潤率}}、{{冠軍部門}}、{{冠軍業績}}\n" +
      "  {{來源試算表}}、{{重點摘要}}\n\n" +
      "🔗 模板連結：\n" + 模板.getUrl()
    );

  } catch (錯誤) {
    Logger.log("❌ 建立模板錯誤：" + 錯誤.message);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

/**
 * 延伸練習 3 - Step B：套用模板，產生本月報告
 * 說明：
 *   1. 複製模板簡報（保留原模板不修改）
 *   2. 讀取試算表資料，計算各項指標
 *   3. 用 replaceAllText() 將所有 {{標記}} 替換為實際數值
 *   4. 在圖表佔位頁插入實際圖表
 */
function 套用模板產生報告() {
  try {
    var ui = SpreadsheetApp.getUi();

    // ── 確認模板 ID ────────────────────────────────────────
    var 模板ID = CONFIG.模板簡報ID;
    if (!模板ID) {
      var 確認 = ui.alert(
        "⚠️ 尚未設定模板簡報 ID！\n\n" +
        "是否要先執行「建立報告模板」？",
        ui.ButtonSet.YES_NO
      );
      if (確認 === ui.Button.YES) {
        建立報告模板();
      }
      return;
    }

    ui.alert("⏳ 正在套用模板產生報告...\n\n請稍候約 15~25 秒。");

    // ── Step 1：讀取試算表資料 ─────────────────────────────
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 部門表 = ss.getSheetByName("圖表資料");
    var 趨勢表 = ss.getSheetByName("月度趨勢");

    if (!部門表 || !趨勢表) {
      ui.alert("❌ 請先執行「初始化圖表資料」！");
      return;
    }

    var 部門資料 = 部門表.getRange("A2:E7").getValues();
    var 趨勢資料 = 趨勢表.getRange("B2:D13").getValues();

    // 計算 KPI 數值
    var 年度總營收 = 0, 總利潤 = 0;
    var 冠軍部門 = "", 冠軍業績 = 0;

    for (var i = 0; i < 部門資料.length; i++) {
      var 部門年合計 = 部門資料[i][1] + 部門資料[i][2] + 部門資料[i][3] + 部門資料[i][4];
      年度總營收 += 部門年合計;
      if (部門年合計 > 冠軍業績) {
        冠軍業績 = 部門年合計;
        冠軍部門 = 部門資料[i][0];
      }
    }
    for (var j = 0; j < 趨勢資料.length; j++) { 總利潤 += 趨勢資料[j][2]; }
    var 利潤率 = Math.round((總利潤 / 年度總營收) * 100);

    // 計算 Q4 vs Q1 整體成長率
    var Q1總 = 0, Q4總 = 0;
    for (var q = 0; q < 部門資料.length; q++) {
      Q1總 += 部門資料[q][1];
      Q4總 += 部門資料[q][4];
    }
    var Q4成長率 = Math.round(((Q4總 - Q1總) / Q1總) * 100);

    var 今日 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月 dd 日");
    var 年份 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy");

    var 重點摘要 =
      "✅ " + 冠軍部門 + " 持續領先，年度業績達 NT$ " + (冠軍業績 / 10000).toFixed(0) + " 萬\n\n" +
      "📈 年度總營收 NT$ " + (年度總營收 / 10000).toFixed(0) + " 萬，利潤率 " + 利潤率 + "%\n\n" +
      "🚀 Q4 對比 Q1 整體成長 " + Q4成長率 + "%，成長動能強勁\n\n" +
      "🎯 下季目標：維持利潤率並持續開拓新市場";

    // ── Step 2：複製模板（不動原模板）─────────────────────
    var 模板檔 = DriveApp.getFileById(模板ID);
    var 日期字串 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy.MM.dd");
    var 新簡報名 = "📊 " + 年份 + " 年度營運報告（套用模板）- " + 日期字串;
    var 新檔 = 模板檔.makeCopy(新簡報名);

    // 開啟複本並編輯
    var ppt = SlidesApp.openById(新檔.getId());

    // ── Step 3：批次替換所有佔位標記 ──────────────────────
    //   replaceAllText(searchPattern, replacement) 會取代全簡報所有投影片中的文字
    var 替換對照表 = {
      "{{年份}}":       年份,
      "{{產生日期}}":   今日,
      "{{年度總營收}}": (年度總營收 / 10000).toFixed(0) + " 萬",
      "{{利潤率}}":     String(利潤率),
      "{{冠軍部門}}":   冠軍部門,
      "{{冠軍業績}}":   "NT$ " + (冠軍業績 / 10000).toFixed(0) + " 萬",
      "{{來源試算表}}": ss.getName(),
      "{{重點摘要}}":   重點摘要
    };

    for (var 標記 in 替換對照表) {
      ppt.replaceAllText(標記, 替換對照表[標記]);
    }

    // ── Step 4：在圖表頁插入實際圖表（附加在最後）─────────
    var 圖表列表 = 部門表.getCharts();
    var 趨勢圖表列表 = 趨勢表.getCharts();
    var 有圖表 = 圖表列表.length > 0 || 趨勢圖表列表.length > 0;

    if (有圖表) {
      // 加入圖表分隔頁
      var 圖表標題頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
      圖表標題頁.getBackground().setSolidFill("#1a237e");
      var 圖表標題框 = 圖表標題頁.insertTextBox("📊 圖表分析", 40, 190, 640, 60);
      圖表標題框.getText().getTextStyle().setFontSize(36).setBold(true).setForegroundColor("#ffffff");
      圖表標題框.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

      // 插入部門圖表
      for (var ci = 0; ci < 圖表列表.length; ci++) {
        var 圖表頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        圖表頁.getBackground().setSolidFill("#ffffff");
        var 圖表名 = 圖表列表[ci].getOptions().get("title") || ("圖表 " + (ci + 1));
        var ctBox = 圖表頁.insertTextBox("📊 " + 圖表名, 40, 10, 640, 40);
        ctBox.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#1a237e");
        var img = 圖表頁.insertImage(圖表列表[ci].getBlob());
        img.setLeft(40).setTop(58).setWidth(640).setHeight(380);
      }

      // 插入趨勢圖表
      for (var ti = 0; ti < 趨勢圖表列表.length; ti++) {
        var 趨勢頁 = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);
        趨勢頁.getBackground().setSolidFill("#ffffff");
        var 趨勢名 = 趨勢圖表列表[ti].getOptions().get("title") || ("趨勢圖 " + (ti + 1));
        var ttBox = 趨勢頁.insertTextBox("📈 " + 趨勢名, 40, 10, 640, 40);
        ttBox.getText().getTextStyle().setFontSize(22).setBold(true).setForegroundColor("#1a237e");
        var tImg = 趨勢頁.insertImage(趨勢圖表列表[ti].getBlob());
        tImg.setLeft(40).setTop(58).setWidth(640).setHeight(380);
      }
    }

    // ── Step 5：儲存並提示 ─────────────────────────────────
    ppt.saveAndClose();
    Logger.log("✅ 模板套用完成！URL：" + 新檔.getUrl());

    ui.alert(
      "🎉 模板套用完成！\n\n" +
      "📎 報告名稱：" + 新簡報名 + "\n" +
      "🔄 替換標記數：" + Object.keys(替換對照表).length + " 個\n" +
      (有圖表 ? "📊 已附加圖表頁\n" : "") +
      "\n🔗 開啟報告：\n" + 新檔.getUrl()
    );

  } catch (錯誤) {
    Logger.log("❌ 模板套用錯誤：" + 錯誤.message);
    Logger.log("錯誤堆疊：" + 錯誤.stack);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 🚀 整合功能：一鍵全流程（產生 → PDF → Email）
// ============================================================

/**
 * 一鍵執行完整流程：
 *   產生報表簡報 → 匯出 PDF 存 Drive → 寄送 Email 給主管
 * 說明：
 *   此函式整合三項延伸練習，執行一次完成所有動作。
 *   請確保 CONFIG.主管Email 已正確設定。
 */
function 一鍵全流程() {
  try {
    var ui = SpreadsheetApp.getUi();

    if (!CONFIG.主管Email || CONFIG.主管Email === "supervisor@example.com") {
      ui.alert(
        "⚠️ 請先設定主管 Email！\n\n" +
        "修改程式頂部 CONFIG.主管Email 後再執行。"
      );
      return;
    }

    ui.alert(
      "🚀 即將執行完整流程：\n\n" +
      "  Step 1：產生報表簡報\n" +
      "  Step 2：匯出 PDF 存至 Google Drive\n" +
      "  Step 3：寄送 Email 至 " + CONFIG.主管Email + "\n\n" +
      "⏳ 預計需要 30~60 秒，點選「確定」開始。"
    );

    // ── Step 1：讀取資料（共用）────────────────────────────
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 部門表 = ss.getSheetByName("圖表資料");
    var 趨勢表 = ss.getSheetByName("月度趨勢");
    if (!部門表 || !趨勢表) {
      ui.alert("❌ 請先執行「初始化圖表資料」！");
      return;
    }

    var 部門資料 = 部門表.getRange("A2:E7").getValues();
    var 趨勢資料 = 趨勢表.getRange("B2:D13").getValues();

    var 年度總營收 = 0, 總利潤 = 0;
    var 冠軍部門 = "", 冠軍業績 = 0;
    for (var i = 0; i < 部門資料.length; i++) {
      var 合計 = 部門資料[i][1] + 部門資料[i][2] + 部門資料[i][3] + 部門資料[i][4];
      年度總營收 += 合計;
      if (合計 > 冠軍業績) { 冠軍業績 = 合計; 冠軍部門 = 部門資料[i][0]; }
    }
    for (var j = 0; j < 趨勢資料.length; j++) { 總利潤 += 趨勢資料[j][2]; }
    var 利潤率 = Math.round((總利潤 / 年度總營收) * 100);

    var 今日 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy 年 MM 月 dd 日");

    // ── Step 2：產生簡報 ────────────────────────────────────
    Logger.log("全流程 Step 1：產生簡報...");
    var ppt = 一鍵產生報表簡報();
    if (!ppt) {
      ui.alert("❌ 簡報產生失敗，流程中止。");
      return;
    }
    var pptUrl = ppt.getUrl();
    var pptId  = ppt.getId();

    // ── Step 3：匯出 PDF ────────────────────────────────────
    Logger.log("全流程 Step 2：匯出 PDF...");
    SpreadsheetApp.flush();
    Utilities.sleep(3000);

    var exportUrl = "https://docs.google.com/presentation/d/" + pptId + "/export/pdf";
    var 回應 = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    var pdfFile = null;
    var pdfBlob = null;
    if (回應.getResponseCode() === 200) {
      var 日期字串 = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmm");
      pdfBlob = 回應.getBlob().setName("營運分析報告_" + 日期字串 + ".pdf");
      var 資料夾 = 取得Drive資料夾_(CONFIG.Drive資料夾名);
      pdfFile = 資料夾.createFile(pdfBlob);
      Logger.log("PDF 已存入：" + pdfFile.getUrl());
    }

    // ── Step 4：重新下載 PDF blob 供 Email 附件使用 ─────────
    var emailPdfBlob = null;
    if (回應.getResponseCode() === 200) {
      var 回應2 = UrlFetchApp.fetch(exportUrl, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      if (回應2.getResponseCode() === 200) {
        emailPdfBlob = 回應2.getBlob().setName("營運分析報告.pdf");
      }
    }

    // ── Step 5：寄送 Email ─────────────────────────────────
    Logger.log("全流程 Step 3：寄送 Email...");

    var htmlBody =
      "<div style='font-family:Arial,sans-serif;max-width:600px;margin:auto;'>" +
      "  <div style='background:#0d47a1;padding:24px;border-radius:8px 8px 0 0;'>" +
      "    <h1 style='color:#fff;margin:0;font-size:22px;'>📊 年度營運分析報告</h1>" +
      "    <p style='color:#90caf9;margin:6px 0 0;'>自動產生日期：" + 今日 + "</p>" +
      "  </div>" +
      "  <div style='background:#fafafa;padding:24px;border:1px solid #e0e0e0;'>" +
      "    <p>您好，</p><p>本月份報告已自動產生，重點摘要如下：</p>" +
      "    <table style='width:100%;border-collapse:collapse;margin:16px 0;'>" +
      "      <tr style='background:#1a237e;color:#fff;'>" +
      "        <th style='padding:8px 12px;text-align:left;'>指標</th>" +
      "        <th style='padding:8px 12px;text-align:right;'>數值</th></tr>" +
      "      <tr style='background:#e8eaf6;'><td style='padding:8px 12px;'>年度總營收</td>" +
      "        <td style='padding:8px 12px;text-align:right;font-weight:bold;'>NT$ " + (年度總營收 / 10000).toFixed(0) + " 萬</td></tr>" +
      "      <tr><td style='padding:8px 12px;'>平均利潤率</td>" +
      "        <td style='padding:8px 12px;text-align:right;font-weight:bold;color:#2e7d32;'>" + 利潤率 + " %</td></tr>" +
      "      <tr style='background:#e8eaf6;'><td style='padding:8px 12px;'>業績冠軍部門</td>" +
      "        <td style='padding:8px 12px;text-align:right;font-weight:bold;'>🏆 " + 冠軍部門 + "</td></tr>" +
      "    </table>" +
      "    <div style='text-align:center;margin:24px 0;'>" +
      "      <a href='" + pptUrl + "' style='background:#1a73e8;color:#fff;padding:12px 28px;" +
      "         border-radius:6px;text-decoration:none;font-size:15px;font-weight:bold;'>" +
      "        📑 開啟完整簡報報告</a>" +
      (pdfFile ? "<br><br><a href='" + pdfFile.getUrl() + "' style='color:#1a73e8;'>📄 在 Drive 查看 PDF</a>" : "") +
      "    </div>" +
      "    <p style='color:#757575;font-size:12px;'>本郵件由 Google Apps Script 自動發送。資料來源：" + ss.getName() + "</p>" +
      "  </div></div>";

    var mailOptions = {
      name: "GAS 自動報告系統",
      htmlBody: htmlBody
    };
    if (emailPdfBlob) { mailOptions.attachments = [emailPdfBlob]; }

    MailApp.sendEmail(
      CONFIG.主管Email,
      "【自動通知】年度營運分析報告 - " + 今日,
      "報告已產生，請開啟 HTML 版本查看。",
      mailOptions
    );

    // ── Step 6：完成摘要 ────────────────────────────────────
    Logger.log("✅ 全流程完成！");
    ui.alert(
      "🎉 全流程執行完成！\n\n" +
      "✅ 報表簡報已產生\n" +
      (pdfFile ? "✅ PDF 已存至 Drive（" + CONFIG.Drive資料夾名 + "）\n" : "⚠️ PDF 匯出失敗（簡報仍可用）\n") +
      "✅ Email 已寄送至 " + CONFIG.主管Email + "\n\n" +
      "🔗 簡報連結：\n" + pptUrl
    );

  } catch (錯誤) {
    Logger.log("❌ 全流程錯誤：" + 錯誤.message);
    Logger.log("錯誤堆疊：" + 錯誤.stack);
    SpreadsheetApp.getUi().alert("❌ 錯誤：" + 錯誤.message);
  }
}

// ============================================================
// 初始化（沿用 Session 9 的資料）
// ============================================================

/**
 * 確認資料是否就緒
 */
function 檢查資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var 圖表表 = ss.getSheetByName("圖表資料");
  var 趨勢表 = ss.getSheetByName("月度趨勢");

  var 訊息 = "📋 資料檢查結果：\n\n";

  if (圖表表) {
    var 圖表數 = 圖表表.getCharts().length;
    訊息 += "✅ 圖表資料：已就緒（" + 圖表數 + " 個圖表）\n";
  } else {
    訊息 += "❌ 圖表資料：未建立\n";
  }

  if (趨勢表) {
    var 趨勢圖表數 = 趨勢表.getCharts().length;
    訊息 += "✅ 月度趨勢：已就緒（" + 趨勢圖表數 + " 個圖表）\n";
  } else {
    訊息 += "❌ 月度趨勢：未建立\n";
  }

  訊息 += "\n⚙️  CONFIG 設定：\n";
  訊息 += "  主管Email：" + CONFIG.主管Email + "\n";
  訊息 += "  Drive資料夾：" + CONFIG.Drive資料夾名 + "\n";
  訊息 += "  模板ID：" + (CONFIG.模板簡報ID || "（尚未設定）") + "\n";

  if (!圖表表 || !趨勢表) {
    訊息 += "\n⚠️ 請先執行「初始化圖表資料」！";
  }

  SpreadsheetApp.getUi().alert(訊息);
}

// ============================================================
// 自訂選單
// ============================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📚 Session 10 工具")
    .addItem("🔍 檢查資料與設定", "檢查資料")
    .addSeparator()
    // ── 基本功能 ──
    .addSubMenu(
      ui.createMenu("📝 基本簡報操作")
        .addItem("📝 建立基本簡報",     "建立基本簡報")
        .addItem("📑 建立多頁簡報",     "建立多頁簡報")
        .addItem("📊 插入圖表到簡報",   "插入圖表到簡報")
        .addSeparator()
        .addItem("🚀 一鍵產生報表簡報", "一鍵產生報表簡報")
    )
    .addSeparator()
    // ── 延伸練習 ──
    .addSubMenu(
      ui.createMenu("🏋️ 延伸練習")
        .addItem("📄 練習1：匯出 PDF 到 Drive",    "匯出PDF到Drive")
        .addSeparator()
        .addItem("📧 練習2：Email 通知主管",       "寄送報告Email")
        .addSeparator()
        .addItem("🗂️ 練習3A：建立報告模板",       "建立報告模板")
        .addItem("🗂️ 練習3B：套用模板產生報告",   "套用模板產生報告")
    )
    .addSeparator()
    // ── 整合功能 ──
    .addItem("🚀 一鍵全流程（產生→PDF→Email）", "一鍵全流程")
    .addToUi();
}
