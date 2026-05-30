// ============================================================
// 進階練習：智慧採購分析與供應商評比
// 對應：Session 6（filter、sort、資料摘要）
// ============================================================

/**
 * 供應商綜合評比（篩選 + 排序 + 統計）
 */
function 供應商評比() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("採購紀錄");
    if (!sheet) { SpreadsheetApp.getUi().alert("❌ 請先初始化"); return; }

    var 資料 = sheet.getDataRange().getValues();
    var 標題 = 資料[0];
    var 紀錄 = [];
    for (var i = 1; i < 資料.length; i++) {
      var obj = {};
      for (var j = 0; j < 標題.length; j++) obj[標題[j]] = 資料[i][j];
      紀錄.push(obj);
    }

    // 依供應商分組統計
    var 供應商統計 = {};
    紀錄.forEach(function(r) {
      var s = r["供應商"];
      if (!供應商統計[s]) {
        供應商統計[s] = { 訂單數: 0, 總金額: 0, 準時交貨: 0, 退貨次數: 0, 品質分數: [] };
      }
      供應商統計[s].訂單數++;
      供應商統計[s].總金額 += r["金額"];
      if (r["交貨狀態"] === "準時") 供應商統計[s].準時交貨++;
      if (r["退貨"] === "是") 供應商統計[s].退貨次數++;
      供應商統計[s].品質分數.push(r["品質評分"]);
    });

    // 計算綜合評分
    var 排名 = [];
    for (var name in 供應商統計) {
      var s = 供應商統計[name];
      var 準時率 = (s.準時交貨 / s.訂單數 * 100);
      var 退貨率 = (s.退貨次數 / s.訂單數 * 100);
      var 平均品質 = s.品質分數.reduce(function(a, b) { return a + b; }, 0) / s.品質分數.length;

      // 綜合評分 = 品質(40%) + 準時率(30%) + (100-退貨率)(20%) + 價格競爭力(10%)
      var 綜合 = 平均品質 * 0.4 + 準時率 * 0.3 + (100 - 退貨率) * 0.2 + 50 * 0.1;

      排名.push({
        供應商: name, 訂單數: s.訂單數, 總金額: s.總金額,
        準時率: 準時率, 退貨率: 退貨率, 品質: 平均品質, 綜合: 綜合
      });
    }

    // 依綜合評分排序
    排名.sort(function(a, b) { return b.綜合 - a.綜合; });

    // 產生評比報表
    var 評比表 = ss.getSheetByName("供應商評比");
    if (評比表) 評比表.clear(); else 評比表 = ss.insertSheet("供應商評比");

    評比表.getRange("A1").setValue("🏆 供應商綜合評比").setFontSize(16).setFontWeight("bold");
    評比表.getRange("A2").setValue("評比日期：" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd"));

    var 表頭 = [["排名", "供應商", "訂單數", "總金額", "準時率", "退貨率", "品質", "綜合評分", "等級"]];
    評比表.getRange(4, 1, 1, 9).setValues(表頭);
    評比表.getRange(4, 1, 1, 9).setBackground("#1565c0").setFontColor("#fff").setFontWeight("bold");

    排名.forEach(function(r, idx) {
      var 等級 = r.綜合 >= 80 ? "⭐ A" : r.綜合 >= 60 ? "🔵 B" : "🔴 C";
      評比表.getRange(5 + idx, 1, 1, 9).setValues([[
        idx + 1, r.供應商, r.訂單數, r.總金額,
        (r.準時率).toFixed(1) + "%", (r.退貨率).toFixed(1) + "%",
        r.品質.toFixed(1), r.綜合.toFixed(1), 等級
      ]]);
      if (idx === 0) 評比表.getRange(5 + idx, 1, 1, 9).setBackground("#e8f5e9"); // 冠軍綠
    });

    評比表.getRange(5, 4, 排名.length, 1).setNumberFormat("#,##0");
    for (var c = 1; c <= 9; c++) 評比表.autoResizeColumn(c);

    SpreadsheetApp.getUi().alert("✅ 供應商評比完成！冠軍：" + 排名[0].供應商);

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); }
}

/**
 * 採購需求預測（根據歷史消耗趨勢）
 */
function 採購預測() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("採購紀錄");
  if (!sheet) return;

  var 資料 = sheet.getDataRange().getValues();
  var 品項統計 = {};

  for (var i = 1; i < 資料.length; i++) {
    var 品項 = 資料[i][1]; // B: 品項
    var 數量 = 資料[i][3]; // D: 數量
    if (!品項統計[品項]) 品項統計[品項] = [];
    品項統計[品項].push(數量);
  }

  var 預測 = [];
  for (var name in 品項統計) {
    var 歷史 = 品項統計[name];
    var 平均 = 歷史.reduce(function(a, b) { return a + b; }, 0) / 歷史.length;
    var 建議量 = Math.ceil(平均 * 1.2); // 多 20% 安全庫存

    預測.push(name + "：月均 " + Math.round(平均) + " → 建議採購 " + 建議量);
  }

  SpreadsheetApp.getUi().alert("📊 採購預測\n\n" + 預測.join("\n"));
}

function 初始化採購資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("採購紀錄");
  if (!sheet) sheet = ss.insertSheet("採購紀錄"); else sheet.clear();

  var 標題 = [["供應商", "品項", "單價", "數量", "金額", "交貨狀態", "退貨", "品質評分", "日期"]];
  var 供應商 = ["宏達文具", "大同辦公", "金鼎耗材", "永豐科技", "佳能事務"];
  var 品項 = ["A4影印紙", "碳粉匣", "原子筆", "資料夾", "白板筆"];
  var 資料 = [];

  for (var i = 0; i < 40; i++) {
    var s = 供應商[Math.floor(Math.random() * 供應商.length)];
    var p = 品項[Math.floor(Math.random() * 品項.length)];
    var 單價 = [150, 1200, 15, 25, 45][品項.indexOf(p)];
    var 數量 = Math.floor(Math.random() * 50) + 5;
    var 交貨 = Math.random() > 0.2 ? "準時" : "延遲";
    var 退貨 = Math.random() > 0.85 ? "是" : "否";
    var 品質 = Math.floor(Math.random() * 40) + 60;
    var 日期 = new Date(2026, Math.floor(Math.random() * 4), Math.floor(Math.random() * 28) + 1);

    資料.push([s, p, 單價, 數量, 單價 * 數量, 交貨, 退貨, 品質, 日期]);
  }

  sheet.getRange(1, 1, 1, 9).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 9).setValues(資料);
  sheet.getRange("A1:I1").setBackground("#6a1b9a").setFontColor("#fff").setFontWeight("bold");
  sheet.getRange("E2:E41").setNumberFormat("#,##0");
  sheet.getRange("I2:I41").setNumberFormat("yyyy/mm/dd");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 9; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 40 筆採購紀錄已建立！");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 智慧採購管理")
    .addItem("📦 初始化採購資料", "初始化採購資料")
    .addItem("🏆 供應商評比", "供應商評比")
    .addItem("📊 採購預測", "採購預測")
    .addToUi();
}
