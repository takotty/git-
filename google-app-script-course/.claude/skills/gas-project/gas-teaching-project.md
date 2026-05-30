---
name: gas-teaching-project
description: Use when writing, editing, or debugging Google Apps Script (GAS) code — covers SpreadsheetApp, DriveApp, MailApp, SlidesApp, triggers, conditional formatting, charts, UrlFetchApp, HtmlService, CacheService, PropertiesService, data validation, and execution limits. Trigger on "寫 GAS", "Apps Script", "試算表自動化", "寄信通知", "外部 API", "自訂選單", "觸發器", or any Google Workspace automation request.
---

# Google Apps Script 核心技能參考

## 參考資源

| 類型 | 網址 |
|------|------|
| 官方文件首頁 | https://developers.google.com/apps-script |
| API 參考（所有服務） | https://developers.google.com/apps-script/reference |
| SpreadsheetApp | https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app |
| DriveApp | https://developers.google.com/apps-script/reference/drive/drive-app |
| MailApp / GmailApp | https://developers.google.com/apps-script/reference/mail/mail-app |
| UrlFetchApp | https://developers.google.com/apps-script/reference/url-fetch/url-fetch-app |
| HtmlService | https://developers.google.com/apps-script/guides/html |
| CacheService | https://developers.google.com/apps-script/reference/cache/cache-service |
| PropertiesService | https://developers.google.com/apps-script/reference/properties |
| 觸發器指南 | https://developers.google.com/apps-script/guides/triggers |
| 執行限制 | https://developers.google.com/apps-script/guides/services/quotas |
| GAS 範例集 | https://github.com/googleworkspace/apps-script-samples |

## SpreadsheetApp

```javascript
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('工作表1');
var newSheet = ss.insertSheet('新工作表');

// 讀寫資料（批次處理，避免逐格操作）
var data = sheet.getRange('A1:D10').getValues();   // 二維陣列
sheet.getRange('A1:D10').setValues(data);

// 常用 Range 操作
var range = sheet.getRange(row, col, numRows, numCols);
range.getValue() / range.setValue(val)
sheet.getLastRow()    // 最後一筆資料列號
sheet.getLastColumn() // 最後一筆資料欄號
```

## 格式設定

```javascript
range.setFontWeight('bold')
     .setFontSize(12)
     .setFontColor('#1a73e8')
     .setBackground('#e8f5e9')
     .setHorizontalAlignment('center')  // left / center / right
     .setNumberFormat('#,##0.00')       // 千分位+小數
     .setNumberFormat('0.00%')          // 百分比
     .setNumberFormat('yyyy/mm/dd')     // 日期
     .setBorder(top, left, bottom, right, vertical, horizontal,
                color, SpreadsheetApp.BorderStyle.SOLID)
     .merge()                           // 合併儲存格
     .setWrap(true);                    // 自動換行

sheet.setColumnWidth(col, pixels);
sheet.setRowHeight(row, pixels);
```

## 條件式格式化

```javascript
sheet.clearConditionalFormatRules();  // 先清除舊規則

var rule = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberGreaterThanOrEqualTo(100000)
  .setBackground('#c8e6c9')
  .setFontColor('#1b5e20')
  .setRanges([sheet.getRange('C2:C13')])
  .build();

sheet.setConditionalFormatRules([rule]);

// 常用條件方法
.whenNumberGreaterThan(n)
.whenNumberLessThan(n)
.whenNumberBetween(a, b)
.whenTextContains(text)
.whenFormulaSatisfied('=A1>B1')  // 自訂公式
```

## 觸發器 (Triggers)

```javascript
// 簡單觸發器（不需授權）
function onOpen(e) {}   // 開啟試算表
function onEdit(e) {    // 編輯儲存格
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  if (sheet.getName() !== '目標表') return;
  if (range.getColumn() !== 3) return;
  if (range.getRow() < 2) return;
  // 處理邏輯
}

// 時間觸發器（需授權，用程式建立）
ScriptApp.newTrigger('函數名')
  .timeBased()
  .everyHours(1)   // 或 .everyDays(1) / .atHour(8)
  .create();

// 刪除所有觸發器
ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
```

## 圖表 (Charts)

```javascript
var chart = sheet.newChart()
  .setChartType(Charts.ChartType.COLUMN)  // BAR / LINE / PIE / AREA / SCATTER
  .addRange(sheet.getRange('A1:E7'))
  .setPosition(2, 7, 0, 0)               // (列, 欄, x偏移, y偏移)
  .setOption('title', '圖表標題')
  .setOption('width', 700)
  .setOption('height', 400)
  .setOption('colors', ['#1a73e8', '#ea4335'])
  .setOption('legend', {position: 'bottom'})
  .setOption('isStacked', true)           // 堆疊
  .setOption('curveType', 'function')     // 平滑折線
  .setOption('pieHole', 0.4)             // 環形圖
  .build();
sheet.insertChart(chart);

// 更新圖表
var existing = sheet.getCharts()[0];
sheet.updateChart(existing.modify().setOption('title', '新標題').build());

// 匯出圖表為 Blob（插入 Slides）
var blob = chart.getBlob();
```

## Google Slides (SlidesApp)

```javascript
// 物件層級：SlidesApp → Presentation → Slide → Shape/Table/Image
var ppt = SlidesApp.create('簡報名稱');
var slide = ppt.getSlides()[0];
var newSlide = ppt.appendSlide(SlidesApp.PredefinedLayout.BLANK);

// 插入元素（單位：點 points，1吋=72點；標準簡報 720×405pt）
slide.insertTextBox('文字', x, y, width, height);
slide.insertTable(rows, cols);
slide.insertImage(blob);           // blob 來自 getBlob()
slide.getBackground().setSolidFill('#1a73e8');

// 取得簡報 URL
Logger.log(ppt.getUrl());
```

## DriveApp

```javascript
var folder = DriveApp.getFolderById('folder_id');
var file = DriveApp.getFileById('file_id');
var files = folder.getFiles();          // FileIterator
while (files.hasNext()) {
  var f = files.next();
  Logger.log(f.getName());
}
folder.createFile('檔名.txt', '內容', MimeType.PLAIN_TEXT);
file.makeCopy('新檔名', folder);
```

## MailApp / GmailApp

```javascript
// 簡單寄信
MailApp.sendEmail('to@example.com', '主旨', '純文字內文');

// 含 HTML 與附件
GmailApp.sendEmail('to@example.com', '主旨', '', {
  htmlBody: '<h1>標題</h1><p>內文</p>',
  attachments: [DriveApp.getFileById('id').getBlob()],
  cc: 'cc@example.com'
});
```

## UrlFetchApp（呼叫外部 API）

```javascript
// GET 請求
var response = UrlFetchApp.fetch('https://api.example.com/data');
var json = JSON.parse(response.getContentText());

// POST 請求（含 JSON body）
var options = {
  method: 'post',
  contentType: 'application/json',
  headers: { 'Authorization': 'Bearer YOUR_TOKEN' },
  payload: JSON.stringify({ key: 'value' }),
  muteHttpExceptions: true   // 不讓 4xx/5xx 直接拋例外
};
var res = UrlFetchApp.fetch('https://api.example.com/post', options);
Logger.log(res.getResponseCode());   // 200 / 400 / 500...

// 批次請求（同時送出多個，更快）
var requests = [
  { url: 'https://api.example.com/a' },
  { url: 'https://api.example.com/b' }
];
var responses = UrlFetchApp.fetchAll(requests);
responses.forEach(r => Logger.log(r.getContentText()));
```

---

## HtmlService（自訂選單 / 對話框 / 側邊欄）

### 建立自訂選單

```javascript
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('我的工具')
    .addItem('執行分析', 'runAnalysis')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('設定')
        .addItem('開啟側邊欄', 'showSidebar')
    )
    .addToUi();
}
```

### Modal 對話框（會阻擋操作）

```javascript
function showDialog() {
  var html = HtmlService.createHtmlOutput('<p>處理完成！</p><button onclick="google.script.host.close()">關閉</button>')
    .setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, '完成通知');
}
```

### 側邊欄（不阻擋操作）

```javascript
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')  // sidebar.html
    .setTitle('輔助面板');
  SpreadsheetApp.getUi().showSidebar(html);
}

// sidebar.html 中呼叫 server 端函式：
// google.script.run.withSuccessHandler(fn).myServerFunction(arg);
```

### 提示框（快速確認）

```javascript
var ui = SpreadsheetApp.getUi();
var result = ui.alert('確認刪除？', ui.ButtonSet.YES_NO);
if (result === ui.Button.YES) { /* 執行刪除 */ }

var name = ui.prompt('輸入名稱').getResponseText();
```

---

## PropertiesService（持久化設定儲存）

> 用途：跨執行保存設定值（API key、上次執行時間等），比 Utilities 快，比 Sheet 讀寫省 quota。

```javascript
// 三種範圍
var scriptProps = PropertiesService.getScriptProperties();   // 所有使用者共用
var userProps   = PropertiesService.getUserProperties();     // 目前使用者
var docProps    = PropertiesService.getDocumentProperties(); // 目前文件

// 讀寫
scriptProps.setProperty('API_KEY', 'abc123');
var key = scriptProps.getProperty('API_KEY');  // 'abc123' or null

// 批次設定
scriptProps.setProperties({ 'KEY1': 'v1', 'KEY2': 'v2' });

// 刪除
scriptProps.deleteProperty('KEY1');
scriptProps.deleteAllProperties();
```

---

## CacheService（短效快取，減少重複請求）

> 最長快取 6 小時（21600 秒）。適合快取外部 API 回應、大量運算結果。

```javascript
var cache = CacheService.getScriptCache();

// 寫入（秒為單位，最長 21600）
cache.put('key', JSON.stringify(data), 300);   // 快取 5 分鐘

// 批次寫入
cache.putAll({ 'k1': 'v1', 'k2': 'v2' }, 600);

// 讀取
var cached = cache.get('key');
if (cached) {
  return JSON.parse(cached);   // 命中快取
}
// 否則重新取得並寫入快取

// 清除
cache.remove('key');
cache.removeAll(['k1', 'k2']);
```

---

## 資料驗證（下拉選單）

```javascript
// 下拉選項清單
var rule = SpreadsheetApp.newDataValidation()
  .requireValueInList(['待處理', '進行中', '已完成'], true)
  .setAllowInvalid(false)
  .build();
sheet.getRange('B2:B100').setDataValidation(rule);

// 從某欄取值作為選項
var rule2 = SpreadsheetApp.newDataValidation()
  .requireValueInRange(sheet.getRange('E2:E10'))
  .build();

// 數字範圍驗證
var rule3 = SpreadsheetApp.newDataValidation()
  .requireNumberBetween(1, 100)
  .setHelpText('請輸入 1~100 的整數')
  .build();
```

---

## Utilities（工具函式）

```javascript
// 日期格式化
Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

// Base64 編碼 / 解碼
var encoded = Utilities.base64Encode('Hello');
var decoded = Utilities.newBlob(Utilities.base64Decode(encoded)).getDataAsString();

// 延遲執行（注意：會消耗執行時間）
Utilities.sleep(1000);   // 等待 1 秒，避免 API rate limit

// MD5 / SHA 雜湊
var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, 'text');
var hash = bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
```

---

## FormApp（Google 表單）

```javascript
// 建立表單
var form = FormApp.create('問卷標題');
form.setDescription('填寫說明');
form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

// 新增題目
form.addTextItem().setTitle('姓名').setRequired(true);
form.addMultipleChoiceItem()
  .setTitle('性別')
  .setChoiceValues(['男', '女', '其他']);
form.addCheckboxItem()
  .setTitle('興趣')
  .setChoiceValues(['閱讀', '旅遊', '音樂']);
form.addScaleItem().setTitle('滿意度').setBounds(1, 5);

// 取得表單 URL
Logger.log(form.getPublishedUrl());
Logger.log(form.getEditUrl());
```

---

## 資料處理常用模式

```javascript
// filter + sort
var result = data
  .filter(row => row[2] > 5000)
  .sort((a, b) => b[2] - a[2]);  // 降冪

// 分組統計
var stats = {};
data.forEach(row => {
  var key = row[0];
  if (!stats[key]) stats[key] = { count: 0, total: 0 };
  stats[key].count++;
  stats[key].total += row[2];
});
```

## 執行限制與最佳實踐

| 限制 | 數值 |
|------|------|
| 腳本執行時間 | 6 分鐘 |
| 跨服務（Sheets+Slides）更容易達到 | 注意複雜操作 |
| 觸發器每日執行次數 | 20 次（免費版） |

```javascript
// ✅ 批次讀寫（快）
var all = sheet.getDataRange().getValues();
// ❌ 逐格操作（慢，易超時）
for (var i = 1; i <= 100; i++) sheet.getRange(i, 1).getValue();

// 防並發（LockService）
var lock = LockService.getScriptLock();
lock.waitLock(30000);
try { /* 操作 */ } finally { lock.releaseLock(); }

// 錯誤處理標準模板
function myFunc() {
  try {
    // 主邏輯
  } catch (e) {
    Logger.log('錯誤：' + e.toString());
    SpreadsheetApp.getUi().alert('執行失敗：' + e.message);
  }
}
```

## 常見錯誤

| 錯誤 | 原因 | 解法 |
|------|------|------|
| `Exception: Service Spreadsheets failed` | 逐格讀寫太多次 | 改用 `getValues()`批次 |
| `Exceeded maximum execution time` | 超過 6 分鐘 | 拆分任務或用觸發器分批執行 |
| 條件式格式規則累積 | 每次執行都疊加 | 先呼叫 `clearConditionalFormatRules()` |
| `insertTextBox` 位置不對 | 單位是 points 非像素 | 標準簡報寬 720pt = 10 吋 |
| 圖表無法直接連結到 Slides | API 限制 | 用 `getBlob()` 轉圖片再插入 |
