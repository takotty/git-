# Session 3：SpreadsheetApp 基本操作

> 📅 日期：115/05/09（六）09:00~12:00 ｜ 講師：林冠廷

---

## 🎯 課程目標

1. 熟悉 SpreadsheetApp 的常用方法
2. 學會用程式建立、管理工作表
3. 掌握儲存格讀取與寫入操作
4. 設定時間觸發器，實現自動化

---

## 📋 前置準備

- 已完成 Session 1、2 的學習
- Google 帳號已授權 Apps Script 使用

---

## 📖 步驟教學

### Step 1：建立試算表與匯入資料

1. 建立新試算表 `GAS課程練習_Session03`
2. 將工作表命名為 `員工資料`
3. 匯入 `sample_data.csv` 到「員工資料」工作表

### Step 2：貼上程式碼並執行

1. 開啟 Apps Script 編輯器
2. 貼入 `code.gs` 的全部內容
3. 儲存後，先執行 `初始化員工資料`（或手動匯入 CSV）
4. 重新整理頁面，使用選單中的各項功能

### Step 3：體驗自動建立月報表

1. 點選「📚 Session 3 工具」>「📅 建立當月報表」
2. 觀察系統自動建立了名為「2026年4月報表」的新工作表
3. 查看標題格式、數字格式、日期格式是否自動套用

### Step 4：設定觸發器

1. 點選「⏰ 設定每日觸發器」
2. 在 Apps Script 編輯器左側，點選「觸發條件」圖示查看已設定的觸發器
3. 可以點選「🗑️ 刪除所有觸發器」清理

---

## 💻 程式碼解說

### SpreadsheetApp 物件層級

```
SpreadsheetApp（服務）
  └─ Spreadsheet（試算表檔案）
       └─ Sheet（工作表/分頁）
            └─ Range（儲存格範圍）
                 └─ getValue / setValue（讀寫值）
```

### 常用方法速查

| 方法 | 說明 | 範例 |
|------|------|------|
| `getActiveSpreadsheet()` | 取得目前試算表 | `var ss = SpreadsheetApp.getActiveSpreadsheet()` |
| `getSheetByName(name)` | 依名稱取得工作表 | `ss.getSheetByName("員工資料")` |
| `insertSheet(name)` | 新增工作表 | `ss.insertSheet("新分頁")` |
| `getRange(a1)` | 取得儲存格範圍 | `sheet.getRange("A1:C10")` |
| `getValue()` | 讀取單一值 | `range.getValue()` |
| `getValues()` | 讀取範圍值（二維陣列）| `range.getValues()` |
| `setValue(val)` | 寫入單一值 | `range.setValue("Hello")` |
| `setValues(arr)` | 寫入範圍值 | `range.setValues([[1,2],[3,4]])` |

### 觸發器類型

| 類型 | 說明 | 範例 |
|------|------|------|
| `onOpen()` | 開啟試算表時 | 建立選單 |
| `onEdit(e)` | 編輯儲存格時 | 自動驗證 |
| 時間觸發 | 定時執行 | 每日報告 |

---

## 🏋️ 延伸練習

### 練習 1：自動建立週報表
修改 `自動建立月報表`，改為建立「週報表」，包含欄位：日期、事項、負責人、進度(%)。

### 練習 2：工作表複製
撰寫一個函數，將「員工資料」工作表複製一份，命名為「員工資料_備份_日期」。

### 練習 3：跨工作表複製
撰寫函數，將 A 工作表的資料複製到 B 工作表的指定位置。

---

## ❓ 常見問題 FAQ

### Q：`getRange("A1")` 和 `getRange(1, 1)` 有什麼差別？
**A**：兩者都取得 A1 儲存格。`getRange("A1")` 使用 A1 表示法，`getRange(1, 1)` 使用列欄數字（1-indexed）。數字表示法在迴圈中更方便。

### Q：`getValue()` 和 `getValues()` 差在哪？
**A**：`getValue()` 回傳單一值；`getValues()` 回傳二維陣列。多個儲存格請用 `getValues()`，效能較好。

### Q：觸發器設定後為什麼沒有執行？
**A**：時間觸發器在設定的時段內「某個時間點」執行（非精確時間）。例如設定 9 點，可能在 9:00~10:00 之間執行。
