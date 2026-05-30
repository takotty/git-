# Session 8：條件式格式化自動化

> 📅 日期：115/05/23（六）13:30~16:30 ｜ 講師：林冠廷

---

## 🎯 課程目標

1. 學會用程式建立條件式格式規則
2. 理解並實作避免重複觸發的技巧
3. 運用多層迴圈生成熱力圖
4. 綜合實作：自動格式化完整報表

---

## 📖 步驟教學

### Step 1：初始化並打下基礎

1. 建立 `GAS課程練習_Session08`，貼入 `code.gs`
2. 執行「📦 初始化業績資料」建立 12 筆業績數據

### Step 2：套用條件式格式

1. 執行「🎨 條件式格式化」
2. 觀察業績和達成率欄位的顏色變化
3. 嘗試手動修改業績數值（C 欄），觀察 onEdit 自動更新

### Step 3：熱力圖

1. 執行「🌡️ 建立熱力圖」
2. 觀察自動生成的紅→黃→綠漸層效果

### Step 4：完整格式化

1. 執行「📋 自動格式化報表」
2. 觀察報表的整體美化效果

---

## 💻 程式碼解說

### 條件式格式規則建立

```javascript
var 規則 = SpreadsheetApp.newConditionalFormatRule()
  .whenNumberGreaterThanOrEqualTo(100000)  // 條件
  .setBackground("#c8e6c9")                // 背景色
  .setFontColor("#1b5e20")                 // 字體色
  .setRanges([sheet.getRange("C2:C13")])   // 套用範圍
  .build();

sheet.setConditionalFormatRules([規則]);
```

### 可用的條件方法

| 方法 | 說明 |
|------|------|
| `whenNumberGreaterThan(n)` | 大於 |
| `whenNumberLessThan(n)` | 小於 |
| `whenNumberBetween(a, b)` | 介於 |
| `whenTextContains(text)` | 包含文字 |
| `whenTextEqualTo(text)` | 等於文字 |
| `whenDateAfter(date)` | 日期之後 |
| `whenFormulaSatisfied(formula)` | 自訂公式 |

### 避免重複觸發的 4 大技巧

1. **限制工作表**：`if (工作表名 !== "目標表") return;`
2. **限制欄位**：`if (range.getColumn() !== 3) return;`
3. **跳過標題列**：`if (range.getRow() < 2) return;`
4. **使用 LockService**：防止並發執行

---

## 🏋️ 延伸練習

### 練習 1：自訂公式條件
使用 `whenFormulaSatisfied()` 建立跨欄位的條件格式（例如：B 欄 > C 欄時標紅）。

### 練習 2：漸層色支援更多色系
修改 `取得漸層顏色` 函數，支援藍紫色系的漸層。

### 練習 3：onEdit 觸發器進階
在 onEdit 中加入驗證：若使用者輸入負數業績，自動清除並顯示警告。

---

## ❓ 常見問題 FAQ

### Q：條件式格式規則會累積嗎？
**A**：會！每次執行 `setConditionalFormatRules` 會替換全部規則。記得先呼叫 `clearConditionalFormatRules()` 清除舊規則。

### Q：onEdit 觸發器能使用 UI 對話框嗎？
**A**：不行。`onEdit` 是簡單觸發器，不能呼叫 `SpreadsheetApp.getUi()`。只能用 `Logger.log` 記錄。

### Q：LockService 的鎖定時間有上限嗎？
**A**：`waitLock()` 最長等待 180 秒。鎖定自動在腳本結束時釋放，但建議在 `finally` 區塊手動釋放。
