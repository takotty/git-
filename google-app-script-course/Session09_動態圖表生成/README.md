# Session 9：動態圖表生成

> 📅 日期：115/05/30（六）09:00~12:00 ｜ 講師：林冠廷

---

## 🎯 課程目標

1. 學會用 Apps Script 建立柱狀圖、折線圖、圓餅圖
2. 了解圖表選項設定（顏色、標題、軸標籤等）
3. 實作動態圖表：數據變更時圖表自動更新
4. 一鍵生成多種圖表

---

## 📖 步驟教學

### Step 1：初始化資料

1. 建立 `GAS課程練習_Session09`，貼入 `code.gs`
2. 執行「📦 初始化圖表資料」
3. 會建立「圖表資料」和「月度趨勢」兩個工作表

### Step 2：逐一建立圖表

1. 📊 建立柱狀圖 → 查看部門 Q1~Q4 對比
2. 📊 建立堆疊柱狀圖 → 查看堆疊效果差異
3. 📈 建立折線圖 → 查看月度趨勢
4. 🥧 建立圓餅圖 → 查看部門年度佔比

### Step 3：動態更新

1. 執行「🔄 模擬資料更新」
2. 觀察柱狀圖的 Q4 數據自動變化
3. 也可以手動修改試算表數值，圖表即時反映

---

## 💻 程式碼解說

### 圖表建立基本結構

```javascript
var chart = sheet.newChart()
  .setChartType(Charts.ChartType.COLUMN)  // 圖表類型
  .addRange(資料範圍)                      // 資料來源
  .setPosition(列, 欄, x偏移, y偏移)       // 位置
  .setOption("title", "標題")              // 選項
  .setOption("width", 700)
  .setOption("height", 400)
  .build();

sheet.insertChart(chart);
```

### 常用圖表類型

| 類型 | 常數 | 用途 |
|------|------|------|
| 柱狀圖 | `Charts.ChartType.COLUMN` | 分類比較 |
| 長條圖 | `Charts.ChartType.BAR` | 水平比較 |
| 折線圖 | `Charts.ChartType.LINE` | 趨勢變化 |
| 圓餅圖 | `Charts.ChartType.PIE` | 佔比分析 |
| 面積圖 | `Charts.ChartType.AREA` | 堆疊趨勢 |
| 散佈圖 | `Charts.ChartType.SCATTER` | 相關性 |

### 常用選項設定

| 選項 | 說明 | 範例 |
|------|------|------|
| `title` | 圖表標題 | `"月度營收"` |
| `width` / `height` | 尺寸 | `700` / `400` |
| `colors` | 自訂顏色陣列 | `["#1a73e8", "#ea4335"]` |
| `legend` | 圖例位置 | `{position: "bottom"}` |
| `isStacked` | 是否堆疊 | `true` / `false` |
| `curveType` | 曲線類型 | `"function"` = 平滑 |
| `pieHole` | 環形圖 | `0.4` = 40% 空心 |

---

## 🏋️ 延伸練習

### 練習 1：組合圖
嘗試建立柱狀圖 + 折線圖的組合圖表。

### 練習 2：圖表匯出
使用 `chart.getBlob()` 將圖表匯出為 PNG 圖片。

### 練習 3：儀表板
在一個工作表中排列多個小型圖表，建立簡易儀表板。

---

## ❓ 常見問題 FAQ

### Q：圖表建立後還能修改嗎？
**A**：可以。取得圖表後用 `chart.modify()` 修改選項，再用 `sheet.updateChart()` 更新。

### Q：為什麼圖表顯示錯誤？
**A**：常見原因：資料範圍不正確、含有空白或文字混在數字中。確認資料範圍正確。
