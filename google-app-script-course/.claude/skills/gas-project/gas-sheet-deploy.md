---
name: gas-sheet-deploy
description: Use when deploying Google Apps Script projects to Google Sheets — single folder or batch across all session folders. Trigger on "部署 GAS", "把 code.gs 上傳到 Google Sheet", "建立試算表並綁定 script", "匯入 CSV 到試算表", "部署 Session N", or "批次部署所有 session". Covers creating spreadsheets, binding Apps Script, uploading code.gs, importing CSV data, and versioning.
---

# Google Sheet + Apps Script 部署流程

## clasp 首次安裝流程

> 若已安裝且登入過，可跳過此節直接到「前置需求」。

### 1. 確認 Node.js 版本（需 v22+）

```bash
node -v
```

若版本過低，升級：
```bash
npm install -g npm
npx n latest
```

### 2. 安裝 clasp

```bash
npm install -g @google/clasp
clasp --version   # 確認安裝成功
```

### 3. 啟用 Google Apps Script API

前往 [https://script.google.com/home/usersettings](https://script.google.com/home/usersettings)，開啟「Google Apps Script API」。

> 若未開啟，`clasp push` 會報 403 錯誤。

### 4. 登入 Google 帳號

```bash
clasp login
```

瀏覽器會自動開啟授權頁面，選擇 Google 帳號並允許存取。  
成功後會在 `~/.clasprc.json` 儲存 token。

**無瀏覽器環境（headless）：**
```bash
clasp login --no-localhost
```
手動複製授權 URL，貼到瀏覽器完成授權，再把回傳的 code 貼回終端機。

### 5. 確認登入狀態

```bash
cat ~/.clasprc.json | python3 -c \
  "import json,sys; d=json.load(sys.stdin); print('clasp OK:', list(d['tokens'].keys()))"
```

---

## 前置需求

```bash
# 確認 clasp 已登入
clasp --version
cat ~/.clasprc.json | python3 -c \
  "import json,sys; d=json.load(sys.stdin); print('clasp OK:', list(d['tokens'].keys()))"
```

**若 token 過期：**
```bash
clasp login
```

> ⚠️ **Sheets 寫入權限（CSV 匯入）**：clasp 預設 token 含 `drive` scope，但不一定含 `spreadsheets` scope。
> 若 CSV 匯入出現 403，需用自訂 OAuth client 重新登入：
> ```bash
> clasp login --creds client_secret.json
> ```
> `client_secret.json` 需在 [GCP Console](https://console.cloud.google.com/) 建立，啟用 Sheets API 並加入 `spreadsheets` scope。

**本地資料夾結構（每個 session）：**
```
Session0X_主題/
  code.gs            ← 必須
  sample_data.csv    ← 選填，有則自動匯入
  appsscript.json    ← 若無則自動建立
  .clasp.json        ← 部署後自動產生（已部署的標記）
```

---

## 單一資料夾部署

```bash
FOLDER="/path/to/session/folder"
NAME=$(basename "$FOLDER")
TOKEN=$(python3 -c \
  "import json; d=json.load(open('$HOME/.clasprc.json')); print(d['tokens']['default']['access_token'])")

# Step 1: 建立試算表（Drive API，只需 drive scope）
SHEET_ID=$(curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d "{\"name\": \"$NAME\", \"mimeType\": \"application/vnd.google-apps.spreadsheet\"}" \
  "https://www.googleapis.com/drive/v3/files" | \
  python3 -c "import json,sys; d=json.load(sys.stdin); print(d.get('id',''))")
echo "Sheet ID: $SHEET_ID"

# Step 2: 移至對應 Drive 資料夾（搜尋同名資料夾）
DRIVE_FOLDER_ID=$(curl -s -G \
  -H "Authorization: Bearer $TOKEN" \
  --data-urlencode "q=name = '${NAME}' and mimeType = 'application/vnd.google-apps.folder'" \
  --data-urlencode "pageSize=1" \
  "https://www.googleapis.com/drive/v3/files" | \
  python3 -c "import json,sys; files=json.load(sys.stdin).get('files',[]); print(files[0]['id'] if files else '')")
[ -n "$DRIVE_FOLDER_ID" ] && curl -s -X PATCH \
  -H "Authorization: Bearer $TOKEN" \
  "https://www.googleapis.com/drive/v3/files/${SHEET_ID}?addParents=${DRIVE_FOLDER_ID}" \
  > /dev/null && echo "移至 Drive 資料夾 ✅"

# Step 3: 建立 Apps Script 並綁定試算表
# ⚠️ clasp create --parentId 是 folder，不能綁定現有 sheet，需直接呼叫 Script API
SCRIPT_ID=$(curl -s -X POST \
  -H "Authorization: Bearer $TOKEN" \
  -H "Content-Type: application/json" \
  -d "{\"title\": \"${NAME}_script\", \"parentId\": \"$SHEET_ID\"}" \
  "https://script.googleapis.com/v1/projects" | \
  python3 -c "import json,sys; d=json.load(sys.stdin); print(d.get('scriptId','ERROR'))")
echo "Script ID: $SCRIPT_ID"

# Step 4: 建立 appsscript.json（若無）並推送 code.gs
[ ! -f "$FOLDER/appsscript.json" ] && \
  echo '{"timeZone":"Asia/Taipei","dependencies":{},"exceptionLogging":"STACKDRIVER","runtimeVersion":"V8"}' \
  > "$FOLDER/appsscript.json"

printf '{"scriptId":"%s","rootDir":""}' "$SCRIPT_ID" > "$FOLDER/.clasp.json"
cd "$FOLDER" && clasp push --force 2>&1 | tail -1

# Step 5: 匯入 CSV（若有，需 spreadsheets scope）
if [ -f "$FOLDER/sample_data.csv" ]; then
  PAYLOAD=$(python3 -c "
import csv, json
with open('$FOLDER/sample_data.csv', encoding='utf-8') as f:
    rows = list(csv.reader(f))
print(json.dumps({'values': rows}))")
  curl -s -X PUT \
    -H "Authorization: Bearer $TOKEN" \
    -H "Content-Type: application/json" \
    -d "$PAYLOAD" \
    "https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1%21A1?valueInputOption=USER_ENTERED" | \
    python3 -c "import json,sys; d=json.load(sys.stdin); print(f'匯入 {d.get(\"updatedRows\",\"?\")}} 筆 ✅' if 'updatedRows' in d else '❌ CSV 匯入失敗（可能缺 spreadsheets scope）：' + d.get('error',{}).get('message',''))"
fi

# Step 6: 建立版本 + 部署
curl -s -X POST -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
  -d '{"description":"v1"}' \
  "https://script.googleapis.com/v1/projects/${SCRIPT_ID}/versions" > /dev/null
curl -s -X POST -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
  -d '{"versionNumber":1,"manifestFileName":"appsscript","description":"API executable"}' \
  "https://script.googleapis.com/v1/projects/${SCRIPT_ID}/deployments" > /dev/null

echo "✅ 完成 | https://docs.google.com/spreadsheets/d/$SHEET_ID"
echo "   Script: https://script.google.com/d/$SCRIPT_ID/edit"
```

---

## 批次部署（所有未部署的 session）

自動跳過已有 `.clasp.json` 的資料夾（冪等，可重複執行）：

```bash
TOKEN=$(python3 -c \
  "import json; d=json.load(open('$HOME/.clasprc.json')); print(d['tokens']['default']['access_token'])")
BASE="/Users/ray-mac/Documents/Claude/Claude_資源/Claude/google app script"

find "$BASE" -name "code.gs" | sort | while read GS_FILE; do
  FOLDER=$(dirname "$GS_FILE")

  # 已部署則跳過
  [ -f "$FOLDER/.clasp.json" ] && echo "⏭ 已部署: $(basename $FOLDER)" && continue

  NAME=$(basename "$FOLDER")
  echo "▶ 部署: $NAME"

  # Step 1: 建立試算表（Drive API）
  SHEET_ID=$(curl -s -X POST \
    -H "Authorization: Bearer $TOKEN" \
    -H "Content-Type: application/json" \
    -d "{\"name\": \"$NAME\", \"mimeType\": \"application/vnd.google-apps.spreadsheet\"}" \
    "https://www.googleapis.com/drive/v3/files" | \
    python3 -c "import json,sys; d=json.load(sys.stdin); print(d.get('id',''))")
  [ -z "$SHEET_ID" ] && echo "  ❌ 試算表建立失敗" && continue

  # Step 2: 移至 Drive 資料夾
  DRIVE_FOLDER_ID=$(curl -s -G \
    -H "Authorization: Bearer $TOKEN" \
    --data-urlencode "q=name = '${NAME}' and mimeType = 'application/vnd.google-apps.folder'" \
    --data-urlencode "pageSize=1" \
    "https://www.googleapis.com/drive/v3/files" | \
    python3 -c "import json,sys; files=json.load(sys.stdin).get('files',[]); print(files[0]['id'] if files else '')")
  [ -n "$DRIVE_FOLDER_ID" ] && curl -s -X PATCH \
    -H "Authorization: Bearer $TOKEN" \
    "https://www.googleapis.com/drive/v3/files/${SHEET_ID}?addParents=${DRIVE_FOLDER_ID}" \
    > /dev/null && echo "  ✅ 移至 Drive 資料夾"

  # Step 3: 建立 Apps Script 並綁定試算表
  SCRIPT_ID=$(curl -s -X POST \
    -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
    -d "{\"title\": \"${NAME}_script\", \"parentId\": \"$SHEET_ID\"}" \
    "https://script.googleapis.com/v1/projects" | \
    python3 -c "import json,sys; d=json.load(sys.stdin); print(d.get('scriptId',''))")
  [ -z "$SCRIPT_ID" ] && echo "  ❌ Script 建立失敗" && continue

  # Step 4: 推送 code.gs
  [ ! -f "$FOLDER/appsscript.json" ] && \
    echo '{"timeZone":"Asia/Taipei","dependencies":{},"exceptionLogging":"STACKDRIVER","runtimeVersion":"V8"}' \
    > "$FOLDER/appsscript.json"
  printf '{"scriptId":"%s","rootDir":""}' "$SCRIPT_ID" > "$FOLDER/.clasp.json"
  cd "$FOLDER" && clasp push --force 2>/dev/null | tail -1

  # Step 5: 匯入 CSV（需 spreadsheets scope）
  if [ -f "$FOLDER/sample_data.csv" ]; then
    PAYLOAD=$(python3 -c "
import csv, json
with open('$FOLDER/sample_data.csv', encoding='utf-8') as f:
    rows = list(csv.reader(f))
print(json.dumps({'values': rows}))")
    ROWS=$(curl -s -X PUT \
      -H "Authorization: Bearer $TOKEN" \
      -H "Content-Type: application/json" \
      -d "$PAYLOAD" \
      "https://sheets.googleapis.com/v4/spreadsheets/${SHEET_ID}/values/Sheet1%21A1?valueInputOption=USER_ENTERED" | \
      python3 -c "import json,sys; d=json.load(sys.stdin); print(d.get('updatedRows','❌ 失敗'))")
    echo "  ✅ CSV 匯入 $ROWS 筆"
  fi

  # Step 6: 建立版本 + 部署
  curl -s -X POST -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
    -d '{"description":"v1"}' \
    "https://script.googleapis.com/v1/projects/${SCRIPT_ID}/versions" > /dev/null
  curl -s -X POST -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
    -d '{"versionNumber":1,"manifestFileName":"appsscript","description":"API executable"}' \
    "https://script.googleapis.com/v1/projects/${SCRIPT_ID}/deployments" > /dev/null

  echo "  ✅ 完成 | Sheet: $SHEET_ID"
  echo ""
done
```

---

## 驗證部署狀態

```bash
BASE="/Users/ray-mac/Documents/Claude/Claude_資源/Claude/google app script"
TOTAL=$(find "$BASE" -name "code.gs" | wc -l | tr -d ' ')
DEPLOYED=$(find "$BASE" -name ".clasp.json" | wc -l | tr -d ' ')
echo "部署進度：$DEPLOYED / $TOTAL"

# 列出未部署的
find "$BASE" -name "code.gs" | while read f; do
  d=$(dirname "$f")
  [ ! -f "$d/.clasp.json" ] && echo "未部署: $(basename $d)"
done
```

---

## 執行 Apps Script 函式

**CLI 限制：** container-bound script 無法透過 `clasp run` 或 `scripts.run` API 執行。

**從試算表執行：**
1. 開啟試算表 → **Extensions → Apps Script**
2. 選擇函式 → 按 ▶️

---

## 常見問題

| 問題 | 原因 | 解法 |
|------|------|------|
| Step 1 回傳空 Sheet ID | Drive API 403 / token 過期 | `clasp login` 重新授權 |
| `clasp push` 報錯 manifest | 缺 appsscript.json | Step 4 自動建立，或手動新增 |
| CSV 匯入 403 | clasp token 缺 `spreadsheets` scope | `clasp login --creds client_secret.json`（需含 spreadsheets scope） |
| Script ID 回傳 ERROR | Script API 未啟用 | 前往 https://script.google.com/home/usersettings 開啟 API |
| `clasp run` 找不到函式 | Container-bound 限制 | 從試算表 Extensions → Apps Script 執行 |
| token 過期 | 超過 1 小時 | `clasp login` |
| Drive 資料夾找不到 | 名稱不完全符合 | 手動指定 `DRIVE_FOLDER_ID=<id>` |
| `clasp create` 建了新試算表 | `--parentId` 是 folder，非現有 sheet | 改用 Step 3 直接呼叫 Script API |
