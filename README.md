# AI 說明文件 — googleGASapp（Google Apps Script）

> 本文件供 AI / 開發者快速理解此 Google Apps Script 專案的結構、部署方式與注意事項。

## 1. 專案概述

本專案為「鬼磕頭」服務的 Google Apps Script（GAS）後端，提供：
- **諮詢表單 Web App**（`諮詢表單.gs`）：接收網站表單、寫入試算表、發送 OTP 驗證、LINE 通知。
- **LINE Bot Webhook**（`LINE.gs`）：處理 LINE 訊息/Postback、案件查詢、訓練管理、小遊戲。
- **互動小遊戲模組**（`GameModule.gs`）：新好友開場遊戲、生日登記、遊戲選擇記錄。

## 2. 部署方式（重要）

- `諮詢表單.gs` 與 `LINE.gs` **不會同時部署**。
- 部署 LINE Bot 時，`諮詢表單.gs` 會**全部註解**，避免 `doPost` 等全域函式衝突。
- 因此兩檔的重複定義（`doPost`、`json_`、`props_`、`targetGroup_` 等）在實際部署中不會同時生效，不構成衝突。
- `GameModule.gs` 一律隨 LINE 部署載入（提供小遊戲與生日功能）。

## 3. 檔案說明

| 檔案 | 角色 |
|------|------|
| `諮詢表單.gs` | Web App 入口（`doPost` 處理 `sendOtp` / `submit`），諮詢表單寫入、reCAPTCHA、Email OTP、LINE 通知 |
| `LINE.gs` | LINE Webhook 入口（`doPost` 處理 `relay` / `events`），指令、案件、訓練、遊戲路由 |
| `GameModule.gs` | 小遊戲卡片、生日選擇器、遊戲選擇記錄、歡迎回來 |
| `TrainModule.gs` | 訓練與庫存管理（已從 LINE.gs 切出） |
| `ConsumableModule.gs` | 耗材庫存與盤點管理（新增/進貨/消耗/盤點/查詢） |

## 4. 試算表工作表

| 工作表名稱 | 用途 |
|------------|------|
| `諮詢表單內容` | 諮詢案件主表（STATUS_COL=12 狀態欄） |
| `LINE_Group` | LINE 群組白名單/角色 |
| `LINE_User` | LINE 使用者資料（含角色、啟用、綁定、生日、遊戲選擇） |
| `遊戲設定` | 小遊戲設定（由 GameModule 自動建立） |
| `SOP與庫存表` | 訓練 SOP 單元與庫存 |
| `訓練紀錄` | 徒弟訓練進度 |

## 5. Script Properties（屬性）

| 屬性 | 用途 |
|------|------|
| `LINE_TOKEN` | LINE Messaging API Token |
| `ADMIN_EMAIL` | 管理員 Email（備援通知） |
| `RELAY_SECRET` | 中繼驗證密鑰 |
| `GROUP_ID` | 備用群組 ID |
| `RECAPTCHA_SECRET` | reCAPTCHA 密鑰（僅諮詢表單） |

## 6. 主要流程

### 諮詢表單
1. 前端送出 `action=sendOtp` → 寄送 6 位數驗證碼（Email）。
2. 前端送出 `action=submit` → 驗證 OTP → 寫入 `諮詢表單內容` → 寄送確認信 → LINE 通知群組。

### LINE Bot
1. `doPost` 解析 `events` → `handleLine_`。
2. 依群組/個人、是否管理員，路由到 `guestEvent_` 或 `handleAllowedEvent_`。
3. Postback 先交給 `handleGamePostback_`，再交給 `handleTrainPostback_`。

### 小遊戲
1. 新好友（或 `測試開場白`）→ `handleFollowEvent_` → `getGameCard_`。
2. 使用者選擇選項 → `handleGamePostback_`（`action=game_opt`）→ 記錄選擇 → 詢問生日。
3. 生日選擇 → `action=select_birthday` → `saveUserBirthday_`（防重複修改）。

## 7. 觸發器（setupTriggers）

- `dailyUnclosed_`：每日 09:00，未結案提醒。
- `dailyTrainAlert_`：每日 08:00，備課/庫存提醒。

## 8. 已知注意事項

1. `handleLine_` 中 `follow` 事件處理目前被註解，新好友不會自動收到開場遊戲（僅測試指令可觸發）。
2. `diagPush` 使用以 `U` 開頭的 ID（使用者），非群組；命名為 gid 但實際推播對象為個人。
3. `formatYMD` 使用 `new Date()` 本地時區解析，GAS（UTC）可能造成日期偏移，建議改 `Utilities.formatDate`。
4. `diagPush_game` 內有未使用的 `btns` 死代碼。
5. 訓練庫存「補蓋章」已修正為優先轉換已列印數量，避免重複計數。
6. 新增庫存盤點功能：指令 `庫存盤點` / `盤點 冊名` / `設庫存 單元 列印|蓋章 數量`。
7. 訓練進度指令改為「敕令/領取/完成」；訓練紀錄 ID 採 `YYYYMMDD-XX`；更新進度按鈕支援防呆反灰（已處理單元顯示灰色）。
8. 新增耗材庫存模組：指令 `耗材新增/進貨/消耗/盤點/查詢`，僅限管理員使用。
9. 「啟用」欄位說明：群組或個人設 `N` 都會完全停用（不回覆、不處理），讓啟用欄位有實際意義。
10. 訓練模組：進度指令為「敕令/領取/完成」；訓練紀錄 ID 採 `YYYYMMDD-XX`；更新進度按鈕支援防呆反灰（已處理單元顯示灰色）；補庫存支援輸入數量。
11. 耗材模組：指令 `耗材新增/進貨/消耗/盤點/查詢/管理`；按鈕採 postback 預填輸入；庫存查詢顯示最後更新時間（含年份）。
12. 耗材管理選單：庫存查詢為第一個按鈕；新增耗材不加顏色。

## 9. 維護建議

- 若需同時啟用諮詢表單與 LINE Bot，請拆成兩個獨立 Apps Script 專案，避免 `doPost` 衝突。
- 修改 `LINE_User` 表頭時，GameModule 的動態欄位（`遊戲_`、`生日`）會自動新增，請勿手動刪除。
- 調整遊戲內容請編輯 `遊戲設定` 工作表，不需改程式碼。

---
*本文件由 AI 產生，最後更新：2026-11-07*