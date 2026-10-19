# ChatGPT Excel AHK Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 建立一個以人工單鍵控制、透過 ChatGPT 原生鍵盤快捷鍵交換問題與回答，並把結果逐列寫入 Excel 的 AutoHotkey v2 工具。

**Architecture:** `ChatGPT_Excel_助手.ahk` 負責啟動流程、Excel COM、前景視窗檢查與快捷鍵編排；`src/ChatGPTExcelCore.ahk` 保存可獨立測試的文字、預覽與狀態決策；`src/MouseStatus.ahk` 專責跟隨滑鼠的非焦點提示 GUI。所有動作由使用者按鍵觸發，不輪詢 ChatGPT、不操作滑鼠座標。

**Tech Stack:** AutoHotkey v2、Windows COM、Microsoft Excel 桌面版、Google Chrome、Office Open XML 測試活頁簿

**Spec:** `docs/superpowers/specs/2026-10-19-chatgpt-excel-ahk-design.md`

## Global Constraints

- 僅支援 AutoHotkey v2，不加入 v1 相容層。
- `F1` 與 `F2` 永遠有效；`Z/X/C/V/A` 只在啟用狀態攔截單鍵。
- 不使用滑鼠座標、DOM、瀏覽器擴充功能、自動輪詢、自動重試或批次迴圈。
- 只有 Excel 寫入與活頁簿儲存都成功後才移動問題與回答列號。
- 紅色警告不可被後續成功訊息蓋掉，直到使用者按 `V`。
- 啟動或測試程式由使用者自行執行；實作過程不替使用者啟動 AHK、Chrome 或 Excel。

## Review Focus

- `Ctrl+Shift+C` 沒有更新剪貼簿：應警告且不寫入、不儲存、不移動列號；由 Task 3 整合測試案例固定。
- 活頁簿唯讀或儲存失敗：應保留目前列號並提示可重試；由 Task 3 整合測試案例固定。
- 使用者取消檔案或範圍選擇、或選取多格：應安全結束啟動；由 Task 3 啟動驗收固定。
- 前景視窗不是 Chrome：`X/C/A` 均不得送出按鍵；由 Task 3 前景視窗案例固定。
- 問題含多行或超過 120 字：貼上保留完整內容，但預覽限制為 120 字與四行；由 Task 1 單元測試固定。

---

## File Structure

- `ChatGPT_Excel_助手.ahk`：入口、啟動選格、Excel COM、快捷鍵與動作編排。
- `src/ChatGPTExcelCore.ahk`：純函式、提示文案、預覽裁切及列移動判斷。
- `src/MouseStatus.ahk`：跟隨滑鼠提示 GUI 與警告鎖定規則。
- `tests/一鍵測試.ahk`：不開啟 Chrome 或 Excel 的核心邏輯測試。
- `AHK功能測試.xlsx`：人工整合測試資料。
- `測試清單.md`：最短人工驗收流程與预期提示。

### Task 1: 核心文字與流程決策

**Files:**
- Create: `src/ChatGPTExcelCore.ahk`
- Create: `tests/一鍵測試.ahk`

**Interfaces:**
- Produces: `BuildPreview(text, maxChars := 120, maxLines := 4) -> String`
- Produces: `BuildStatus(kind, questionLabel := "", answerLabel := "", nextQuestionLabel := "", preview := "") -> String`
- Produces: `CanAdvance(copyOk, writeOk, saveOk) -> Boolean`
- Produces: `IsChromeProcess(processName) -> Boolean`

- [ ] **Step 1: 先寫核心測試**

在 `tests/一鍵測試.ahk` 加入無外部相依的斷言，涵蓋：120 字截斷、四行截斷、短文字不加省略號、所有既定提示文案、三項成功才允許移列、`chrome.exe` 大小寫處理及非 Chrome 拒絕。

- [ ] **Step 2: 保留預期失敗證據**

不啟動測試檔；以靜態檢查確認測試已引用尚不存在的四個介面，記錄「執行交由使用者」限制。

- [ ] **Step 3: 實作最小核心函式**

在 `src/ChatGPTExcelCore.ahk` 實作四個介面。`BuildPreview` 只裁切顯示文字，不修改實際送出的完整問題；`BuildStatus` 使用規格中的繁體中文固定文案。

- [ ] **Step 4: 靜態核對測試與實作名稱**

使用 `rg` 確認測試引用的函式名稱與實作完全一致，且沒有 `TODO/TBD`。

### Task 2: 滑鼠旁提示元件

**Files:**
- Create: `src/MouseStatus.ahk`
- Modify: `tests/一鍵測試.ahk`

**Interfaces:**
- Consumes: Task 1 `BuildStatus(...)`
- Produces: `MouseStatus.ShowSuccess(text, ttlMs := 2500)`
- Produces: `MouseStatus.ShowInfo(text, sticky := false)`
- Produces: `MouseStatus.ShowWarning(text)`
- Produces: `MouseStatus.ClearWarning()`
- Produces: `MouseStatus.Shutdown()`

- [ ] **Step 1: 先寫提示狀態測試**

加入可在不顯示 GUI 的測試模式下驗證：警告會阻擋成功訊息、`ClearWarning()` 後可再次顯示成功訊息、成功訊息含 2500 ms 到期時間。

- [ ] **Step 2: 實作提示 GUI**

建立無標題列、永遠置頂、不取得焦點的小型 GUI；使用計時器跟隨滑鼠，綠色代表成功、黃色代表資訊、紅色代表警告。公開方法維持上述介面，測試模式只更新内部狀態不建立視窗。

- [ ] **Step 3: 靜態檢查提示生命週期**

確認每個計時器都有停止路徑、`Shutdown()` 可關閉 GUI，且警告鎖定只能由 `ClearWarning()` 解除。

### Task 3: Excel、Chrome 與快捷鍵整合

**Files:**
- Create: `ChatGPT_Excel_助手.ahk`
- Modify: `tests/一鍵測試.ahk`

**Interfaces:**
- Consumes: Task 1 全部核心函式
- Consumes: Task 2 `MouseStatus` 公開方法
- Produces: `InitializeWorkbook() -> Boolean`
- Produces: `OpenIncognitoChatGPT()`
- Produces: `SendCurrentQuestion()`
- Produces: `CopyAndSaveAnswer()`
- Produces: `CloseFocusedChromeWindow()`
- Produces: `ToggleEnabled()`、`DismissWarning()`、`ExitAssistant()`

- [ ] **Step 1: 加入可替换边界的测试案例**

在測試檔以假物件驗證流程決定：取消選檔、取消選格、多格選取、空白問題、非 Chrome 前景、剪貼簿逾時、Excel 寫入失敗、Excel 儲存失敗都不前進；完整成功才把偏移量加一。

- [ ] **Step 2: 實作啟動與 Excel 選格**

使用 `FileSelect` 選活頁簿，以 Excel COM 開啟並顯示；使用 `Application.InputBox(Type := 8)` 依序取得單一問題 Range 與回答 Range。取消或多格時釋放 COM 參照並結束腳本，但不關閉使用者已開啟的 Excel。

- [ ] **Step 3: 實作 Chrome 開啟與前景檢查**

依規格順序尋找 `chrome.exe`，以 `--incognito --new-window https://chatgpt.com/` 啟動。成功後讀取目前問題並顯示 120 字、四行預覽；空白問題顯示紅色警告。

- [ ] **Step 4: 實作 X 送出流程**

只在前景程序為 `chrome.exe` 且問題非空時，將完整問題放入剪貼簿，依序傳送 `Shift+Esc`、`Ctrl+V`、`Enter`；不移動列號。

- [ ] **Step 5: 實作 C 複製與儲存流程**

清空剪貼簿後傳送 `Ctrl+Shift+C` 並限時等待；取得新文字後寫入回答 Range、儲存活頁簿。使用 `CanAdvance()` 保證只有複製、寫入、儲存全部成功時偏移量加一。

- [ ] **Step 6: 實作其餘快捷鍵**

註冊 `F1/F2/Z/X/C/V/A`。暫停時讓五個字母恢復正常輸入；`A` 只對目前聚焦的 Chrome 傳送 `Alt+F4`；`F2` 釋放資源但保留 Excel 視窗。

- [ ] **Step 7: 靜態整合檢查**

使用 `rg` 核对七个快捷键、所有公开介面、ChatGPT URL、`--incognito --new-window`、`Shift+Esc`、`Ctrl+Shift+C` 与完整提示文案均已出现一次且没有占位词。

### Task 4: 測試活頁簿與懶人驗收清單

**Files:**
- Create: `AHK功能測試.xlsx`
- Create: `測試清單.md`

**Interfaces:**
- Consumes: Task 3 全部使用者可見流程
- Produces: 可直接选择的测试活页簿与不需阅读程式码的测试步骤

- [ ] **Step 1: 建立測試活頁簿**

建立 `測試` 工作表：`A1=問題`、`B1=回答`，從 `A2` 放入三條無害且容易辨識的測試問題，`B2:B4` 留白。凍結首列並設定適合長文字的欄寬與自動換行。

- [ ] **Step 2: 建立一頁式測試清單**

清單只要求使用者依序執行：雙擊主腳本、選擇測試活頁簿、點 `A2/B2`、按 `Z/X/C`、檢查 `B2` 與下一題提示、製造一次非 Chrome 警告、按 `V`、按 `A`、按 `F2`。每一步附唯一預期提示。

- [ ] **Step 3: 驗證檔案結構**

檢查活頁簿可由標準 Office Open XML 讀取，且工作表名稱、六個目標儲存格與三條測試問題正確；不啟動 Excel GUI。

### Task 5: 最終一致性檢查與交付

**Files:**
- Modify: `測試清單.md`（仅在检查发现差异时）

**Interfaces:**
- Consumes: Tasks 1–4 的全部交付物
- Produces: 可由使用者自行运行的完整资料夹

- [ ] **Step 1: 規格對照**

逐項對照設計規格驗收條件，記錄每項對應的檔案與函式；發現差異時先修正再繼續。

- [ ] **Step 2: UTF-8 與路徑檢查**

以 PowerShell UTF-8 安全读取文字档，确认中文路径、中文提示与 AHK 源码未乱码。

- [ ] **Step 3: 交付未執行聲明**

明確告訴使用者：已完成靜態檢查與 `.xlsx` 結構檢查，但遵循其要求，沒有替他啟動 AHK、Chrome 或 Excel；實際 GUI 與快捷鍵由使用者依 `測試清單.md` 驗收。
