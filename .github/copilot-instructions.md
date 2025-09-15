# Copilot Instructions for uploadExamPaper

## 專案架構與主要元件

-   本專案為 Google Apps Script (GAS) 專案，結合 HTML 前端（`index.html`）與 Apps Script 後端（`程式碼.js`）。
-   使用者登入與 email 顯示流程：
    -   前端 `index.html` 透過 `google.script.run.getUserEmail()` 呼叫 GAS 端函式取得登入者 email。
    -   `程式碼.js` 需提供 `getUserEmail` 函式，回傳目前登入使用者的 email。
-   `appsscript.json` 設定專案時區、例外處理與執行環境（V8）。

## 關鍵檔案

-   `index.html`：首頁 UI，負責登入流程與 email 顯示，僅透過 Apps Script API 與後端互動。
-   `程式碼.js`：GAS 端邏輯，需暴露 `getUserEmail` 供前端呼叫。
-   `appsscript.json`：GAS 專案設定檔。

## 開發與部署流程

-   請使用 [clasp](https://github.com/google/clasp) 進行本地開發與部署。
    -   推薦指令：
        -   `clasp push`：將本地程式碼上傳至 Apps Script 專案。
        -   `clasp pull`：同步雲端程式碼至本地。
        -   `clasp open`：於瀏覽器開啟 Apps Script 編輯器。
-   前端 HTML 需嵌入於 Apps Script Web App，並透過 `google.script.run` 呼叫後端函式。

## 專案慣例

-   所有 GAS 端 API 需以 function 形式暴露於全域（勿嵌套於其他 function 內）。
-   前端與後端僅透過 `google.script.run` 進行通訊。
-   登入驗證與 email 取得僅依賴 `Session.getActiveUser().getEmail()`。
-   主要程式碼皆以繁體中文註解。
-   嚴格控制變數的 scope，避免全域變數污染。

## 重要注意事項

-   若需新增 GAS API，請確保其為全域 function，否則前端無法呼叫。
-   若需擴充登入資訊，請同步調整前端與後端資料結構。
-   本專案無自動化測試，請以本地與 Apps Script 編輯器手動驗證。

---

如需範例：

-   前端呼叫：`google.script.run.withSuccessHandler(cb).getUserEmail();`
-   後端暴露：
    ```javascript
    function getUserEmail() {
        return Session.getActiveUser().getEmail();
    }
    ```
