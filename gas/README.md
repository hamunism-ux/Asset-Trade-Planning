# gas/ — Apps Script 端的程式

- `Code.gs`：**載入器**（唯一要貼進試算表 Apps Script 的檔案）。每次執行時從這裡的 `core.js` 取得最新程式。
- `core.js`：伺服器程式（由 build 從原始 Code.gs 產生，附匯出清單）。
- `App.html`：網頁介面（Apps Script 版；網頁會直接讀這裡的最新版）。
- `appsscript.json`：參考用，載入器方式不需要額外權限，用預設資訊清單即可。

資料與金鑰都不在這裡。
