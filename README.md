# schedule

Google Apps Script HTML Service 行程網站。

## 檔案職責

| 檔案 | 職責 |
| --- | --- |
| `Code.gs` | 讀取試算表、A 欄圖片與提供 HTML 範本 |
| `Index.html` | 頁面骨架與模組載入順序 |
| `TripRenderer.html` | 共用文字／連結工具、住宿／餐飲渲染、整體渲染流程 |
| `ScheduleRenderer.html` | 日期與行程解析、每日卡片、展開動畫 |
| `SpotRenderer.html` | 景點解析、卡片與資訊區塊 |
| `SpotGallery.html` | 圖庫 HTML、左右切換、popup 生命週期與縮放手勢 |
| `Maps.html` | 地圖按鈕、交通路線與地圖展開動畫 |
| `App.html` | 資料請求、錯誤處理與啟動 |
| `Styles.html` | 共用版面與行程／景點樣式 |
| `SpotGalleryStyles.html` | 圖庫與 popup 樣式 |

## 載入與更新

`Index.html` 先載入共用工具、行程、圖庫、景點與地圖，最後才由 `App.html` 請求資料。
前端同時發出兩個獨立請求：`getScheduleData()` 只讀取每日行程，完成後立即顯示；`getTripDetails()` 在背景讀取景點圖片、餐飲、住宿與地圖位置。後者完成時不會重建行程卡片，保留使用者已展開的日期。兩者完成後才初始化地圖；其中一組資料讀取失敗，只顯示該區的錯誤，不清除另一組已載入的內容。

後端優先使用 Sheets API 的多範圍讀取與欄位篩選：行程只需一次 API 請求，景點、餐飲與住宿合計另一次，直接取得顯示文字、連結、公式、合併範圍與地圖智慧方塊，不再逐張表分次讀取文字／RichText／公式，也不再逐個合併儲存格讀取內容。直接插入儲存格的圖片仍需各景點表 A 欄的一次 `getValues()` 取得 `CellImage` 網址；工作表清單與使用範圍仍由 SpreadsheetApp 讀取。因此「兩次」指資料的 Sheets API 請求，並非所有 Google 服務呼叫的總數。地圖網址快取改用一次 `getAll()`，各區清單也移除尾端空白欄列以減少傳輸量，保留內部空白列的位置。

每次開啟仍直接讀取試算表，不使用行程／景點資料快取。兩個請求分開讀取，若載入途中編輯試算表，各區可能反映不同時間點的資料。舊端點 `getTripData()` 保留既有回傳格式（包含舊別名與 `spotReferences`）；新版端點不再計算或傳送這些未使用的資料。前端只保留目前使用的分段渲染函式。

Sheets API 未啟用或讀取失敗時會記錄原因並退回相容讀取，維持資料可用，但此時不會有批次讀取的效能改善。兩條路徑共用工作表探索、圖片處理與回傳資料組裝，備援時不再重新讀取工作表清單。圖片公式、智慧方塊與圖片說明使用共用解析，避免兩條路徑產生不同結果。

瀏覽器主控台會記錄「行程顯示耗時」及「其他旅遊資料顯示耗時」（從前端發出請求起算，不含 HTML 頁面本身的載入）。

Apps Script 執行記錄會標示 `Sheets API batch` 或 `compatibility reader` 與後端耗時。

景點與餐飲清單的「顯示地圖」固定在項目下方向下展開小地圖，再次點擊「隱藏地圖」收合；保留 260ms 動畫與首次展開才載入 iframe 的行為。小地圖沿用原本的地點名稱查詢，只有面板內的「在 Google Maps 開啟」連結會另開分頁。此流程與每日行程的交通路線按鈕分開初始化，不因資料中有地圖連結就改成直接跳轉。

交通路線使用景點、餐飲與住宿資料的 Google Maps 位置；優先讀取第二欄地圖連結，也支援名稱欄的地圖連結。地點智慧方塊的網址透過 Google Sheets API 的 `chipRuns.chip.richLinkProperties.uri` 讀取，以 `cell.mapUrl` 加入既有位置索引，保留原本的顯示文字與一般超連結。

初次載入的地圖智慧方塊隨其他資料一併批次取得，只取 A、B 欄作為位置來源，不連線解析地圖短網址。相容讀取路徑才另發一次 Sheets API 請求讀取 A、B 欄智慧方塊。點擊路線時才由 `resolveMapLocations()` 展開該段起終點的分享短網址，前端取出座標、Place ID 或連結內的位置查詢值。缺少可用位置時顯示「無法取得位置」；相容路徑的智慧方塊 API 也讀取失敗時顯示「地圖資料未載入」，不影響行程顯示。短網址解析失敗可點擊重試，不再用行程名稱加上「日本」搜尋。

部署前，請在 Apps Script 編輯器左側「服務」按「＋」，加入 **Google Sheets API（v4，識別碼 Sheets）**。使用 Apps Script 預設 Cloud 專案時，API 會一起啟用；若使用自訂標準 Cloud 專案，也要在該專案啟用 Google Sheets API。儲存程式、完成授權並更新部署版本後，智慧方塊內的地圖位置才會讀入。

`renderSpotSheets()` 在替換 DOM 前關閉既有 popup，完成後初始化新圖庫。
事件附加在每次建立的新元素上；不要對同一組圖庫重複呼叫 `enableSpotGalleries()`。
`closeSpotImageDialog()` 可重複呼叫，負責關閉預覽並還原頁面捲動與焦點。

部署時須將所有 HTML 檔加入 Apps Script，包括新增的四個模組檔，並更新部署版本。

此次效能修改及清理需同步更新 `Code.gs`、`App.html`、`TripRenderer.html` 與 `Styles.html`，再更新部署版本。

API 參考：[多範圍讀取與欄位篩選](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/get)、[儲存格回應格式](https://developers.google.com/workspace/sheets/api/reference/rest/v4/spreadsheets/cells)、[CellImage](https://developers.google.com/apps-script/reference/spreadsheet/cell-image)。
