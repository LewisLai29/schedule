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
`renderTrip()` 完成各區塊渲染後呼叫 `initializeTripMaps()`；不透過 DOM 變動監聽器重跑地圖初始化。

交通路線使用景點、餐飲與住宿資料的 Google Maps 位置；優先讀取第二欄地圖連結，也支援名稱欄的地圖連結。地點智慧方塊的網址透過 Google Sheets API 的 `chipRuns.chip.richLinkProperties.uri` 讀取，以 `cell.mapUrl` 加入既有位置索引，保留原本的顯示文字與一般超連結。

初次載入以一次 Sheets API 請求批次讀取景點、餐飲與住宿分頁的 A、B 欄智慧方塊資料，不連線解析地圖短網址。點擊路線時才由 `resolveMapLocations()` 展開該段起終點的分享短網址，前端取出座標、Place ID 或連結內的位置查詢值。缺少可用位置時顯示「無法取得位置」；智慧方塊 API 讀取失敗則顯示「地圖資料未載入」，不影響行程顯示。短網址解析失敗可點擊重試，不再用行程名稱加上「日本」搜尋。

部署前，請在 Apps Script 編輯器左側「服務」按「＋」，加入 **Google Sheets API（v4，識別碼 Sheets）**。使用 Apps Script 預設 Cloud 專案時，API 會一起啟用；若使用自訂標準 Cloud 專案，也要在該專案啟用 Google Sheets API。儲存程式、完成授權並更新部署版本後，智慧方塊內的地圖位置才會讀入。

`renderSpotSheets()` 在替換 DOM 前關閉既有 popup，完成後初始化新圖庫。
事件附加在每次建立的新元素上；不要對同一組圖庫重複呼叫 `enableSpotGalleries()`。
`closeSpotImageDialog()` 可重複呼叫，負責關閉預覽並還原頁面捲動與焦點。

部署時須將所有 HTML 檔加入 Apps Script，包括新增的四個模組檔，並更新部署版本。
