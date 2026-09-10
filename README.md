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

`renderSpotSheets()` 在替換 DOM 前關閉既有 popup，完成後初始化新圖庫。
事件附加在每次建立的新元素上；不要對同一組圖庫重複呼叫 `enableSpotGalleries()`。
`closeSpotImageDialog()` 可重複呼叫，負責關閉預覽並還原頁面捲動與焦點。

部署時須將所有 HTML 檔加入 Apps Script，包括新增的四個模組檔，並更新部署版本。
