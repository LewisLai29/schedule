const SPREADSHEET_ID = '1VNJX4mWQhhdLI-dv_yp7yoomvDIGOkFdt1Irnkr4Iy0';

const SOURCE_RANGES = {
  schedule: { sheetName: 'schedule', range: 'A1:K100', rowCount: 100, columnCount: 11, expandMergedCells: true, readNotes: true },
  food: { sheetName: 'food', range: 'A1:K100', rowCount: 100, columnCount: 11 },
};

function doGet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tripSettings = readTripSettings_(spreadsheet);
  const template = HtmlService.createTemplateFromFile('Index');
  template.tripSettings = tripSettings;

  return template
    .evaluate()
    .setTitle(tripSettings.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Optional settings tab: column A contains keys; column B contains display text.
function readTripSettings_(spreadsheet) {
  const settings = {
    title: spreadsheet.getName(),
    subtitle: '',
    flights: '',
  };
  const sheet = spreadsheet.getSheetByName('settings');
  if (!sheet || !sheet.getLastRow()) return settings;

  sheet.getRange(1, 1, sheet.getLastRow(), 2).getDisplayValues().forEach(([key, value]) => {
    key = key.trim().toLowerCase();
    value = value.trim();
    if (Object.prototype.hasOwnProperty.call(settings, key) && value) settings[key] = value;
  });
  return settings;
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getTripData() {
  // Preserve the original endpoint for older HTML and editor diagnostics.
  const scheduleData = getScheduleData();
  const details = getTripDetails();
  return {
    ...details,
    ...scheduleData,
    kobeSpot: (details.spotSheets.find((sheet) => sheet.name.toLowerCase() === 'kobe spot') || {}).rows || [],
    spot: (details.spotSheets.find((sheet) => sheet.name.toLowerCase() === 'spot') || {}).rows || [],
    spotReferences: collectSpotReferences_(details.spotSheets, details.hotel || []),
  };
}

// Keep the first visible itinerary independent of images and Maps services.
function getScheduleData() {
  const startedAt = Date.now();
  let schedule;
  try {
    schedule = readGridRanges_([SOURCE_RANGES.schedule])[0];
    console.info('schedule: Sheets API batch', Date.now() - startedAt);
  } catch (error) {
    console.warn('行程批次讀取未啟用，改用相容讀取：', error.message || String(error));
    schedule = readRichRange_(SpreadsheetApp.openById(SPREADSHEET_ID), SOURCE_RANGES.schedule);
    console.info('schedule: compatibility reader', Date.now() - startedAt);
  }
  return {
    schedule,
    ...readTrafficData_(),
    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
  };
}

// Optional sheet: a failed traffic read must not hide the itinerary or Maps.
function readTrafficData_() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheets().find((item) => item.getName().trim().toLowerCase() === 'traffic');
    if (!sheet || !sheet.getLastRow()) return { traffic: [] };
    const range = sheet.getDataRange();
    const config = { sheetName: sheet.getName(), range: range.getA1Notation(),
      rowCount: range.getNumRows(), columnCount: range.getNumColumns(), imageStartColumn: 1, readLinks: true };
    let rows;
    try {
      rows = readGridRanges_([config])[0];
      if (config.columnCount > 1) applyColumnImages_(rows, sheet, 1, 1, config.columnCount - 1);
    } catch (error) {
      rows = readRichRange_(spreadsheet, config, sheet);
    }
    return { traffic: rows };
  } catch (error) {
    console.warn('交通筆記讀取失敗：', error.message || String(error));
    return { traffic: [], trafficError: '交通筆記暫時無法載入，請重新整理再試。' };
  }
}

function getTripDetails() {
  return readTripDetails_();
}

const MAP_URL_PATTERN = /^https:\/\/(?:maps\.app\.goo\.gl\/|goo\.gl\/maps(?:\/|\?)|(?:www\.|maps\.)?google\.(?:com|co\.jp|com\.tw)\/(?:maps(?:\/|\?)|\?))/i;

// Read formatted text, all link types, formulas and merge bounds together.
// Include notes only for the itinerary, linked text only when requested.
function readGridRanges_(configs) {
  if (typeof Sheets === 'undefined') throw new Error('請啟用 Google Sheets API（v4）。');
  const response = Sheets.Spreadsheets.get(SPREADSHEET_ID, {
    ranges: configs.map((config) => "'" + config.sheetName.replace(/'/g, "''") + "'!" + config.range),
    fields: 'sheets(properties(title),merges(startRowIndex,endRowIndex,startColumnIndex,endColumnIndex),data(startRow,startColumn,rowData(values(formattedValue,hyperlink,userEnteredValue(formulaValue),textFormatRuns(startIndex,format(link(uri))),chipRuns(chip(richLinkProperties(uri)))'
      + (configs.some((config) => config.readNotes) ? ',note' : '') + '))))',
  });
  return configs.map((config) => {
    const sheet = (response.sheets || []).find((item) => item.properties.title === config.sheetName);
    if (!sheet) throw new Error('批次資料缺少工作表：' + config.sheetName);
    return gridRows_(sheet, config);
  });
}

// All configured ranges start at A1. Empty API cells/rows are omitted, so
// restore their positions before expanding merges or overlaying image cells.
function gridRows_(sheet, config) {
  const rows = Array.from({ length: config.rowCount }, () =>
    Array.from({ length: config.columnCount }, () => ({ text: '', url: '' })));
  (sheet.data || []).forEach((grid) => {
    (grid.rowData || []).forEach((row, rowOffset) => {
      const rowIndex = (grid.startRow || 0) + rowOffset;
      if (rowIndex >= rows.length) return;
      (row.values || []).forEach((value, columnOffset) => {
        const columnIndex = (grid.startColumn || 0) + columnOffset;
        if (columnIndex >= config.columnCount) return;
        const formula = value.userEnteredValue?.formulaValue || '';
        const links = [...new Set((value.textFormatRuns || [])
          .map((run) => run.format?.link?.uri).filter(Boolean))];
        const cell = {
          text: value.formattedValue || '',
          url: value.hyperlink || (links.length === 1 ? links[0] : '') || richCellUrl_(null, formula),
        };
        if (config.readNotes && value.note) cell.note = value.note;
        if (config.readLinks && value.textFormatRuns?.length) {
          cell.runs = value.textFormatRuns.map((run, index, runs) => ({
            text: (value.formattedValue || '').slice(run.startIndex || 0, runs[index + 1]?.startIndex),
            url: run.format?.link?.uri || '',
          }));
        }
        if (columnIndex < (config.mapColumnCount || 2)) {
          const mapUrl = mapChipUrl_(value);
          if (mapUrl) cell.mapUrl = mapUrl;
        }
        if ((config.readFirstColumnImages && columnIndex === 0) || columnIndex === config.imageColumn
          || (Number.isInteger(config.imageStartColumn) && columnIndex >= config.imageStartColumn)) {
          const imageUrl = imageFormulaUrl_(formula);
          if (imageUrl) {
            rows[rowIndex][columnIndex] = { text: '', url: '', imageUrl, imageAlt: '' };
            return;
          }
        }
        rows[rowIndex][columnIndex] = cell;
      });
    });
  });
  if (config.expandMergedCells) {
    (sheet.merges || []).forEach((merge) => {
      const firstRow = merge.startRowIndex || 0;
      const firstColumn = merge.startColumnIndex || 0;
      const lastRow = Math.min(merge.endRowIndex, rows.length) - 1;
      const lastColumn = Math.min(merge.endColumnIndex, config.columnCount) - 1;
      const anchor = rows[firstRow]?.[firstColumn];
      if (!anchor) return;
      for (let row = firstRow; row <= lastRow; row += 1) {
        for (let column = firstColumn; column <= lastColumn; column += 1) {
          rows[row][column] = { ...anchor, mergeStartRow: firstRow, mergeEndRow: lastRow };
        }
      }
    });
  }
  return rows;
}

// The explicit modes are for editor diagnostics; normal loading uses auto.
function readTripDetails_(mode = 'auto') {
  const startedAt = Date.now();
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = spreadsheet.getSheets();
  const spots = sheets.filter((sheet) => /spot$/i.test(sheet.getName().trim()));
  const hotelSheet = sheets.find((sheet) => sheet.getName().trim().toLowerCase() === 'hotel');
  const settingsSheet = sheets.find((sheet) => sheet.getName() === 'settings');
  const configFor = (sheet, images) => {
    const range = sheet.getDataRange();
    return {
      sheetName: sheet.getName(), range: range.getA1Notation(),
      rowCount: range.getNumRows(), columnCount: range.getNumColumns(),
      readFirstColumnImages: images,
    };
  };
  const configs = [
    SOURCE_RANGES.food,
    ...spots.map((sheet) => configFor(sheet, true)),
    ...(hotelSheet ? [configFor(hotelSheet, false)] : []),
    ...(settingsSheet ? [{
      sheetName: settingsSheet.getName(), range: 'A1:C' + Math.max(1, settingsSheet.getLastRow()),
      rowCount: Math.max(1, settingsSheet.getLastRow()), columnCount: 3, mapColumnCount: 3,
    }] : []),
  ];
  let result;
  if (mode !== 'compatible') {
    try { result = readGridRanges_(configs); }
    catch (error) {
      if (mode === 'batch') throw error;
      console.warn('旅遊資料批次讀取未啟用，改用相容讀取：', error.message || String(error));
    }
  }
  let mapLocationError = '';
  const usedBatch = Boolean(result);
  if (usedBatch) {
    spots.forEach((sheet, index) => applyFirstColumnImages_(result[index + 1], sheet));
  } else {
    // Reuse discovery results; only the data reads need a fallback.
    result = configs.map((config) => readRichRange_(spreadsheet, config,
      sheets.find((sheet) => sheet.getName() === config.sheetName)));
    try {
      readMapChipUrls_(configs.map((config, index) => ({
        name: config.sheetName, rows: result[index], columnCount: config.mapColumnCount || 2,
      })));
    } catch (error) {
      mapLocationError = '地圖智慧方塊讀取失敗：' + (error.message || String(error));
      console.warn(mapLocationError);
    }
  }
  const food = trimEmptyEdges_(result[0]);
  const spotSheets = spots.map((sheet, index) => ({
    name: sheet.getName().trim(), rows: trimEmptyEdges_(result[index + 1]),
  }));
  const settingsHotels = settingsSheet ? settingsHotelRows_(result[result.length - 1]) : null;
  const hotelIndex = 1 + spots.length;
  const hotel = trimEmptyEdges_(settingsHotels === null
    ? (hotelSheet ? result[hotelIndex] : []) : settingsHotels);
  const mapLocations = collectMapLocations_(spotSheets, hotel, food);
  console.info('details: ' + (usedBatch ? 'Sheets API batch + image columns' : 'compatibility reader'), Date.now() - startedAt);
  return { food, spotSheets, hotel, mapLocationError, mapLocations };
}

// A hotels key starts a block; the next non-empty key ends it.
// null means no override, while an empty block explicitly clears old hotels.
function settingsHotelRows_(rows) {
  let inHotels = false;
  let found = false;
  const hotels = [];
  rows.forEach((row) => {
    const key = String(row[0]?.text || '').trim().toLowerCase();
    if (key) {
      inHotels = key === 'hotels';
      if (inHotels) found = true;
    }
    if (!inHotels) return;
    const cells = [row[1] || { text: '', url: '' }, row[2] || { text: '', url: '' }];
    if (cells.some((cell) => cell.text || cell.url || cell.mapUrl)) hotels.push(cells);
  });
  return found ? hotels : null;
}

function imageFormulaUrl_(formula) {
  const match = String(formula || '').match(/^=\s*IMAGE\(\s*"((?:[^"]|"")*)"\s*(?:[,;)]|$)/i);
  const url = match ? match[1].replace(/""/g, '"') : '';
  return /^https?:\/\//i.test(url) ? url : '';
}

// Native CellImage URLs are unavailable in Sheets API. Share the overlay
// between both readers, keeping IMAGE formula URLs when content URLs are absent.
function applyFirstColumnImages_(rows, sheet, firstRow = 1) {
  applyColumnImages_(rows, sheet, 0, firstRow);
}

function applyColumnImages_(rows, sheet, column, firstRow = 1, columnCount = 1) {
  if (!rows.length) return;
  const images = sheet.getRange(firstRow, column + 1, rows.length, columnCount).getValues();
  images.forEach((row, index) => {
    row.forEach((value, offset) => {
      if (!value || typeof value.getContentUrl !== 'function') return;
      const cell = rows[index][column + offset];
      cell.imageRef = { sheet: sheet.getName(), row: firstRow + index, column: column + offset + 1 };
      try {
        const imageUrl = value.getContentUrl() || cell.imageUrl || '';
        if (/^https?:\/\//i.test(imageUrl)) {
          cell.text = '';
          cell.url = '';
          cell.imageUrl = imageUrl;
          cell.imageAlt = value.getAltTextDescription() || value.getAltTextTitle() || '';
        }
      } catch (error) {
        // Keep the reference so the browser can retry without losing the whole sheet.
        console.warn('儲存格圖片讀取失敗：', sheet.getName(), firstRow + index, error.message || String(error));
      }
    });
  });
}

// Re-read native images only after a browser load failure. Never accepts a URL
// from the client: the image must belong to an allowed cell in this workbook.
function getTripImage(sheetName, rowNumber, columnNumber) {
  if (typeof sheetName !== 'string' || !Number.isInteger(rowNumber) || rowNumber < 1) throw new Error('無效的圖片位置。');
  const isTraffic = sheetName.trim().toLowerCase() === 'traffic';
  if (!isTraffic && !/spot$/i.test(sheetName.trim())) throw new Error('不支援的圖片工作表。');
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet || rowNumber > sheet.getLastRow()) throw new Error('找不到圖片儲存格。');
  const column = columnNumber === undefined ? (isTraffic ? 2 : 1) : columnNumber;
  if (!Number.isInteger(column) || (isTraffic ? column < 2 || column > sheet.getLastColumn() : column !== 1)) throw new Error('無效的圖片欄位。');
  const value = sheet.getRange(rowNumber, column).getValue();
  if (!value || typeof value.getContentUrl !== 'function') throw new Error('此儲存格沒有儲存格內圖片。');
  const url = value.getContentUrl();
  // The content URL comes from CellImage, not from editable cell text.
  if (!/^https:\/\//i.test(url)) throw new Error('無法取得圖片網址。');
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error('原圖讀取失敗，請確認執行帳號可存取試算表。');
  const blob = response.getBlob();
  const mime = String(blob.getContentType() || '').split(';')[0].trim().toLowerCase();
  if (!/^image\/(png|jpeg|gif|webp|bmp|avif)$/.test(mime)) throw new Error('圖片來源未回傳可顯示的圖片。');
  const bytes = blob.getBytes();
  if (bytes.length > 8 * 1024 * 1024) throw new Error('圖片超過 8 MB，請縮小圖片後再試。');
  return { dataUrl: 'data:' + mime + ';base64,' + Utilities.base64Encode(bytes) };
}

function mapChipUrl_(value) {
  const urls = [...new Set((value.chipRuns || [])
    .map((run) => run.chip?.richLinkProperties?.uri || '')
    .filter((url) => MAP_URL_PATTERN.test(url)))];
  // Multiple places in a cell do not identify a single route endpoint.
  return urls.length === 1 ? urls[0] : '';
}

function trimEmptyEdges_(rows) {
  const hasData = (cell) => cell.text || cell.url || cell.mapUrl || cell.imageUrl || cell.imageRef;
  rows.forEach((row) => {
    while (row.length && !hasData(row[row.length - 1])) row.pop();
  });
  while (rows.length && !rows[rows.length - 1].length) rows.pop();
  return rows;
}

// Place smart chips store their URLs in chipRuns, not RichTextValue hyperlinks.
// Read location columns, including settings column C, in one API request.
function readMapChipUrls_(mapSheets) {
  const sheets = mapSheets.filter((sheet) => sheet.rows.length);
  if (!sheets.length) return;
  if (typeof Sheets === 'undefined') {
    throw new Error('請在 Apps Script「服務」加入 Google Sheets API（v4）後重新部署。');
  }
  const byName = new Map(sheets.map((sheet) => [sheet.name.trim(), sheet.rows]));
  const response = Sheets.Spreadsheets.get(SPREADSHEET_ID, {
    ranges: sheets.map((sheet) => "'" + sheet.name.replace(/'/g, "''") + "'!A1:"
      + (sheet.columnCount === 3 ? 'C' : 'B') + sheet.rows.length),
    fields: 'sheets(properties(title),data(startRow,startColumn,rowData(values(chipRuns(chip(richLinkProperties(uri)))))))',
  });
  (response.sheets || []).forEach((sheet) => {
    const rows = byName.get(sheet.properties.title.trim());
    if (!rows) return;
    (sheet.data || []).forEach((grid) => {
      (grid.rowData || []).forEach((row, rowOffset) => {
        (row.values || []).forEach((value, columnOffset) => {
          const cell = rows[(grid.startRow || 0) + rowOffset]?.[(grid.startColumn || 0) + columnOffset];
          if (!cell) return;
          const url = mapChipUrl_(value);
          if (url) cell.mapUrl = url;
        });
      });
    });
  });
}

// Initial page loading must never wait for Google Maps network requests.
function collectMapLocations_(spotSheets, hotelRows, foodRows = []) {
  const locations = [];
  spotSheets.map((sheet) => sheet.rows).concat([hotelRows, foodRows]).forEach((rows) => {
    rows.forEach((row) => {
      const name = (row[0]?.text || row[1]?.text || '').trim();
      const secondary = row[1] || {};
      const url = [secondary.mapUrl, secondary.url, secondary.text, row[0]?.mapUrl, row[0]?.url, row[0]?.text]
        .map((value) => String(value || '').trim()).find((value) => MAP_URL_PATTERN.test(value)) || '';
      if (!name || !url) return;
      locations.push({ name, alias: secondary.text || '', originalUrl: url, url });
    });
  });
  // Long full URLs exceed CacheService's key limit and need no expansion.
  const keys = [...new Set(locations.map((location) => 'map-url:v2:' + location.url))]
    .filter((key) => key.length <= 250);
  try {
    const cached = keys.length ? CacheService.getScriptCache().getAll(keys) : {};
    locations.forEach((location) => { location.url = cached['map-url:v2:' + location.url] || location.url; });
  } catch (error) {
    console.warn('地圖快取暫時無法讀取，使用原始連結：', error.message || String(error));
  }
  return locations;
}

// Resolve only the two endpoints requested by a route button.
function resolveMapLocations(urls) {
  if (!Array.isArray(urls) || urls.length > 2) throw new Error('路線最多包含兩個位置。');
  const cache = CacheService.getScriptCache();
  return [...new Set(urls)].map((url) => {
    if (typeof url !== 'string' || !MAP_URL_PATTERN.test(url)) throw new Error('地圖連結格式不正確。');
    const key = 'map-url:v2:' + url;
    let expanded = cache.get(key) || url;
    // Stop at the full Maps URL; fetching the Maps page itself is unnecessary.
    for (let hop = 0; hop < 5 && /^https:\/\/(?:maps\.app\.goo\.gl|goo\.gl)\//i.test(expanded); hop += 1) {
      const response = UrlFetchApp.fetch(expanded, { followRedirects: false, muteHttpExceptions: true });
      if (response.getResponseCode() < 300 || response.getResponseCode() >= 400) break;
      const headers = response.getAllHeaders();
      let next = headers.Location || headers.location;
      if (typeof next === 'string' && /^\/(?!\/)/.test(next)) {
        next = expanded.match(/^https:\/\/[^/]+/i)[0] + next;
      }
      if (typeof next !== 'string' || !MAP_URL_PATTERN.test(next)) break;
      expanded = next;
    }
    if (expanded !== url && !/^https:\/\/(?:maps\.app\.goo\.gl|goo\.gl)\//i.test(expanded)) {
      cache.put(key, expanded, 21600);
    }
    return { originalUrl: url, url: expanded };
  });
}

/**
 * Name references retained only for clients using the legacy getTripData API.
 */
function collectSpotReferences_(spotSheets, hotelRows) {
  const references = [];
  const seen = new Set();

  const addReference = (name) => {
    const normalized = name.toLowerCase().replace(/\s+/g, ' ').trim();
    if (!normalized || seen.has(normalized)) return;
    seen.add(normalized);
    references.push(name);
  };

  spotSheets.forEach(({ rows }) => {
    rows.forEach((row) => {
      const primary = row[0] || {};
      const secondary = row[1] || {};
      const primaryText = (primary.text || '').trim();
      const secondaryText = (secondary.text || '').trim();
      const isItem = primaryText && (secondaryText || primary.url || secondary.url);

      if (!isItem) return;

      [primaryText, secondaryText].forEach((name) => {
        addReference(name);
      });
    });
  });

  hotelRows.forEach((row) => {
    row.slice(0, 2).forEach((cell) => {
      const name = (cell.text || '').trim();
      if (name) addReference(name);
    });
  });

  return references;
}

/**
 * Read display text and cell-level hyperlinks.
 * Display values preserve the formatting used for dates, times and prices.
 */
function readRichRange_(spreadsheet, config, sheet = spreadsheet.getSheetByName(config.sheetName)) {
  if (!sheet) {
    throw new Error('找不到工作表：' + config.sheetName);
  }

  const range = sheet.getRange(config.range);
  const values = range.getDisplayValues();
  const richTextValues = range.getRichTextValues();
  const formulas = range.getFormulas();
  const notes = config.readNotes ? range.getNotes() : [];
  const rangeRow = range.getRow();
  const rangeColumn = range.getColumn();
  const mergedCells = new Map();

  if (config.expandMergedCells) {
    range.getMergedRanges().forEach((mergedRange) => {
      const sourceRow = mergedRange.getRow() - rangeRow;
      const sourceColumn = mergedRange.getColumn() - rangeColumn;
      let mergedText;
      let mergedUrl;
      if (sourceRow >= 0 && sourceColumn >= 0) {
        // Reuse the bulk read instead of two service reads per merged cell.
        mergedText = values[sourceRow][sourceColumn] || '';
        mergedUrl = richCellUrl_(richTextValues[sourceRow][sourceColumn], formulas[sourceRow][sourceColumn]);
      } else {
        // A range can intersect a merge whose anchor lies outside it.
        const topLeft = mergedRange.getCell(1, 1);
        mergedText = topLeft.getDisplayValue() || '';
        mergedUrl = richCellUrl_(topLeft.getRichTextValue(), topLeft.getFormula());
      }
      const firstRow = Math.max(0, sourceRow);
      const lastRow = Math.min(
        values.length - 1,
        mergedRange.getLastRow() - rangeRow
      );
      const firstColumn = Math.max(0, sourceColumn);
      const lastColumn = Math.min(
        values[0].length - 1,
        mergedRange.getLastColumn() - rangeColumn
      );

      for (let rowIndex = firstRow; rowIndex <= lastRow; rowIndex += 1) {
        for (let columnIndex = firstColumn; columnIndex <= lastColumn; columnIndex += 1) {
          mergedCells.set(`${rowIndex}:${columnIndex}`, {
            text: mergedText,
            url: mergedUrl,
            mergeStartRow: firstRow,
            mergeEndRow: lastRow,
            ...(config.readNotes ? { note: sourceRow >= 0 && sourceColumn >= 0
              ? notes[sourceRow]?.[sourceColumn] || '' : mergedRange.getCell(1, 1).getNote() } : {}),
          });
        }
      }
    });
  }

  const rows = values.map((row, rowIndex) => row.map((text, columnIndex) => {
    if ((config.readFirstColumnImages && rangeColumn + columnIndex === 1) || columnIndex === config.imageColumn
      || (Number.isInteger(config.imageStartColumn) && columnIndex >= config.imageStartColumn)) {
      const imageUrl = imageFormulaUrl_(formulas[rowIndex][columnIndex]);
      if (imageUrl) return { text: '', url: '', imageUrl, imageAlt: '' };
    }
    const mergedCell = mergedCells.get(`${rowIndex}:${columnIndex}`);
    if (mergedCell) return mergedCell;

    const richText = richTextValues[rowIndex][columnIndex];
    return {
      text: text || '',
      url: richCellUrl_(richText, formulas[rowIndex][columnIndex]),
      ...(config.readNotes ? { note: notes[rowIndex]?.[columnIndex] || '' } : {}),
      ...(config.readLinks && richText ? { runs: richText.getRuns().map((run) => ({ text: run.getText(), url: run.getLinkUrl() || '' })) } : {}),
    };
  }));
  if (config.readFirstColumnImages && rangeColumn === 1) applyFirstColumnImages_(rows, sheet, rangeRow);
  if (Number.isInteger(config.imageColumn)) applyColumnImages_(rows, sheet, config.imageColumn, rangeRow);
  if (Number.isInteger(config.imageStartColumn) && config.columnCount > config.imageStartColumn) {
    applyColumnImages_(rows, sheet, config.imageStartColumn, rangeRow, config.columnCount - config.imageStartColumn);
  }
  return rows;
}

function richCellUrl_(richText, formula) {
  const url = richText ? richText.getLinkUrl() : '';
  if (url) return url;
  const links = richText ? [...new Set(richText.getRuns().map((run) => run.getLinkUrl()).filter(Boolean))] : [];
  if (links.length === 1) return links[0];
  const hyperlink = String(formula || '').match(/^=\s*HYPERLINK\(\s*"((?:[^"]|"")*)"\s*[,;]/i);
  return hyperlink ? hyperlink[1].replace(/""/g, '"') : '';
}

/**
 * Run this once in the Apps Script editor to check authorization and reading.
 */
function testRead() {
  const data = getTripData();

  Logger.log(JSON.stringify({
    scheduleRows: data.schedule.length,
    foodRows: data.food.length,
    spotSheets: data.spotSheets.map((sheet) => ({ name: sheet.name, rows: sheet.rows.length })),
    mapLocations: data.mapLocations,
    updatedAt: data.updatedAt,
  }));
}

/**
 * Run manually in the editor to compare actual backend reads, without the
 * automatic compatibility fallback hiding a broken Sheets API setup.
 * Re-run for representative results; service warm-up affects each sample.
 */
function testReadPerformance() {
  const results = [];
  const readers = [
    ['schedule-compatible', () => readRichRange_(SpreadsheetApp.openById(SPREADSHEET_ID), SOURCE_RANGES.schedule)],
    ['schedule-batch', () => readGridRanges_([SOURCE_RANGES.schedule])[0]],
    ['details-compatible', () => readTripDetails_('compatible')],
    ['details-batch', () => readTripDetails_('batch')],
  ];
  readers.forEach(([name, read]) => {
    const startedAt = Date.now();
    try {
      const data = read();
      results.push({ name, durationMs: Date.now() - startedAt,
        payloadBytes: Utilities.newBlob(JSON.stringify(data)).getBytes().length });
    } catch (error) {
      results.push({ name, durationMs: Date.now() - startedAt, error: error.message || String(error) });
    }
  });
  Logger.log(JSON.stringify(results));
  return results;
}
