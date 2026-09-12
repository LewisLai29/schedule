const SPREADSHEET_ID = '1sJZjqZ60cUJrMvfbPUM6EwoG4wZf5S-BTWs_5or5yEw';

const SOURCE_RANGES = {
  schedule: { sheetName: 'schedule', range: 'A1:K100', rowCount: 100, columnCount: 11, expandMergedCells: true },
  food: { sheetName: 'food', range: 'A1:K100', rowCount: 100, columnCount: 11 },
};

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('岡山保佑一定好天氣')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
  };
}

const MAP_URL_PATTERN = /^https:\/\/(?:maps\.app\.goo\.gl\/|goo\.gl\/maps(?:\/|\?)|(?:www\.|maps\.)?google\.(?:com|co\.jp|com\.tw)\/(?:maps(?:\/|\?)|\?))/i;

// Read formatted text, all link types, formulas and merge bounds together.
// Keep the field mask narrow: no styling, notes or unrelated sheets.
function readGridRanges_(configs) {
  if (typeof Sheets === 'undefined') throw new Error('請啟用 Google Sheets API（v4）。');
  const response = Sheets.Spreadsheets.get(SPREADSHEET_ID, {
    ranges: configs.map((config) => "'" + config.sheetName.replace(/'/g, "''") + "'!" + config.range),
    fields: 'sheets(properties(title),merges(startRowIndex,endRowIndex,startColumnIndex,endColumnIndex),data(startRow,startColumn,rowData(values(formattedValue,hyperlink,userEnteredValue(formulaValue),textFormatRuns(format(link(uri))),chipRuns(chip(richLinkProperties(uri)))))))',
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
        if (columnIndex < 2) {
          const mapUrl = mapChipUrl_(value);
          if (mapUrl) cell.mapUrl = mapUrl;
        }
        if (config.readFirstColumnImages && columnIndex === 0) {
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

function getTripDetails() {
  const startedAt = Date.now();
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = spreadsheet.getSheets();
  const spots = sheets.filter((sheet) => /spot$/i.test(sheet.getName().trim()));
  const hotelSheet = sheets.find((sheet) => sheet.getName().trim().toLowerCase() === 'hotel');
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
  ];
  let result;
  try { result = readGridRanges_(configs); }
  catch (error) {
    console.warn('旅遊資料批次讀取未啟用，改用相容讀取：', error.message || String(error));
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
      readMapChipUrls_(configs.map((config, index) => ({ name: config.sheetName, rows: result[index] })));
    } catch (error) {
      mapLocationError = '地圖智慧方塊讀取失敗：' + (error.message || String(error));
      console.warn(mapLocationError);
    }
  }
  const food = trimEmptyEdges_(result[0]);
  const spotSheets = spots.map((sheet, index) => ({
    name: sheet.getName().trim(), rows: trimEmptyEdges_(result[index + 1]),
  }));
  const hotel = hotelSheet ? trimEmptyEdges_(result[result.length - 1]) : [];
  const mapLocations = collectMapLocations_(spotSheets, hotel, food);
  console.info('details: ' + (usedBatch ? 'Sheets API batch + image columns' : 'compatibility reader'), Date.now() - startedAt);
  return { food, spotSheets, hotel, mapLocationError, mapLocations };
}

function imageFormulaUrl_(formula) {
  const match = String(formula || '').match(/^=\s*IMAGE\(\s*"((?:[^"]|"")*)"\s*(?:[,;)]|$)/i);
  const url = match ? match[1].replace(/""/g, '"') : '';
  return /^https?:\/\//i.test(url) ? url : '';
}

// Native CellImage URLs are unavailable in Sheets API. Share the overlay
// between both readers, keeping IMAGE formula URLs when content URLs are absent.
function applyFirstColumnImages_(rows, sheet, firstRow = 1) {
  const images = sheet.getRange(firstRow, 1, rows.length, 1).getValues();
  images.forEach((row, index) => {
    const value = row[0];
    if (!value || value.valueType !== SpreadsheetApp.ValueType.IMAGE) return;
    const imageUrl = value.getContentUrl() || rows[index][0].imageUrl || '';
    if (/^https?:\/\//i.test(imageUrl)) {
      rows[index][0] = { text: '', url: '', imageUrl,
        imageAlt: value.getAltTextDescription() || value.getAltTextTitle() || '' };
    }
  });
}

function mapChipUrl_(value) {
  const urls = [...new Set((value.chipRuns || [])
    .map((run) => run.chip?.richLinkProperties?.uri || '')
    .filter((url) => MAP_URL_PATTERN.test(url)))];
  // Multiple places in a cell do not identify a single route endpoint.
  return urls.length === 1 ? urls[0] : '';
}

function trimEmptyEdges_(rows) {
  const hasData = (cell) => cell.text || cell.url || cell.mapUrl || cell.imageUrl;
  rows.forEach((row) => {
    while (row.length && !hasData(row[row.length - 1])) row.pop();
  });
  while (rows.length && !rows[rows.length - 1].length) rows.pop();
  return rows;
}

// Place smart chips store their URLs in chipRuns, not RichTextValue hyperlinks.
// Read only the two location columns, in a single Sheets API request.
function readMapChipUrls_(mapSheets) {
  const sheets = mapSheets.filter((sheet) => sheet.rows.length);
  if (!sheets.length) return;
  if (typeof Sheets === 'undefined') {
    throw new Error('請在 Apps Script「服務」加入 Google Sheets API（v4）後重新部署。');
  }
  const byName = new Map(sheets.map((sheet) => [sheet.name.trim(), sheet.rows]));
  const response = Sheets.Spreadsheets.get(SPREADSHEET_ID, {
    ranges: sheets.map((sheet) => "'" + sheet.name.replace(/'/g, "''") + "'!A1:B" + sheet.rows.length),
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
          });
        }
      }
    });
  }

  const rows = values.map((row, rowIndex) => row.map((text, columnIndex) => {
    if (config.readFirstColumnImages && rangeColumn + columnIndex === 1) {
      const imageUrl = imageFormulaUrl_(formulas[rowIndex][columnIndex]);
      if (imageUrl) return { text: '', url: '', imageUrl, imageAlt: '' };
    }
    const mergedCell = mergedCells.get(`${rowIndex}:${columnIndex}`);
    if (mergedCell) return mergedCell;

    const richText = richTextValues[rowIndex][columnIndex];
    return {
      text: text || '',
      url: richCellUrl_(richText, formulas[rowIndex][columnIndex]),
    };
  }));
  if (config.readFirstColumnImages && rangeColumn === 1) applyFirstColumnImages_(rows, sheet, rangeRow);
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
