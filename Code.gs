const SPREADSHEET_ID = '1sJZjqZ60cUJrMvfbPUM6EwoG4wZf5S-BTWs_5or5yEw';

const SOURCE_RANGES = {
  schedule: { sheetName: 'schedule', range: 'A1:K100', expandMergedCells: true },
  food: { sheetName: 'food', range: 'A1:K100' },
};

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('岡山保佑一定好天氣')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Return only the travel data needed by Index.html.
 * The spreadsheet stays private because this function runs as the deployer.
 */
function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getTripData() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const spotSheets = collectSpotSheets_(spreadsheet);
  const food = readRichRange_(spreadsheet, SOURCE_RANGES.food);
  const hotelSheet = spreadsheet.getSheets().find(
    (sheet) => sheet.getName().trim().toLowerCase() === 'hotel'
  );
  const hotel = hotelSheet ? readRichRange_(spreadsheet, {
    sheetName: hotelSheet.getName(),
    range: hotelSheet.getDataRange().getA1Notation(),
  }) : [];

  let mapLocationError = '';
  try {
    const mapSheets = spotSheets.concat([{ name: SOURCE_RANGES.food.sheetName, rows: food }]);
    if (hotelSheet) mapSheets.push({ name: hotelSheet.getName(), rows: hotel });
    readMapChipUrls_(mapSheets);
  } catch (error) {
    // A Maps integration problem must not prevent the itinerary from loading.
    mapLocationError = '地圖智慧方塊讀取失敗：' + (error.message || String(error));
    console.warn(mapLocationError);
  }

  return {
    schedule: readRichRange_(spreadsheet, SOURCE_RANGES.schedule),
    food,
    // Keep older deployed HTML working while the frontend files are updated.
    // These aliases use discovered sheets; neither sheet is required to exist.
    kobeSpot: (spotSheets.find((sheet) => sheet.name.toLowerCase() === 'kobe spot') || {}).rows || [],
    spot: (spotSheets.find((sheet) => sheet.name.toLowerCase() === 'spot') || {}).rows || [],
    spotSheets,
    mapLocations: collectMapLocations_(spotSheets, hotel, food),
    mapLocationError,
    hotel,
    spotReferences: collectSpotReferences_(spreadsheet, spotSheets, hotel),
    updatedAt: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm'
    ),
  };
}

const MAP_URL_PATTERN = /^https:\/\/(?:maps\.app\.goo\.gl\/|goo\.gl\/maps(?:\/|\?)|(?:www\.|maps\.)?google\.(?:com|co\.jp|com\.tw)\/(?:maps(?:\/|\?)|\?))/i;

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
          const urls = [...new Set((value.chipRuns || [])
            .map((run) => run.chip?.richLinkProperties?.uri || '')
            .filter((url) => MAP_URL_PATTERN.test(url)))];
          // Multiple places in a cell do not identify a single route endpoint.
          if (urls.length === 1) cell.mapUrl = urls[0];
        });
      });
    });
  });
}

// Initial page loading must never wait for Google Maps network requests.
function collectMapLocations_(spotSheets, hotelRows, foodRows = []) {
  const resolved = new Map();
  const locations = [];
  const cache = CacheService.getScriptCache();
  spotSheets.map((sheet) => sheet.rows).concat([hotelRows, foodRows]).forEach((rows) => {
    rows.forEach((row) => {
      const name = (row[0]?.text || row[1]?.text || '').trim();
      const secondary = row[1] || {};
      const url = [secondary.mapUrl, secondary.url, secondary.text, row[0]?.mapUrl, row[0]?.url, row[0]?.text]
        .map((value) => String(value || '').trim()).find((value) => MAP_URL_PATTERN.test(value)) || '';
      if (!name || !url) return;
      if (!resolved.has(url)) {
        resolved.set(url, cache.get('map-url:v2:' + url) || url);
      }
      locations.push({ name, alias: secondary.text || '', originalUrl: url, url: resolved.get(url) });
    });
  });
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
 * Read every sheet whose name ends with "spot" so each one can be rendered
 * and used as a schedule-name reference.
 */
function collectSpotSheets_(spreadsheet) {
  return spreadsheet.getSheets()
    .filter((sheet) => /spot$/i.test(sheet.getName().trim()))
    .map((sheet) => ({
      name: sheet.getName().trim(),
      rows: readRichRange_(spreadsheet, {
        sheetName: sheet.getName(),
        range: sheet.getDataRange().getA1Notation(),
        readFirstColumnImages: true,
      }),
    }));
}

/**
 * Collect names from every spot sheet and the optional Hotel reference sheet.
 */
function collectSpotReferences_(spreadsheet, spotSheets, hotelRows) {
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
function readRichRange_(spreadsheet, config) {
  const sheet = spreadsheet.getSheetByName(config.sheetName);

  if (!sheet) {
    throw new Error('找不到工作表：' + config.sheetName);
  }

  const range = sheet.getRange(config.range);
  const values = range.getDisplayValues();
  const richTextValues = range.getRichTextValues();
  const formulas = range.getFormulas();
  const imageColumn = config.readFirstColumnImages
    ? sheet.getRange(range.getRow(), 1, range.getNumRows(), 1) : null;
  const imageValues = imageColumn ? imageColumn.getValues() : [];
  const imageFormulas = imageColumn ? imageColumn.getFormulas() : [];
  const mergedCells = new Map();

  if (config.expandMergedCells) {
    range.getMergedRanges().forEach((mergedRange) => {
      const topLeft = mergedRange.getCell(1, 1);
      const mergedText = topLeft.getDisplayValue() || '';
      const mergedRichText = topLeft.getRichTextValue();
      const mergedUrl = mergedRichText ? mergedRichText.getLinkUrl() || '' : '';
      const firstRow = Math.max(0, mergedRange.getRow() - range.getRow());
      const lastRow = Math.min(
        values.length - 1,
        mergedRange.getLastRow() - range.getRow()
      );
      const firstColumn = Math.max(0, mergedRange.getColumn() - range.getColumn());
      const lastColumn = Math.min(
        values[0].length - 1,
        mergedRange.getLastColumn() - range.getColumn()
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

  return values.map((row, rowIndex) => row.map((text, columnIndex) => {
    if (imageColumn && range.getColumn() + columnIndex === 1) {
      const value = imageValues[rowIndex][0];
      const formula = imageFormulas[rowIndex][0] || '';
      const imageFormula = formula.match(/^=\s*IMAGE\(\s*"((?:[^"]|"")*)"\s*(?:[,;)]|$)/i);
      let imageUrl = imageFormula ? imageFormula[1].replace(/""/g, '"') : '';
      let imageAlt = '';
      if (value && value.valueType === SpreadsheetApp.ValueType.IMAGE) {
        imageUrl = value.getContentUrl() || imageUrl;
        imageAlt = value.getAltTextDescription() || value.getAltTextTitle() || '';
      }
      if (/^https?:\/\//i.test(imageUrl)) {
        return { text: '', url: '', imageUrl, imageAlt };
      }
    }
    const mergedCell = mergedCells.get(`${rowIndex}:${columnIndex}`);
    if (mergedCell) return mergedCell;

    const richText = richTextValues[rowIndex][columnIndex];
    return {
      text: text || '',
      url: richCellUrl_(richText, formulas[rowIndex][columnIndex]),
    };
  }));
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
