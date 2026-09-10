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
  const hotelSheet = spreadsheet.getSheets().find(
    (sheet) => sheet.getName().trim().toLowerCase() === 'hotel'
  );
  const hotel = hotelSheet ? readRichRange_(spreadsheet, {
    sheetName: hotelSheet.getName(),
    range: hotelSheet.getDataRange().getA1Notation(),
  }) : [];

  return {
    schedule: readRichRange_(spreadsheet, SOURCE_RANGES.schedule),
    food: readRichRange_(spreadsheet, SOURCE_RANGES.food),
    // Keep older deployed HTML working while the frontend files are updated.
    // These aliases use discovered sheets; neither sheet is required to exist.
    kobeSpot: (spotSheets.find((sheet) => sheet.name.toLowerCase() === 'kobe spot') || {}).rows || [],
    spot: (spotSheets.find((sheet) => sheet.name.toLowerCase() === 'spot') || {}).rows || [],
    spotSheets,
    hotel,
    spotReferences: collectSpotReferences_(spreadsheet, spotSheets, hotel),
    updatedAt: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm'
    ),
  };
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
      url: richText ? richText.getLinkUrl() || '' : '',
    };
  }));
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
    updatedAt: data.updatedAt,
  }));
}
