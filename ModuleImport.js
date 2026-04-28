/**
 * ══════════════════════════════════════════════
 * ModuleImport.gs — Modul: Import souborů
 * Vyhodnocení OZ | Lidl interní nástroj
 *
 * KOMPLETNÍ REWORK PRO MAXIMÁLNÍ STABILITU
 * ══════════════════════════════════════════════
 */

// Regex kompilovaný jednou pro celý modul — ne uvnitř map() kde by se re-kompiloval per buňce
const _ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/;

const ModuleImport = {

  CASH_COLUMNS: ['ITEMNR', 'DESCRIPTION', 'ITEMGROUP', 'ITEMCOUNT', 'KG', 'CZK', 'STORE'],
  CASH_DATE_COLUMN: 3,
  CASH_FIRST_DATA_COLUMN: 4,
  CASH_CLEAR_FIRST_COLUMN: 3,
  CASH_CLEAR_LAST_COLUMN: 10,

  /**
   * Import datové dávky do listu
   * @param {string} jsonPayload - JSON řetězec s 2D polem dat
   * @param {string} sheetName - název cílového listu
   * @param {Object} options - { isFirstChunk: boolean, overwrite: boolean }
   */
  importChunk(jsonPayload, sheetName, options) {
    let data;
    try {
      data = JSON.parse(jsonPayload);

      // Detekce a převod ISO dat zpět na objekty Date
      data = data.map(row => row.map(val => {
        if (typeof val === 'string' && _ISO_DATE_RE.test(val)) {
          const d = new Date(val);
          return isNaN(d.getTime()) ? val : d;
        }
        return val;
      }));
    } catch (e) {
      return { success: false, error: 'Chyba při čtení datového balíčku: ' + e.message };
    }

    if (!data || data.length === 0) return { success: true, count: 0 };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);

    // 1. Příprava listu (pouze u první dávky souboru)
    if (options.isFirstChunk) {
      if (sheet) {
        if (options.overwrite) {
          AppLogger.info('Mazání listu: ' + sheetName);
          sheet.clear();
        } else {
          AppLogger.info('Přidávám k existujícím datům: ' + sheetName);
        }
      } else {
        AppLogger.info('Vytvářím nový list: ' + sheetName);
        sheet = ss.insertSheet(sheetName);
      }
    }

    if (!sheet) throw new Error('Nepodařilo se získat cílový list: ' + sheetName);

    // 2. Zjištění startovní pozice a rozměrů
    let startRow = (options.startRow !== undefined) ? options.startRow : (sheet.getLastRow() + 1);
    const numRows = data.length;
    const maxCols = data.reduce((max, row) => Math.max(max, row.length), 0);
    if (maxCols === 0) return { success: true, count: 0, actualStartRow: startRow };

    // AUTOMATICKÉ ROZŠÍŘENÍ LISTU
    const maxRows = sheet.getMaxRows();
    const neededRows = startRow + numRows - 1;
    const totalToPrepare = (options.isFirstChunk && options.totalRows) ? (options.totalRows) : neededRows;

    if (totalToPrepare > maxRows) {
      const toAdd = totalToPrepare - maxRows;
      sheet.insertRowsAfter(maxRows, toAdd);
      AppLogger.info('Příprava prostoru: +' + toAdd + ' řádků.');
    }

    // 3. PŘÍPRAVA DAT PRO ADVANCED SHEETS API (v4)
    const normalizedData = data.map(row => {
      const r = row.slice(0, maxCols);
      while (r.length < maxCols) r.push('');
      return r;
    });

    try {
      const spreadsheetId = ss.getId();
      const range = sheetName + '!' + Utils.columnToLetter(1) + startRow + ':' + Utils.columnToLetter(maxCols) + (startRow + numRows - 1);

      const valueRange = Sheets.newValueRange();
      valueRange.values = normalizedData;

      Sheets.Spreadsheets.Values.update(valueRange, spreadsheetId, range, { valueInputOption: 'RAW' });

      // Formátování záhlaví (podmíněno parametrem formatHeader)
      if (startRow === 1 && options.formatHeader !== false) {
        const header = sheet.getRange(1, 1, 1, maxCols);
        header.setFontWeight('bold');
        header.setBackground('#eeeeee');
        sheet.setFrozenRows(1);
      }

      return { success: true, count: numRows, actualStartRow: startRow };
    } catch (e) {
      AppLogger.error('Chyba V4 (fallback): ' + e.message);
      try {
        sheet.getRange(startRow, 1, numRows, maxCols).setValues(normalizedData);
        return { success: true, count: numRows, actualStartRow: startRow };
      } catch (err) {
        return { success: false, error: 'Kritické selhání zápisu: ' + err.message };
      }
    }
  },

  /**
   * Aktivní artikly pro filtr pokladních dat.
   * @returns {{success: boolean, articles?: string[], error?: string}}
   */
  getActiveArticles() {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName('Vyhodnocení akce');
      if (!sheet) return { success: false, error: 'List "Vyhodnocení akce" nebyl nalezen.' };

      const values = sheet.getRange('A6:A28').getValues();
      const articles = values
        .map(row => ModuleImport.normalizeArticle_(row[0]))
        .filter(value => value !== '');

      return { success: true, articles: articles };
    } catch (e) {
      return { success: false, error: 'Nelze načíst aktivní artikly: ' + e.message };
    }
  },

  /**
   * Specializovaný zápis pokladních dat.
   * Řádek 1 s hlavičkou v cílovém listu nikdy nemaže.
   * @param {string} jsonPayload
   * @param {string} sheetName
   * @param {Object} options
   * @returns {{success: boolean, count?: number, actualStartRow?: number, error?: string}}
   */
  importCashChunk(jsonPayload, sheetName, options) {
    options = options || {};
    let data;
    try {
      data = JSON.parse(jsonPayload);
    } catch (e) {
      return { success: false, error: 'Chyba při čtení pokladních dat: ' + e.message };
    }

    if (!data || data.length === 0) return { success: true, count: 0, actualStartRow: 2 };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, ModuleImport.CASH_DATE_COLUMN).setValue('Datum');
      sheet.getRange(1, ModuleImport.CASH_FIRST_DATA_COLUMN, 1, ModuleImport.CASH_COLUMNS.length).setValues([ModuleImport.CASH_COLUMNS]);
      sheet.setFrozenRows(1);
    }

    const header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), ModuleImport.CASH_CLEAR_LAST_COLUMN)).getValues()[0];
    const colMap = ModuleImport.getCashTargetColumnMap_(header);
    const missing = ModuleImport.CASH_COLUMNS.filter(col => !colMap[col]);
    if (missing.length) {
      return { success: false, error: 'V cílovém listu chybí sloupce: ' + missing.join(', ') };
    }

    if (options.isFirstChunk) {
      const maxRows = sheet.getMaxRows();
      if (maxRows > 1) {
        ModuleImport.clearCashData_(sheet, maxRows);
      }
    }

    const startRow = options.startRow !== undefined ? Number(options.startRow) : 2;
    const neededRows = startRow + data.length - 1;
    const preparedRows = (options.isFirstChunk && options.totalRows) ? Number(options.totalRows) + 1 : neededRows;
    const rowsToPrepare = Math.max(neededRows, preparedRows);
    if (rowsToPrepare > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), rowsToPrepare - sheet.getMaxRows());
    }

    const firstCol = Math.min.apply(null, ModuleImport.CASH_COLUMNS.map(col => colMap[col]));
    const lastCol = Math.max.apply(null, ModuleImport.CASH_COLUMNS.map(col => colMap[col]));
    const isContiguous = (lastCol - firstCol + 1) === ModuleImport.CASH_COLUMNS.length &&
      ModuleImport.CASH_COLUMNS.every((col, idx) => colMap[col] === firstCol + idx);

    try {
      const importDate = new Date();
      sheet.getRange(startRow, ModuleImport.CASH_DATE_COLUMN, data.length, 1)
        .setValues(data.map(() => [importDate]))
        .setNumberFormat('dd.MM.yyyy');

      if (isContiguous) {
        sheet.getRange(startRow, firstCol, data.length, ModuleImport.CASH_COLUMNS.length).setValues(data);
      } else {
        ModuleImport.CASH_COLUMNS.forEach((col, idx) => {
          const columnValues = data.map(row => [row[idx]]);
          sheet.getRange(startRow, colMap[col], columnValues.length, 1).setValues(columnValues);
        });
      }
      return { success: true, count: data.length, actualStartRow: startRow };
    } catch (e) {
      return { success: false, error: 'Chyba zápisu pokladních dat: ' + e.message };
    }
  },

  /**
   * Import souborů z Google Drive
   */
  importToDrive(fileIds, options) {
    const results = [];
    const targetFolderId = options.targetFolderId;
    const prefix = options.prefix || '';

    let targetFolder = targetFolderId ? DriveApp.getFolderById(targetFolderId) : DriveApp.getRootFolder();

    fileIds.forEach(id => {
      try {
        const file = DriveApp.getFileById(id);
        const name = prefix + file.getName();
        file.makeCopy(name, targetFolder);
        results.push({ name: name, success: true });
        AppLogger.ok('Kopírováno: ' + name);
      } catch (e) {
        results.push({ name: id, success: false, error: e.message });
      }
    });
    return { success: true, results: results };
  },

  /**
   * Nahrání souboru na Drive
   */
  uploadToDrive(fileObj, options) {
    try {
      const blob = ModuleImport.base64ToBlob_(fileObj.data, fileObj.name);
      const folder = options.targetFolderId ? DriveApp.getFolderById(options.targetFolderId) : DriveApp.getRootFolder();
      const file = folder.createFile(blob);
      if (options.prefix) file.setName(options.prefix + file.getName());
      return { success: true, name: file.getName() };
    } catch (e) {
      return { success: false, error: e.message };
    }
  },

  base64ToBlob_(dataUrl, fileName) {
    const parts = dataUrl.split(',');
    const mime = parts[0].match(/:(.*?);/)[1];
    const decoded = Utilities.base64Decode(parts[1]);
    return Utilities.newBlob(decoded, mime, fileName);
  },

  getCashTargetColumnMap_(header) {
    const map = {};
    header.forEach((cell, idx) => {
      const key = String(cell || '').trim().toUpperCase();
      const col = idx + 1;
      if (col >= ModuleImport.CASH_FIRST_DATA_COLUMN && ModuleImport.CASH_COLUMNS.indexOf(key) !== -1) {
        map[key] = col;
      }
    });
    return map;
  },

  clearCashData_(sheet, maxRows) {
    const width = ModuleImport.CASH_CLEAR_LAST_COLUMN - ModuleImport.CASH_CLEAR_FIRST_COLUMN + 1;
    sheet.getRange(2, ModuleImport.CASH_CLEAR_FIRST_COLUMN, maxRows - 1, width).clearContent();
  },

  normalizeArticle_(value) {
    if (value === null || value === undefined) return '';
    let text = String(value).trim();
    if (/^\d+(\.0+)?$/.test(text)) text = String(parseInt(text, 10));
    return text;
  }
};

/** API Wrappery */
function moduleImport_chunk(data, sheetName, options) {
  return ModuleImport.importChunk(data, sheetName, options);
}

function moduleImport_getActiveArticles() {
  return ModuleImport.getActiveArticles();
}

function moduleImport_cashChunk(data, sheetName, options) {
  return ModuleImport.importCashChunk(data, sheetName, options);
}

function moduleImport_uploadToDrive(fileObj, options) {
  return ModuleImport.uploadToDrive(fileObj, options);
}

function moduleImport_runDrive(fileIds, options) {
  return ModuleImport.importToDrive(fileIds, options);
}
