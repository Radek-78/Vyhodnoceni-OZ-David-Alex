/**
 * ══════════════════════════════════════════════
 * ModuleKT.gs — Modul: Nový KT soubor
 * Vyhodnocení OZ | Lidl interní nástroj
 *
 * Vytváří nový Google Sheets soubor
 * zkopírováním aktuálního souboru. První kopie
 * vzniká z mustru, další kopie vznikají ze souboru
 * předchozího dne a drží se stejné KT složky.
 * ══════════════════════════════════════════════
 */

const ModuleKT = {

  /**
   * Vytvoření nového KT souboru.
   * Zdroj = aktuální tabulka. Pokud aktuální tabulka ještě nemá
   * ktSeriesFolderId, jde o první kopii z mustru a smí vytvořit podsložky.
   * Pokud ktSeriesFolderId existuje, další kopie se vytvoří ve stejné složce.
   * @param {Object} params
   * @param {number} params.week - číslo kalendářního týdne
   * @param {number} params.year - rok
   * @param {string} [params.nameMode] - 'dayMorning' nebo 'date'
   * @param {string} [params.copyDate] - datum ve formátu yyyy-MM-dd
   * @param {string} [params.targetFolderId] - URL nebo ID cílové složky (prázdné = auto)
   * @param {boolean} [params.useSubfolders] - true = vytvořit rok/KTxx podsložky
   * @returns {{ success: boolean, fileId?: string, fileName?: string, url?: string, error?: string }}
   */
  create(params) {
    params = params || {};
    AppLogger.info('═══ Nový KT soubor ═══');

    try {
      var week = params.week;
      var year = params.year;
      var nameMode = params.nameMode || 'date';
      var copyDate = ModuleKT.parseCopyDate_(params.copyDate);
      var useSubfolders = (params.useSubfolders === true || params.useSubfolders === 'true');

      var fileName = ModuleKT.buildFileName_(week, nameMode, copyDate);

      AppLogger.info('Týden: KT ' + String(week).padStart(2, '0') + ' / ' + year);
      AppLogger.info('Název souboru: ' + fileName);
      AppLogger.info('Režim názvu: ' + (nameMode === 'dayMorning' ? 'den DOPOLEDNE' : 'datum'));

      // Zdroj = aktuální tabulka
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sourceFile = DriveApp.getFileById(ss.getId());
      AppLogger.ok('Zdroj: ' + sourceFile.getName() + ' ✓');

      // Uložení preferencí pro příště
      ModuleKT.savePreferences_(params);

      // Cílová složka. Jakmile existuje ktSeriesFolderId, další kopie už
      // nevytváří nové podsložky a drží se stejné složky jako první KT kopie.
      var existingSeriesFolderId = AppConfig.get('ktSeriesFolderId');
      var targetFolder;
      var isFirstSeriesCopy = !existingSeriesFolderId;

      if (existingSeriesFolderId) {
        try {
          targetFolder = DriveApp.getFolderById(existingSeriesFolderId);
          AppLogger.dim('Navazuji na existující KT složku: ' + targetFolder.getName());
        } catch (e) {
          AppLogger.warn('Uložená KT složka nebyla nalezena, použiji složku aktuálního souboru.');
          existingSeriesFolderId = '';
          isFirstSeriesCopy = true;
        }
      }

      if (!existingSeriesFolderId) {
        var baseFolder = ModuleKT.resolveBaseFolder_(params.targetFolderId, sourceFile);
        AppLogger.dim('Výchozí složka: ' + baseFolder.getName());

        targetFolder = baseFolder;
        if (useSubfolders) {
          var yearFolder = ModuleKT.getOrCreateSubfolder_(baseFolder, String(year));
          var ktFolder = ModuleKT.getOrCreateSubfolder_(yearFolder, 'KT' + String(week).padStart(2, '0'));
          targetFolder = ktFolder;
          AppLogger.dim('Cesta: ' + baseFolder.getName() + '/' + String(year) + '/KT' + String(week).padStart(2, '0'));
        }
      }

      AppLogger.dim('Podsložky: ' + (isFirstSeriesCopy && useSubfolders ? 'ano, pouze první kopie' : 'ne'));

      // Kopírování zdroje
      AppLogger.info('Kopíruji soubor...');
      var newFile = sourceFile.makeCopy(fileName, targetFolder);
      AppLogger.ok('Soubor vytvořen: ' + fileName + ' ✓');

      // Otevření a nastavení parametrů
      AppLogger.info('Nastavuji parametry...');
      var newSS = SpreadsheetApp.open(newFile);

      // Pokud nový soubor má _Config list, nastavíme week/year
      var configSheet = newSS.getSheetByName('_Config');
      if (configSheet) {
        var cfgLastRow = configSheet.getLastRow();
        var data = cfgLastRow > 0 ? configSheet.getRange(1, 1, cfgLastRow, 2).getValues() : [];
        var updates = {
          ktWeek: week,
          ktYear: year,
          ktSeriesFolderId: targetFolder.getId(),
          ktLastCopyDate: ModuleKT.formatDate_(copyDate),
          ktLastNameMode: nameMode,
          ktSourceFileId: ss.getId()
        };

        for (var i = 0; i < data.length; i++) {
          var key = data[i][0];
          if (key in updates) {
            configSheet.getRange(i + 1, 2).setValue(updates[key]);
            delete updates[key];
          }
        }

        for (var updateKey in updates) {
          var lr = configSheet.getLastRow() + 1;
          configSheet.getRange(lr, 1, 1, 2).setValues([[updateKey, updates[updateKey]]]);
        }

        AppLogger.ok('Parametry nastaveny (KT' + String(week).padStart(2, '0') + '/' + year + ') ✓');
      } else {
        AppLogger.dim('_Config list v novém souboru nenalezen — přeskakuji parametry');
      }

      var url = newFile.getUrl();
      AppLogger.ok('═══ KT soubor připraven ═══');
      AppLogger.info('URL: ' + url);

      return {
        success: true,
        fileId: newFile.getId(),
        fileName: fileName,
        url: url,
        targetFolderId: targetFolder.getId(),
        targetFolderName: targetFolder.getName(),
        firstSeriesCopy: isFirstSeriesCopy
      };

    } catch (e) {
      AppLogger.error('Chyba: ' + e.message);
      return {
        success: false,
        error: e.message
      };
    }
  },

  /**
   * Složí pevný název KT souboru.
   * @private
   * @param {number} week
   * @param {string} nameMode
   * @param {Date} copyDate
   * @returns {string}
   */
  buildFileName_(week, nameMode, copyDate) {
    var label = nameMode === 'dayMorning'
      ? ModuleKT.getCzechDayName_(copyDate) + ' DOPOLEDNE'
      : Utils.formatDate(copyDate, 'dd.MM.yyyy');
    return 'Vyhodnocení akce KT ' + Number(week) + ' - ' + label;
  },

  /**
   * @private
   * @param {string|Date} value
   * @returns {Date}
   */
  parseCopyDate_(value) {
    if (value instanceof Date) return value;
    if (value) {
      var m = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    }
    return new Date();
  },

  /**
   * @private
   * @param {Date} date
   * @returns {string}
   */
  formatDate_(date) {
    return Utilities.formatDate(date, 'Europe/Prague', 'yyyy-MM-dd');
  },

  /**
   * @private
   * @param {Date} date
   * @returns {string}
   */
  getCzechDayName_(date) {
    var names = ['NEDĚLE', 'PONDĚLÍ', 'ÚTERÝ', 'STŘEDA', 'ČTVRTEK', 'PÁTEK', 'SOBOTA'];
    return names[date.getDay()];
  },

  /**
   * Resolve cílové složky — uživatelský přepis nebo složka mustru
   * @private
   * @param {string} targetFolderInput - URL nebo ID nebo prázdné
   * @param {GoogleAppsScript.Drive.File} templateFile
   * @returns {GoogleAppsScript.Drive.Folder}
   */
  resolveBaseFolder_(targetFolderInput, templateFile) {
    // 1) Přímý vstup z formuláře (URL nebo ID)
    if (targetFolderInput && String(targetFolderInput).trim()) {
      var trimmed = String(targetFolderInput).trim();
      var folderId = trimmed;

      if (trimmed.startsWith('http')) {
        var extracted = Utils.extractFileIdFromUrl(trimmed);
        if (extracted) folderId = extracted;
      }

      try {
        return DriveApp.getFolderById(folderId);
      } catch (e) {
        AppLogger.warn('Zadaná složka nenalezena, zkouším nastavení...');
      }
    }

    // 2) Uloženou override z nastavení
    var ktFolderOverride = AppConfig.get('ktFolderId');
    if (ktFolderOverride) {
      try {
        return DriveApp.getFolderById(ktFolderOverride);
      } catch (e) {
        AppLogger.warn('Nastavená složka nenalezena, používám výchozí');
      }
    }

    // 3) Výchozí = složka tabulky
    var autoId = Utils.getBaseFolderId();
    if (autoId) {
      try {
        return DriveApp.getFolderById(autoId);
      } catch (e) {
        // Fallback níže
      }
    }

    return DriveApp.getRootFolder();
  },

  /**
   * Najde nebo vytvoří podsložku v dané složce
   * @private
   * @param {GoogleAppsScript.Drive.Folder} parent
   * @param {string} name
   * @returns {GoogleAppsScript.Drive.Folder}
   */
  getOrCreateSubfolder_(parent, name) {
    var folders = parent.getFoldersByName(name);
    if (folders.hasNext()) {
      return folders.next();
    }
    AppLogger.dim('Vytvářím podsložku: ' + name);
    return parent.createFolder(name);
  },

  /**
   * Uloží uživatelské preference do _Config
   * @private
   * @param {Object} params
   */
  savePreferences_(params) {
    try {
      // Dávkový zápis — 1 čtení + 1 zápis místo 4 × (čtení + zápis)
      const updates = {};
      if (params.useSubfolders !== undefined) updates['ktUseSubfolders'] = params.useSubfolders ? '1' : '0';
      if (params.targetFolderId !== undefined) updates['ktTargetFolderId'] = params.targetFolderId || '';
      if (params.nameMode !== undefined) updates['ktNameMode'] = params.nameMode || 'date';
      if (Object.keys(updates).length > 0) AppConfig.setMultiple(updates);
    } catch (e) {
      // Preferences save is non-critical
      AppLogger.dim('Preference neuloženy: ' + e.message);
    }
  },

  /**
   * Načtení uložených preferencí
   * @returns {Object}
   */
  getPreferences() {
    var subVal = AppConfig.get('ktUseSubfolders');
    // Default = true (zaškrtnuto) — pokud klíč neexistuje (null), default je true
    // Sheets může konvertovat '1' na číslo 1, proto porovnáváme přes String()
    var useSub;
    if (subVal === null || subVal === undefined || subVal === '') {
      useSub = true; // default
    } else {
      useSub = (String(subVal) === '1' || subVal === true);
    }
    return {
      useSubfolders: useSub,
      targetFolderId: AppConfig.get('ktTargetFolderId') || '',
      nameMode: AppConfig.get('ktNameMode') || 'date',
      seriesFolderId: AppConfig.get('ktSeriesFolderId') || ''
    };
  },

  /**
   * Získání informací o šabloně (= aktuální tabulka) + auto složka
   * @returns {{ configured: boolean, name?: string, id?: string, autoFolderName?: string, autoFolderId?: string }}
   */
  getTemplateInfo() {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var autoFolderId = Utils.getBaseFolderId() || '';
      var autoFolderName = 'Kořen Drive';
      if (autoFolderId) {
        try {
          autoFolderName = DriveApp.getFolderById(autoFolderId).getName();
        } catch (e) {
          autoFolderName = 'Složka nenalezena';
        }
      }

      return {
        configured: true,
        name: ss.getName(),
        id: ss.getId(),
        autoFolderName: autoFolderName,
        autoFolderId: autoFolderId
      };
    } catch (e) {
      return {
        configured: false,
        error: 'Nelze získat info o aktuální tabulce'
      };
    }
  }
  ,

  /**
   * Aktuální KT kontext pro formulář.
   * @returns {{week: number|string, year: number|string}}
   */
  getCurrentContext() {
    var current = Utils.getCurrentWeek();
    var configuredWeek = parseInt(AppConfig.get('ktWeek'), 10);
    var configuredYear = parseInt(AppConfig.get('ktYear'), 10);

    return {
      week: (!isNaN(configuredWeek) && configuredWeek >= 1 && configuredWeek <= 53) ? configuredWeek : current.week,
      year: (!isNaN(configuredYear) && configuredYear >= 2020 && configuredYear <= 2099) ? configuredYear : current.year
    };
  }

};


// ═══ Top-level wrappery pro google.script.run ═══

function moduleKT_create(params) {
  return ModuleKT.create(params);
}

function moduleKT_getTemplateInfo() {
  return ModuleKT.getTemplateInfo();
}

function moduleKT_getPreferences() {
  return ModuleKT.getPreferences();
}

function moduleKT_getCurrentContext() {
  return ModuleKT.getCurrentContext();
}
