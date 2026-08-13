function normalizeSheetValue(value) {
  return String(value || '').trim();
}

function getAssetPrefixFromText(sourceText) {
  var text = normalizeSheetValue(sourceText || 'AST').replace(/[^A-Za-z0-9]/g, ' ');
  var words = text.split(/\s+/).filter(function(word) { return word.length > 0; });
  if (words.length === 0) return 'AST';
  if (words.length === 1) return words[0].slice(0, 3).toUpperCase() || 'AST';

  var initials = words.slice(0, 3).map(function(word) {
    return word.charAt(0).toUpperCase();
  }).join('');

  return (initials || 'AST').slice(0, 3).toUpperCase() || 'AST';
}

function padAssetNumber(value) {
  return String(value).padStart(3, '0');
}

function generateAssetId(prefix, sequenceNumber) {
  return prefix + '-' + padAssetNumber(sequenceNumber);
}

function getNextAssetNumber(sheet, prefix) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxNumber = 0;

  values.forEach(function(row) {
    var idValue = normalizeSheetValue(row[0]);
    if (!idValue) return;

    var match = idValue.match(new RegExp('^' + prefix + '-(\\d+)$'));
    if (match && match[1]) {
      var numericValue = parseInt(match[1], 10);
      if (numericValue > maxNumber) {
        maxNumber = numericValue;
      }
    }
  });

  return maxNumber;
}

function applyConditionValidation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var totalRows = Math.max(1, sheet.getLastRow() - 1);
  var conditionRange = sheet.getRange(2, 7, totalRows, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Baik', 'Rusak Ringan', 'Rusak Parah'], true)
    .setAllowInvalid(false)
    .build();
  conditionRange.setDataValidation(rule);
}

function ensureAssetIdsForExistingRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  var changed = false;

  rows.forEach(function(row, index) {
    var currentId = normalizeSheetValue(row[0]);
    var hasData = row.slice(1).some(function(value) {
      return normalizeSheetValue(value) !== '';
    });

    if (!currentId && hasData) {
      var prefix = getAssetPrefixFromText(row[3] || row[1] || 'AST');
      var nextNumber = getNextAssetNumber(sheet, prefix) + 1;
      rows[index][0] = generateAssetId(prefix, nextNumber);
      changed = true;
    }
  });

  if (changed) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
}

function ensureSheetHeader() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var firstCell = normalizeSheetValue(sheet.getRange(1, 1).getValue());

  if (firstCell !== 'ID Aset') {
    var legacyValues = sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), Math.min(6, sheet.getLastColumn())).getValues();
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1, 1, 7).setValues([
      ['ID Aset', 'Nama Alat', 'Merek', 'Kategori', 'Jumlah', 'Spesifikasi', 'Kondisi']
    ]);

    if (legacyValues.length > 1) {
      var migratedRows = legacyValues.slice(1).map(function(row) {
        return [
          '',
          row[0] || '',
          row[1] || '',
          row[2] || '',
          row[3] || '',
          row[4] || '',
          row[5] || ''
        ];
      });

      if (migratedRows.length > 0) {
        sheet.getRange(2, 1, migratedRows.length, 7).setValues(migratedRows);
      }
    }
  }

  ensureAssetIdsForExistingRows();
  applyConditionValidation();
}

function findAssetRowById(sheet, assetId) {
  var values = sheet.getRange(2, 1, Math.max(0, sheet.getLastRow() - 1), 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (normalizeSheetValue(values[i][0]) === assetId) {
      return i + 2;
    }
  }
  return -1;
}

function insertItemizedAssetRows(data) {
  var quantity = parseInt(data.jumlah, 10) || 0;
  if (quantity <= 0) {
    throw new Error('Jumlah harus lebih dari 0.');
  }

  ensureSheetHeader();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var prefix = getAssetPrefixFromText(data.kategori || data.namaAlat || 'AST');
  var startIndex = getNextAssetNumber(sheet, prefix) + 1;
  var rows = [];

  for (var i = 0; i < quantity; i++) {
    rows.push([
      generateAssetId(prefix, startIndex + i),
      data.namaAlat || '',
      data.merek || '',
      data.kategori || '',
      1,
      data.spesifikasi || '',
      data.kondisi || 'Baik'
    ]);
  }

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  }

  applyConditionValidation();
}

function doGet() {
  ensureSheetHeader();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var rows = data.slice(1);

  var result = rows
    .filter(function(row) {
      return row.some(function(value) {
        return normalizeSheetValue(value) !== '';
      });
    })
    .map(function(row, index) {
      var normalizedRow = row.length >= 7 ? row.slice(0, 7) : row.concat(Array(7 - row.length).fill(''));
      return {
        rowIndex: index + 2,
        idAset: normalizeSheetValue(normalizedRow[0]),
        namaAlat: normalizedRow[1] || '',
        merek: normalizedRow[2] || '',
        kategori: normalizedRow[3] || '',
        jumlah: parseInt(normalizedRow[4], 10) || 0,
        spesifikasi: normalizedRow[5] || '',
        kondisi: normalizedRow[6] || 'Baik'
      };
    });

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    ensureSheetHeader();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    if (data.action === 'updateKondisi') {
      var assetId = normalizeSheetValue(data.idAset || data.assetId || '');
      if (!assetId) {
        throw new Error('ID Aset tidak valid.');
      }

      var targetRow = findAssetRowById(sheet, assetId);
      if (targetRow === -1) {
        throw new Error('ID Aset tidak ditemukan di Google Sheets.');
      }

      sheet.getRange(targetRow, 7).setValue(data.kondisi || 'Baik');
      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Kondisi berhasil diperbarui!'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'initializeSheet') {
      ensureSheetHeader();
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    insertItemizedAssetRows(data);

    return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}