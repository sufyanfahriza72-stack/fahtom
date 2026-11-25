const SPREADSHEET_ID = '177PCcDP638NfZCodr79eETVR5QcmsHbDDsWu0qEyM8s';
const LOGO_URL = 'https://lh3.googleusercontent.com/d/1v9XLw_NjWKP9zdtIgREkPYmgakhKMq_n';
const TTD_URL  = 'https://lh3.googleusercontent.com/d/1mDzoyRWgLuyOutFn81FJd2XQ_JcAQtCo';

const SHEET_SCHEMA = {
  'Agenda': ['ID', 'Judul', 'Lokasi', 'JatuhTempo', 'Durasi', 'Prioritas', 'Status', 'DisposisiKepada', 'TamuKhusus'],
  'Pelajaran': ['ID', 'MataPelajaran', 'Guru', 'Hari', 'Waktu'],
  'Ujian': ['ID', 'Hari', 'Jam', 'MataPelajaran', 'StatusAktif'],
  'Tugas': ['ID', 'Deadline', 'MataPelajaran', 'Guru', 'DetailTugas'],
  'Notulensi': ['ID', 'Tanggal', 'JudulRapat', 'IsiNotulensi', 'StatusTTE']
};

function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.keys(SHEET_SCHEMA).forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    const headers = SHEET_SCHEMA[name];
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#0f172a')
        .setFontColor('#ffffff');
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }
  });
}

function doGet(e) {
  try {
    setupDatabase();
  } catch (err) {
    Logger.log('Setup warning: ' + err.message);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('MUSAFA Mobile — Sistem Manajemen Agenda & Tugas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .addMetaTag('theme-color', '#0f172a')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-status-bar-style', 'black-translucent')
    .addMetaTag('apple-mobile-web-app-title', 'MUSAFA Mobile')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function verifyLogin(username) {
  return String(username || '').trim() === 'fahriztommy';
}

function verifyPin(pin) {
  return String(pin || '').trim().toUpperCase() === 'FT';
}

function getDynamicTTEToken() {
  const minuteStamp = Math.floor(Date.now() / 60000);
  const raw = 'MUSAFA_TTE_' + minuteStamp + '_FT_SECURE';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  let code = '';
  for (let i = 0; i < 4; i++) {
    let b = digest[i];
    if (b < 0) b += 256;
    code += ('0' + b.toString(16)).slice(-2);
  }
  return code.toUpperCase();
}

function verifyTTEToken(inputToken) {
  const current = getDynamicTTEToken();
  const clean = String(inputToken || '').trim().toUpperCase();
  if (clean && clean === current) {
    return {
      valid: true,
      ttdUrl: TTD_URL,
      signerName: 'Muhammad Sufyan Fahriza',
      signerRole: 'Eselon III (Pejabat Rohis)',
      verifiedAt: new Date().toISOString()
    };
  }
  return { valid: false };
}

function getAllData() {
  setupDatabase();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const result = { Agenda: [], Pelajaran: [], Ujian: [], Tugas: [], Notulensi: [] };

  Object.keys(result).forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() <= 1) return;

    const values = sheet.getDataRange().getValues();
    const headers = values[0];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const obj = {};
      let hasContent = false;
      for (let j = 0; j < headers.length; j++) {
        let cell = row[j];
        if (cell instanceof Date) {
          cell = Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
        }
        obj[headers[j]] = cell !== null && cell !== undefined ? String(cell) : '';
        if (obj[headers[j]] !== '') hasContent = true;
      }
      if (hasContent && obj.ID) {
        result[name].push(obj);
      }
    }
  });

  return result;
}

function saveRecord(sheetName, recordObj) {
  setupDatabase();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan: ' + sheetName);

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const cleanRec = Object.assign({}, recordObj);

  if (!cleanRec.ID || String(cleanRec.ID).trim() === '') {
    cleanRec.ID = sheetName.substring(0, 3).toUpperCase() + '-' + new Date().getTime();
    const newRow = headers.map(function(h) {
      return cleanRec[h] !== undefined ? cleanRec[h] : '';
    });
    sheet.appendRow(newRow);
  } else {
    let found = false;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(cleanRec.ID)) {
        const updatedRow = headers.map(function(h, idx) {
          return cleanRec[h] !== undefined ? cleanRec[h] : values[i][idx];
        });
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([updatedRow]);
        found = true;
        break;
      }
    }
    if (!found) {
      const newRow = headers.map(function(h) {
        return cleanRec[h] !== undefined ? cleanRec[h] : '';
      });
      sheet.appendRow(newRow);
    }
  }

  return getAllData();
}

function deleteRecord(sheetName, id) {
  setupDatabase();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return getAllData();

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  return getAllData();
}