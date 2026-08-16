/***************************************************************
 * AMelur-Thirukkural
 * Code.gs — REAL KURALS DIRECTLY FROM USER'S GOOGLE SHEET
 *
 * SOURCE:
 * Spreadsheet ID:
 * 13JdpPtzsZOZ4s6S9Hfih6lN0c5-0VJTdwBTLBtT0y4k
 *
 * Source range:
 * C1:D1331
 *
 * C1 = கு_எண்
 * D1 = குறள்
 *
 * Rows 2:1331 contain the 1330 real Kurals.
 ***************************************************************/

const KURAL_SOURCE_SPREADSHEET_ID =
  '13JdpPtzsZOZ4s6S9Hfih6lN0c5-0VJTdwBTLBtT0y4k';

const KURAL_SOURCE_START_ROW = 2;
const KURAL_SOURCE_NUM_COLUMN = 3;
const KURAL_SOURCE_TEXT_COLUMN = 4;
const KURAL_SOURCE_COUNT = 1330;
const KURAL_CACHE_SECONDS = 21600;

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle(
      'ஊராட்சி ஒன்றியத் தொடக்கப் பள்ளி, அஞ்சலம் மேலூர் - திருக்குறள் செயலி'
    )
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getKuralsData(numbersArray) {
  if (!Array.isArray(numbersArray)) {
    throw new Error('குறள் எண்கள் சரியான பட்டியலாக இல்லை.');
  }

  const numbers = [...new Set(
    numbersArray
      .map(Number)
      .filter(n => Number.isInteger(n) && n >= 1 && n <= 1330)
  )];

  if (!numbers.length) {
    throw new Error('1 முதல் 1330 வரை செல்லுபடியாகும் குறள் எண் தேவை.');
  }

  const result = {};
  const missing = [];

  numbers.forEach(n => {
    const cached = CacheService.getScriptCache()
      .get('SHEET_REAL_KURAL_' + n);

    if (cached) {
      try {
        const record = JSON.parse(cached);
        if (isValidKuralRecord_(record, n)) {
          result[n] = record;
          return;
        }
      } catch (e) {}
    }
    missing.push(n);
  });

  if (missing.length) {
    const sourceRecords = loadAllSourceKurals_();

    missing.forEach(n => {
      const record = sourceRecords[n];
      if (record && isValidKuralRecord_(record, n)) {
        result[n] = record;
        CacheService.getScriptCache().put(
          'SHEET_REAL_KURAL_' + n,
          JSON.stringify(record),
          KURAL_CACHE_SECONDS
        );
      }
    });
  }

  const unavailable = numbers.filter(n => !result[n]);

  if (unavailable.length) {
    throw new Error(
      'உங்கள் Google Sheet-ல் இந்த குறள் எண்கள் கிடைக்கவில்லை: ' +
      unavailable.join(', ') +
      '. C1:D1331 தரவைச் சரிபார்க்கவும்.'
    );
  }

  return result;
}

function loadAllSourceKurals_() {
  const ss = SpreadsheetApp.openById(KURAL_SOURCE_SPREADSHEET_ID);
  const sheets = ss.getSheets();

  if (!sheets || !sheets.length) {
    throw new Error('குறள் தரவு Spreadsheet-ல் Sheet கிடைக்கவில்லை.');
  }

  const sheet = sheets[0];
  const lastRow = sheet.getLastRow();

  if (lastRow < KURAL_SOURCE_START_ROW) {
    throw new Error('குறள் தரவு Sheet காலியாக உள்ளது.');
  }

  const rowCount = Math.min(
    KURAL_SOURCE_COUNT,
    lastRow - KURAL_SOURCE_START_ROW + 1
  );

  const values = sheet.getRange(
    KURAL_SOURCE_START_ROW,
    KURAL_SOURCE_NUM_COLUMN,
    rowCount,
    2
  ).getDisplayValues();

  const records = {};

  values.forEach(row => {
    const numberText = String(row[0] || '').trim();
    const number = Number(numberText.replace(/[^\d]/g, ''));
    const fullText = String(row[1] || '').trim();

    if (
      !Number.isInteger(number) ||
      number < 1 ||
      number > 1330 ||
      !fullText
    ) {
      return;
    }

    const record = {
      num: number,
      text: fullText.replace(/\s*\|\s*/g, ' ').replace(/\s+/g, ' ').trim(),
      voiceParts: splitVoiceParts_(fullText)
    };

    if (isValidKuralRecord_(record, number)) {
      records[number] = record;
    }
  });

  return records;
}

/**
 * The "|" characters in column D define the Voice Recognition parts.
 * Example:
 * "முதல் பகுதி | இரண்டாம் பகுதி"
 * becomes two recognition boxes.
 */
function splitVoiceParts_(text) {
  const value = String(text || '').trim();

  if (value.indexOf('|') === -1) {
    return [value];
  }

  return value
    .split('|')
    .map(s => s.trim())
    .filter(Boolean)
    .slice(0, 2);
}

function isValidKuralRecord_(record, expectedNumber) {
  if (!record) return false;

  if (Number(record.num) !== Number(expectedNumber)) {
    return false;
  }

  if (typeof record.text !== 'string' || !record.text.trim()) {
    return false;
  }

  if (!Array.isArray(record.voiceParts) || !record.voiceParts.length) {
    return false;
  }

  const combined = record.text.toLowerCase();

  const forbidden = [
    'முதலடி தமிழ் வாசிப்பு',
    'இரண்டாமடி தமிழ் வாசிப்பு',
    'placeholder',
    'demo',
    'sample text'
  ];

  for (let i = 0; i < forbidden.length; i++) {
    if (combined.indexOf(forbidden[i].toLowerCase()) !== -1) {
      return false;
    }
  }

  return true;
}

function getAllKuralsForSearch() {
  return loadAllSourceKurals_();
}

/**
 * Student learning report.
 */
function logUserLearningStatus(
  studentName,
  whatsapp,
  emis,
  kuralNumbers,
  startedAt,
  usedSeconds,
  score
) {
  try {
    const name = String(studentName || '').trim();
    const wa = String(whatsapp || '').trim();
    const emisId = String(emis || '').trim();
    const completed = String(kuralNumbers || '').trim();

    if (!name) return 'மாணவர் பெயர் தேவை.';
    if (!wa) return 'WhatsApp எண் தேவை.';
    if (!emisId) return 'EMIS ID தேவை.';

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return 'Google Spreadsheet இணைக்கப்படவில்லை.';

    let sheet = ss.getSheetByName('Report');
    if (!sheet) sheet = ss.insertSheet('Report');

    const headers = [
      'Date',
      'Student Name',
      'WhatsApp',
      'EMIS ID',
      'Completed Kurals',
      'Score',
      'Used Time (sec)'
    ];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    } else {
      const width = Math.max(sheet.getLastColumn(), headers.length);
      const current = sheet.getRange(1,1,1,width).getDisplayValues()[0];

      headers.forEach((header,index) => {
        if (String(current[index] || '').trim() !== header) {
          sheet.getRange(1,index+1).setValue(header);
        }
      });
    }

    sheet.appendRow([
      new Date(),
      name,
      wa,
      emisId,
      completed,
      Number(score || 0),
      Number(usedSeconds || 0)
    ]);

    return 'வெற்றிகரமாகப் பதிவு செய்யப்பட்டது!';
  } catch (e) {
    return 'பதிவு செய்வதில் பிழை: ' + e.message;
  }
}
