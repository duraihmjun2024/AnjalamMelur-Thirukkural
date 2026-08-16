/***************************************************************
 * AMelur-Thirukkural
 * Code.gs — REST API BACKEND FOR GITHUB PAGES & APPS SCRIPT
 ***************************************************************/

const KURAL_SOURCE_SPREADSHEET_ID = '13JdpPtzsZOZ4s3s9Hfih6lN0c5-0VJTdwBTLBtT0y4k';
const KURAL_SOURCE_START_ROW = 2;
const KURAL_SOURCE_NUM_COLUMN = 3;
const KURAL_SOURCE_TEXT_COLUMN = 4;
const KURAL_SOURCE_COUNT = 1330;
const KURAL_CACHE_SECONDS = 21600;

function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : null;
    
    if (action === 'searchAll') {
      const allData = loadAllSourceKurals_();
      return createJsonResponse({ status: 'success', data: allData });
    }

    if (e && e.parameter && e.parameter.numbers) {
      const numbersArray = e.parameter.numbers.split(',').map(Number);
      const data = getKuralsData(numbersArray);
      return createJsonResponse({ status: 'success', data: data });
    }

    // Default: Return initial 10 Kurals (1-10)
    const defaultNumbers = Array.from({ length: 10 }, (_, i) => i + 1);
    const initialData = getKuralsData(defaultNumbers);
    return createJsonResponse({ status: 'success', data: initialData });

  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

function doPost(e) {
  try {
    const contents = e.postData ? e.postData.contents : null;
    if (!contents) {
      return createJsonResponse({ status: 'error', message: 'No payload received' });
    }

    const payload = JSON.parse(contents);
    const resultMsg = logUserLearningStatus(
      payload.studentName,
      payload.whatsapp,
      payload.emis,
      payload.kuralNumbers,
      payload.startedAt,
      payload.usedSeconds,
      payload.score
    );

    return createJsonResponse({ status: 'success', message: resultMsg });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
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
    const cached = CacheService.getScriptCache().get('SHEET_REAL_KURAL_' + n);
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
    throw new Error('குறள் எண்கள் கிடைக்கவில்லை: ' + unavailable.join(', '));
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

    if (!Number.isInteger(number) || number < 1 || number > 1330 || !fullText) {
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
  if (!record || Number(record.num) !== Number(expectedNumber)) return false;
  if (typeof record.text !== 'string' || !record.text.trim()) return false;
  if (!Array.isArray(record.voiceParts) || !record.voiceParts.length) return false;
  return true;
}

function logUserLearningStatus(studentName, whatsapp, emis, kuralNumbers, startedAt, usedSeconds, score) {
  try {
    const name = String(studentName || '').trim();
    const wa = String(whatsapp || '').trim();
    const emisId = String(emis || '').trim();
    const completed = String(kuralNumbers || '').trim();

    if (!name || !wa || !emisId) return 'மாணவர் விவரங்கள் அரைகுறையாக உள்ளன.';

    const ss = SpreadsheetApp.openById(KURAL_SOURCE_SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Report');
    if (!sheet) sheet = ss.insertSheet('Report');

    const headers = ['Date', 'Student Name', 'WhatsApp', 'EMIS ID', 'Completed Kurals', 'Score', 'Used Time (sec)'];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
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
