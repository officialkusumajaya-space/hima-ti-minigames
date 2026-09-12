// === HIMA TI Minigames - Google Apps Script Backend ===
// Cara setup:
// 1. Buka Google Sheet baru
// 2. Extensions > Apps Script
// 3. Hapus semua code, paste ini
// 4. Save (Ctrl+S)
// 5. Deploy > New deployment > Web app
//    - Execute as: Me
//    - Who has access: Anyone
// 6. Copy URL, kasih ke gw

const SHEET_NAME = 'scores';

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // Create sheet if missing
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['name', 'game', 'score', 'unit', 'time']);
  }
  
  const data = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd');
  
  // Filter today's scores
  const scores = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = Utilities.formatDate(new Date(row[4]), 'Asia/Makassar', 'yyyy-MM-dd');
    if (rowDate === today) {
      scores.push({
        name: row[0],
        game: row[1],
        score: row[2],
        unit: row[3],
        time: Utilities.formatDate(new Date(row[4]), 'Asia/Makassar', 'HH:mm')
      });
    }
  }
  
  // Sort: reaction ascending, others descending
  scores.sort((a, b) => {
    if (a.game === 'reaction') return a.score - b.score;
    return b.score - a.score;
  });
  
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, scores: scores.slice(0, 20) }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // Create sheet if missing
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['name', 'game', 'score', 'unit', 'time']);
  }
  
  try {
    const body = JSON.parse(e.postData.contents);
    const name = (body.name || 'anon').substring(0, 20);
    const game = body.game;
    const score = Number(body.score);
    const unit = body.unit || '';
    
    if (!game || isNaN(score)) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'invalid data' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    sheet.appendRow([name, game, score, unit, new Date()]);
    
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet_today() {
  // Debug helper - run this to check today's scores
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('No sheet'); return; }
  const data = sheet.getDataRange().getValues();
  Logger.log('Rows: ' + data.length);
  data.forEach(r => Logger.log(r));
}
