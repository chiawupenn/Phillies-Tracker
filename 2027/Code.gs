/**
 * Data feed for the 2027 tracker page (2027/index.html).
 *
 * Paste this into Extensions > Apps Script of the "Phillies Season Ticket 2027"
 * Google Sheet, then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Setup steps are in README.md.
 *
 * The feed only publishes date, time, opponent, special event and whether the
 * game is gone. Buyer names and prices stay in the sheet.
 */
const SHEET_NAME = '2027 Games';
const COLUMNS = {
  date: 'Game Date',
  time: 'Game Time',
  subject: 'SUBJECT',
  event: 'Special Event',
  sold: 'Ticket Forwarded',
};
const CHECKBOX_COLUMNS = ['Ticket Forwarded', 'Ticket Platform Sale', 'Self Sell', 'Attended Game'];

function doGet() {
  let body;
  try {
    body = readGames_();
  } catch (e) {
    body = { error: e.message };
  }
  return ContentService.createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

function readGames_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getGamesSheet_(ss);
  const tz = ss.getSpreadsheetTimeZone();
  const range = sheet.getDataRange();
  const values = range.getValues();
  const shown = range.getDisplayValues();
  const col = findColumns_(values[0], COLUMNS);

  const games = [];
  for (let r = 1; r < values.length; r++) {
    const date = values[r][col.date];
    if (!(date instanceof Date)) continue; // blank rows and the Grand Total row
    const subject = String(values[r][col.subject]).trim();
    games.push({
      date: Utilities.formatDate(date, tz, 'yyyy-MM-dd'),
      // display value, so the time reads exactly as it does in the sheet ("6:40 PM")
      time: col.time >= 0 ? shown[r][col.time].trim() : '',
      opponent: subject.replace(/\s+at\s+Phillies\b.*$/i, '').trim(),
      event: col.event >= 0 ? String(values[r][col.event]).trim() : '',
      sold: values[r][col.sold] === true,
    });
  }
  return games;
}

/**
 * Run once from the editor: turns the TRUE/FALSE columns into checkboxes.
 * (Same as selecting those cells and choosing Insert > Checkbox.)
 */
function setupCheckboxes() {
  const sheet = getGamesSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const values = sheet.getDataRange().getValues();
  const header = values[0].map(h => String(h).trim());
  const dateCol = findColumns_(values[0], COLUMNS).date;
  let lastGameRow = 1;
  values.forEach((row, i) => {
    if (row[dateCol] instanceof Date) lastGameRow = i + 1;
  });
  if (lastGameRow < 2) throw new Error('No game rows found');
  CHECKBOX_COLUMNS.forEach(name => {
    const c = header.indexOf(name);
    if (c < 0) throw new Error('Missing column: ' + name);
    sheet.getRange(2, c + 1, lastGameRow - 1, 1).insertCheckboxes();
  });
}

function getGamesSheet_(ss) {
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

function findColumns_(headerRow, wanted) {
  const header = headerRow.map(h => String(h).trim().toLowerCase());
  const col = {};
  Object.keys(wanted).forEach(key => {
    col[key] = header.indexOf(wanted[key].toLowerCase());
  });
  ['date', 'subject', 'sold'].forEach(key => {
    if (col[key] < 0) throw new Error('Missing column: ' + wanted[key]);
  });
  return col;
}
