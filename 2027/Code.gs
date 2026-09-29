/**
 * Data feed for the 2027 tracker page (2027/index.html).
 *
 * Paste this into Extensions > Apps Script of the "Phillies Season Ticket 2027"
 * Google Sheet, run setupSheet once, then Deploy > New deployment > Web app
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

// Layout of the "2027 Games" tab, used by setupSheet()
const SHEET_HEADERS = ['Game Date', 'Game Time', 'SUBJECT', 'Special Event', 'Ticket Forwarded', 'Buyer',
  'Forward to', 'Sale Price', 'Net Profit', 'Net Profit Per Seat', 'Ticket Platform Sale', 'Self Sell', 'Attended Game'];
const FORWARD_TO = ['StubHub', 'SeatGeek', 'Facebook Marketplace', 'Facebook Messenger', 'Chris & Jae', 'Self'];
const COLUMN_WIDTHS = [132, 100, 152, 146, 128, 226, 181, 103, 103, 131, 145, 89, 110, 26, 159, 103]; // A:P, pixels
const DATE_FORMAT = 'ddd", "m" / "d';   // Thu, 4 / 1 (same as 2026)
const TIME_FORMAT = 'h:mm am/pm';
const MONEY_FORMAT = '"$"#,##0.00';
const STRIKE_FORMULA = '=$E2=TRUE';

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
 * Run once from the editor after the sheet is uploaded. Applies the 2026 look
 * (date, time and currency formats, bold headers, blue inputs, frozen panes,
 * column widths, hidden Net Profit Per Seat), the Forward to dropdown, the
 * strike-through for forwarded games, and the checkboxes. Safe to run again.
 */
function setupSheet() {
  const sheet = getGamesSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const header = sheet.getRange(1, 1, 1, SHEET_HEADERS.length).getValues()[0].map(h => String(h).trim());
  if (header.join('|') !== SHEET_HEADERS.join('|')) {
    throw new Error('Row 1 is not the expected 2027 header row, so nothing was changed');
  }
  const last = lastGameRow_(sheet);
  const games = last - 1;
  const total = last + 1; // Grand Total row

  sheet.getRange(1, 1, total, 16).setFontFamily('Arial').setFontSize(10);
  sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1, games, 1).setNumberFormat(DATE_FORMAT).setHorizontalAlignment('right');
  sheet.getRange(2, 2, games, 1).setNumberFormat(TIME_FORMAT).setHorizontalAlignment('right');
  sheet.getRange(2, 8, games, 1).setHorizontalAlignment('center');
  sheet.getRange(2, 8, games + 1, 3).setNumberFormat(MONEY_FORMAT); // Sale Price to Per Seat, with totals
  sheet.getRange(total, 7, 1, 4).setFontWeight('bold');
  sheet.getRangeList(CHECKBOX_COLUMNS.map(name => {
    const c = header.indexOf(name) + 1;
    return sheet.getRange(2, c, games, 1).getA1Notation();
  })).setHorizontalAlignment('center');

  // Season Setup / Season Summary block
  sheet.getRange('O1:O15').setFontWeight('bold');
  sheet.getRange('P2').setNumberFormat(MONEY_FORMAT).setFontColor('#0000ff').setBackground('#ffff00')
    .setNote('Placeholder: the 2026 season ticket total ($13,472). Replace with the 2027 invoice amount; ' +
      'Per Game, Per Seat and Net Profit update from it.');
  sheet.getRangeList(['P4', 'P7']).setFontColor('#0000ff');
  sheet.getRange('P5:P6').setNumberFormat(MONEY_FORMAT);
  sheet.getRange('O17').setFontStyle('italic').setFontSize(9).setFontColor('#666666');

  sheet.getRange(2, 7, games, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(FORWARD_TO, true).setAllowInvalid(false).build());

  const rules = sheet.getConditionalFormatRules().filter(rule => !isStrikeRule_(rule));
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(STRIKE_FORMULA)
    .setStrikethrough(true).setRanges([sheet.getRange(2, 1, games, 3)]).build());
  sheet.setConditionalFormatRules(rules);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  COLUMN_WIDTHS.forEach((px, i) => sheet.setColumnWidth(i + 1, px));
  sheet.hideColumns(10);

  setupCheckboxes();
}

/**
 * Turns the TRUE/FALSE columns into checkboxes (setupSheet runs this too).
 * insertCheckboxes() sets every cell to FALSE, so the current ticks are saved
 * first and put back, which makes it safe to run after sales are entered.
 */
function setupCheckboxes() {
  const sheet = getGamesSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const values = sheet.getDataRange().getValues();
  const header = values[0].map(h => String(h).trim());
  const lastGameRow = lastGameRow_(sheet);
  CHECKBOX_COLUMNS.forEach(name => {
    const c = header.indexOf(name);
    if (c < 0) throw new Error('Missing column: ' + name);
    const range = sheet.getRange(2, c + 1, lastGameRow - 1, 1);
    const ticks = range.getValues().map(row => [row[0] === true]);
    range.insertCheckboxes();
    range.setValues(ticks);
  });
}

function getGamesSheet_(ss) {
  return ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
}

// Row number of the last game (the last row with a date in Game Date)
function lastGameRow_(sheet) {
  const values = sheet.getDataRange().getValues();
  const dateCol = findColumns_(values[0], COLUMNS).date;
  let last = 1;
  values.forEach((row, i) => {
    if (row[dateCol] instanceof Date) last = i + 1;
  });
  if (last < 2) throw new Error('No game rows found');
  return last;
}

function isStrikeRule_(rule) {
  const condition = rule.getBooleanCondition();
  return condition !== null &&
    condition.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
    condition.getCriteriaValues()[0] === STRIKE_FORMULA;
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
