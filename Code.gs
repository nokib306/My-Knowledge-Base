// FINAL KNOWLEDGE BASE - CODE.GS
// Simple 5 column structure: Date, Time, Category, Content, ID

const SHEET_NAME = 'Knowledge Base';
const HEADER_ROW = ['Date', 'Time', 'Category', 'Content', 'ID'];
const CATEGORIES = ['Text', 'Code', 'Link', 'Idea', 'Photo', 'Other'];

function doGet() {
  ensureSheet();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Knowledge Base')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ensureSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME, 0);
  }
  const range = sheet.getRange(1, 1, 1, HEADER_ROW.length);
  const existing = range.getValues()[0];
  const needsHeaders = HEADER_ROW.some((header, index) => existing[index] !== header);
  if (needsHeaders) {
    range.setValues([HEADER_ROW]);
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet not found. Please open the spreadsheet to create the "Knowledge Base" sheet.');
  }
  return sheet;
}

function getEntries(query) {
  ensureSheet();
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values = sheet.getRange(2, 1, lastRow - 1, HEADER_ROW.length).getValues();
  const normalizedQuery = (query || '').toString().trim().toLowerCase();

  return values
    .map(row => ({
      date: row[0],
      time: row[1],
      category: row[2],
      content: row[3],
      id: row[4]
    }))
    .filter(entry => {
      if (!normalizedQuery) return true;
      return Object.values(entry)
        .join(' ')
        .toLowerCase()
        .includes(normalizedQuery);
    })
    .sort((a, b) => new Date(`${b.date} ${b.time}`) - new Date(`${a.date} ${a.time}`));
}

function saveEntry(entry) {
  ensureSheet();
  const sheet = getSheet();
  const timestamp = new Date();
  const date = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const time = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'HH:mm:ss');

  if (!entry || !entry.content) {
    throw new Error('Content is required');
  }
  if (!entry.category) {
    throw new Error('Category is required');
  }

  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  if (entry.id) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] === entry.id) {
        rowIndex = i + 1; // account for 1-based rows
        break;
      }
    }
  }

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, HEADER_ROW.length).setValues([[date, time, entry.category, entry.content, entry.id]]);
    return entry.id;
  }

  const newId = entry.id || Utilities.getUuid();
  sheet.appendRow([date, time, entry.category, entry.content, newId]);
  return newId;
}

function deleteEntry(id) {
  ensureSheet();
  if (!id) {
    throw new Error('Missing ID');
  }
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][4] === id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error('Entry not found');
}

function getCategories() {
  return CATEGORIES;
}
