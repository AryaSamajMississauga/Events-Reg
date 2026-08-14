/**
 * Geet Sangeet Registration — Google Sheets backend
 *
 * Writes/reads by HEADER NAME (not fixed column position), so it keeps
 * working even if the sheet's columns are reordered, and it auto-adds any
 * missing columns. All actions use GET because Apps Script's /exec URL
 * redirects and browsers drop POST bodies on that redirect.
 *
 * SETUP / UPDATE:
 * 1. Google Sheet -> Extensions -> Apps Script.
 * 2. Replace all code with this file, Save.
 * 3. Deploy -> Manage deployments -> (pencil) -> Version: New version -> Deploy.
 *    Execute as: Me   |   Who has access: Anyone
 *
 * NOTE: If you already have a "Registrations" tab from an earlier version
 * that used a "DOB" column, this code will simply add a "Phone" column at
 * the end and leave DOB empty — no data is lost. If you'd rather start
 * clean, delete the "Registrations" tab once and it will be recreated with
 * the correct headers.
 */

const SHEET_NAME = 'Registrations';
const HEADERS = ['Code','PrimaryName','Email','Phone','AttendeesJSON',
                 'Total','Paid','PaidAt','CreatedAt','CheckedIn','CheckedInAt'];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    return sheet;
  }
  // Ensure every expected header exists; append any that are missing.
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  let existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = HEADERS.filter(h => existing.indexOf(h) === -1);
  if (missing.length) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
  return sheet;
}

function headerMap(sheet) {
  const lastCol = sheet.getLastColumn();
  const row = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const map = {};
  row.forEach((h, i) => { map[h] = i; }); // 0-based
  return map;
}

function doGet(e)  { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  const action = e.parameter.action;
  const sheet = getSheet();
  const map = headerMap(sheet);

  if (action === 'getAll') {
    return jsonOut(readAllRows(sheet, map));
  }

  if (action === 'getByCode') {
    const code = (e.parameter.code || '').trim().toUpperCase();
    const rows = readAllRows(sheet, map);
    const found = rows.find(r => (r.code || '').toUpperCase() === code);
    return jsonOut(found || {});
  }

  if (action === 'create') {
    let record;
    try { record = JSON.parse(e.parameter.data); }
    catch (err) { return jsonOut({ success: false, error: 'bad data' }); }

    const width = sheet.getLastColumn();
    const row = new Array(width).fill('');
    const put = (h, v) => { if (map[h] !== undefined) row[map[h]] = v; };
    put('Code', record.code);
    put('PrimaryName', record.primaryName);
    put('Email', record.email);
    put('Phone', record.phone || '');
    put('AttendeesJSON', JSON.stringify(record.attendees || []));
    put('Total', record.total);
    put('Paid', false);
    put('PaidAt', '');
    put('CreatedAt', new Date().toISOString());
    put('CheckedIn', false);
    put('CheckedInAt', '');
    sheet.appendRow(row);
    return jsonOut({ success: true });
  }

  if (action === 'markPaid' || action === 'checkin') {
    const data = sheet.getDataRange().getValues();
    const codeCol = map['Code'];
    const target = (e.parameter.code || '').trim().toUpperCase();

    for (let i = 1; i < data.length; i++) {
      if ((data[i][codeCol] || '').toString().toUpperCase() === target) {
        const rowNum = i + 1;
        if (action === 'markPaid') {
          sheet.getRange(rowNum, map['Paid'] + 1).setValue(true);
          sheet.getRange(rowNum, map['PaidAt'] + 1).setValue(new Date().toISOString());
        } else {
          const current = data[i][map['CheckedIn']] === true;
          const newVal = !current;
          sheet.getRange(rowNum, map['CheckedIn'] + 1).setValue(newVal);
          sheet.getRange(rowNum, map['CheckedInAt'] + 1).setValue(newVal ? new Date().toISOString() : '');
        }
        const refreshed = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
        return jsonOut(rowToObj(map, refreshed));
      }
    }
    return jsonOut({ success: false, error: 'not found' });
  }

  return jsonOut({ error: 'unknown action' });
}

function readAllRows(sheet, map) {
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][map['Code']]) continue;
    rows.push(rowToObj(map, data[i]));
  }
  return rows;
}

function rowToObj(map, row) {
  const g = (h) => (map[h] !== undefined ? row[map[h]] : '');
  let attendees = [];
  try { attendees = JSON.parse(g('AttendeesJSON') || '[]'); } catch (err) { attendees = []; }
  return {
    code: g('Code'),
    primaryName: g('PrimaryName'),
    email: g('Email'),
    phone: g('Phone'),
    attendees: attendees,
    total: g('Total'),
    paid: g('Paid') === true,
    paidAt: g('PaidAt'),
    createdAt: g('CreatedAt'),
    checkedIn: g('CheckedIn') === true,
    checkedInAt: g('CheckedInAt')
  };
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
