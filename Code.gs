/**
 * Geet Sangeet Registration — Google Sheets backend
 *
 * Uses GET for every action (including writes) because Apps Script's
 * /exec URL issues a redirect, and browsers silently downgrade POST
 * requests to GET on that kind of redirect — which drops the request
 * body. GET requests survive the redirect intact, so all data is sent
 * as URL query parameters instead.
 *
 * SETUP:
 * 1. Open the REAL spreadsheet (from Google Drive — not the pubhtml link).
 * 2. Extensions -> Apps Script.
 * 3. Replace all code with this file's contents, save.
 * 4. Deploy -> Manage deployments -> Edit (pencil) -> New version -> Deploy.
 *    (Or Deploy -> New deployment if this is your first time.)
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web app URL (ends in /exec) into both HTML apps as API_URL.
 */

const SHEET_NAME = 'Registrations';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Code', 'PrimaryName', 'Email', 'DOB', 'AttendeesJSON',
      'Total', 'Paid', 'PaidAt', 'CreatedAt', 'CheckedIn', 'CheckedInAt'
    ]);
  }
  return sheet;
}

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  // Kept for completeness, but the front-end apps only use GET (see note above).
  return handleRequest(e);
}

function handleRequest(e) {
  const action = e.parameter.action;
  const sheet = getSheet();

  if (action === 'getByCode') {
    const code = (e.parameter.code || '').trim().toUpperCase();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1).filter(r => r[headers.indexOf('Code')]).map(r => rowToObj(headers, r));
    const found = rows.find(r => (r.code || '').toUpperCase() === code);
    return jsonOut(found || {});
  }

  if (action === 'getAll') {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1).filter(r => r[headers.indexOf('Code')]).map(r => rowToObj(headers, r));
    return jsonOut(rows);
  }

  if (action === 'create') {
    let record;
    try {
      record = JSON.parse(e.parameter.data);
    } catch (err) {
      return jsonOut({ success: false, error: 'bad data' });
    }
    sheet.appendRow([
      record.code,
      record.primaryName,
      record.email,
      record.dob,
      JSON.stringify(record.attendees || []),
      record.total,
      false,
      '',
      new Date().toISOString(),
      false,
      ''
    ]);
    return jsonOut({ success: true });
  }

  if (action === 'markPaid' || action === 'checkin') {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const codeCol = headers.indexOf('Code');
    const target = (e.parameter.code || '').trim().toUpperCase();

    for (let i = 1; i < data.length; i++) {
      if ((data[i][codeCol] || '').toString().toUpperCase() === target) {
        const rowNum = i + 1;

        if (action === 'markPaid') {
          sheet.getRange(rowNum, headers.indexOf('Paid') + 1).setValue(true);
          sheet.getRange(rowNum, headers.indexOf('PaidAt') + 1).setValue(new Date().toISOString());
        } else {
          const checkedInColIdx = headers.indexOf('CheckedIn');
          const current = data[i][checkedInColIdx] === true;
          const newVal = !current;
          sheet.getRange(rowNum, checkedInColIdx + 1).setValue(newVal);
          sheet.getRange(rowNum, headers.indexOf('CheckedInAt') + 1)
            .setValue(newVal ? new Date().toISOString() : '');
        }

        const refreshed = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
        return jsonOut(rowToObj(headers, refreshed));
      }
    }
    return jsonOut({ success: false, error: 'not found' });
  }

  return jsonOut({ error: 'unknown action' });
}

function rowToObj(headers, row) {
  const idx = (h) => headers.indexOf(h);
  let attendees = [];
  try {
    attendees = JSON.parse(row[idx('AttendeesJSON')] || '[]');
  } catch (err) {
    attendees = [];
  }
  return {
    code: row[idx('Code')],
    primaryName: row[idx('PrimaryName')],
    email: row[idx('Email')],
    dob: row[idx('DOB')],
    attendees: attendees,
    total: row[idx('Total')],
    paid: row[idx('Paid')] === true,
    paidAt: row[idx('PaidAt')],
    createdAt: row[idx('CreatedAt')],
    checkedIn: row[idx('CheckedIn')] === true,
    checkedInAt: row[idx('CheckedInAt')]
  };
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
