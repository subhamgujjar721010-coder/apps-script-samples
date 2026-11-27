const SHEET_NAME = 'Bookings';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'id',
      'patientName',
      'ageSex',
      'mobile',
      'centreCode',
      'tests',
      'status',
      'createdAt'
    ]);
  }
  return sheet;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) {
      return jsonResponse({ ok: true, bookings: [] });
    }

    const headers = rows[0];
    const requestedCentre = (e && e.parameter && e.parameter.centreCode)
      ? String(e.parameter.centreCode).trim()
      : '';

    const bookings = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[0] && !row[1]) continue;

      const item = {};
      for (let j = 0; j < headers.length; j++) {
        item[headers[j]] = row[j];
      }

      if (!requestedCentre || String(item.centreCode || '').trim() === requestedCentre) {
        bookings.push(item);
      }
    }

    return jsonResponse({ ok: true, bookings: bookings });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const sheet = getSheet();
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    const id = data.id || ('BK-' + Date.now());
    const patientName = data.patientName || '';
    const ageSex = data.ageSex || '';
    const mobile = data.mobile || '';
    const centreCode = data.centreCode || '';
    const tests = Array.isArray(data.tests) ? data.tests.join(', ') : (data.tests || '');
    const status = data.status || 'Booked';
    const createdAt = data.createdAt || new Date().toISOString();

    sheet.appendRow([id, patientName, ageSex, mobile, centreCode, tests, status, createdAt]);

    return jsonResponse({
      ok: true,
      id: id,
      booking: { id, patientName, ageSex, mobile, centreCode, tests, status, createdAt }
    });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}
