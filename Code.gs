/*** KirimDigi RSVP - Salman & Pia ***/
/*** Paste di Extensions > Apps Script, lalu Deploy > New deployment > Web app ***/
var SHEET_ID = '1hPD_1-mvQ6DoYZmSOudAtl6UDq7vtiWVgnHxMJ5ItC8';
var SHEET_NAME = 'Sheet1';

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'list';
  if (action === 'add') {
    return addUcapan(e.parameter);
  }
  return listUcapan();
}

function doPost(e) {
  var p = {};
  try {
    if (e && e.postData && e.postData.contents) {
      p = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    p = (e && e.parameter) || {};
  }
  if (!p || Object.keys(p).length === 0) {
    p = (e && e.parameter) || {};
  }
  return addUcapan(p);
}

function listUcapan() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var values = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    // 1.timestamp 2.nama tamu 3.ucapan 4.konfirmasi kehadiran 5.jumlah tamu
    out.push({
      timestamp: values[i][0],
      nama: values[i][1],
      ucapan: values[i][2],
      konfirmasi: values[i][3],
      jumlah: values[i][4]
    });
  }
  out.reverse(); // terbaru dulu
  return jsonOut({ status: 'ok', data: out });
}

function addUcapan(p) {
  var nama = String((p.nama || p.name || '')).trim();
  var ucapan = String((p.ucapan || p.message || '')).trim();
  var konfirmasi = String((p.konfirmasi || p.attendance || '')).trim();
  var jumlah = parseInt(p.jumlah || p.guestCount || 1, 10) || 1;
  if (!nama || !ucapan) {
    return jsonOut({ status: 'error', message: 'Nama dan ucapan wajib diisi.' });
  }
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  sh.appendRow([new Date(), nama, ucapan, konfirmasi, jumlah]);
  return jsonOut({ status: 'ok', message: 'Terima kasih atas ucapannya!' });
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
