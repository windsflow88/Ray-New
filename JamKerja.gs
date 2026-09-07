/**
 * JamKerja.gs - sheet "Jam Kerja Standar" (per Bagian Kerja), dipakai
 * Dashboard.gs untuk menghitung keterlambatan.
 */

var SHEET_JAM_KERJA = 'Jam Kerja Standar';

function getOrBuatSheetJamKerja_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_JAM_KERJA);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_JAM_KERJA);
    sheet.appendRow(['Bagian Kerja', 'Jam Masuk Standar (HH:mm)', 'Toleransi (menit)']);
  }
  return sheet;
}

function getJamStandarMap_() {
  var sheet = getOrBuatSheetJamKerja_();
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var bagianKerja = String(data[i][0] || '').trim();
    var jamStr = formatJamSel_(data[i][1]);
    var toleransi = parseInt(data[i][2], 10);
    if (isNaN(toleransi)) toleransi = 0;
    if (!bagianKerja || !jamStr) continue;
    map[bagianKerja] = { jam: jamStr, menit: menitDariJam_(jamStr), toleransi: toleransi };
  }
  return map;
}
