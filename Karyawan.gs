/**
 * Karyawan.gs - data master karyawan & lokasi kerja, login karyawan
 * (nama + PIN), dan status absen hari ini.
 */

// ============ DATA MASTER: KARYAWAN & LOKASI ============

function getDaftarKaryawanPublik() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][0] || '').trim();
    if (nama) out.push(nama);
  }
  return out;
}

function cariKaryawanByNamaPin_(nama, pin) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  var namaTarget = String(nama || '').trim();
  var pinTarget = String(pin || '').trim();

  for (var i = 1; i < data.length; i++) {
    var namaSheet = String(data[i][0] || '').trim();
    var pinSheet = String(data[i][4] || '').trim();
    if (namaSheet && namaSheet === namaTarget) {
      if (pinSheet && pinSheet === pinTarget) {
        return { nama: namaSheet, bagianKerja: String(data[i][3] || '').trim() };
      }
      return null;
    }
  }
  return null;
}

// Mengambil daftar lokasi kerja yang valid saja (melindungi dari data lat/lng yang salah format)
// Juga dipakai frontend untuk menampilkan peta live & hitung jarak sisi client.
function getDaftarLokasiValid_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOKASI);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][0] || '').trim();
    var lat = parseFloat(String(data[i][1]).replace(',', '.'));
    var lng = parseFloat(String(data[i][2]).replace(',', '.'));
    var radius = parseFloat(data[i][3]) || 50;
    var satuan = String(data[i][4] || 'Meter').toLowerCase();

    if (!nama || isNaN(lat) || isNaN(lng)) continue;
    if (Math.abs(lat) > 90 || Math.abs(lng) > 180) continue;

    if (satuan.indexOf('km') !== -1) radius = radius * 1000;

    out.push({ nama: nama, lat: lat, lng: lng, radius: radius });
  }
  return out;
}


// ============ STATUS ABSEN HARI INI ============

function getStatusAbsenHariIni_(nama) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ABSENSI);
  var data = sheet.getDataRange().getValues();
  var tglHariIni = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var status = { sudahClockIn: false, sudahClockOut: false, jamMasuk: '', jamPulang: '' };

  for (var i = 1; i < data.length; i++) {
    var tglBaris = data[i][2];
    var tglBarisStr = (tglBaris instanceof Date)
      ? Utilities.formatDate(tglBaris, Session.getScriptTimeZone(), 'dd/MM/yyyy')
      : String(tglBaris).trim();

    if (String(data[i][0]).trim() === nama && tglBarisStr === tglHariIni) {
      var jamBaris = data[i][3];
      var jamStr = (jamBaris instanceof Date)
        ? Utilities.formatDate(jamBaris, Session.getScriptTimeZone(), 'HH:mm')
        : String(jamBaris).trim().substring(0, 5);

      if (data[i][1] === 'CLOCK IN') { status.sudahClockIn = true; status.jamMasuk = jamStr; }
      if (data[i][1] === 'CLOCK OUT') { status.sudahClockOut = true; status.jamPulang = jamStr; }
    }
  }
  return status;
}


// ============ LOGIN (NAMA + PIN) ============

function loginKaryawan(nama, pin) {
  var karyawan = cariKaryawanByNamaPin_(nama, pin);
  if (!karyawan) {
    return { sukses: false, pesan: 'Nama atau PIN salah.' };
  }
  return {
    sukses: true,
    bagianKerja: karyawan.bagianKerja,
    status: getStatusAbsenHariIni_(karyawan.nama)
  };
}

// Dipakai untuk refresh status setelah absen, tetap re-verifikasi PIN
function getStatusUlang(nama, pin) {
  var karyawan = cariKaryawanByNamaPin_(nama, pin);
  if (!karyawan) return { sukses: false, pesan: 'Sesi tidak valid, silakan masuk ulang.' };
  return { sukses: true, status: getStatusAbsenHariIni_(karyawan.nama) };
}
