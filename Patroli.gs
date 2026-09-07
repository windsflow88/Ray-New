/**
 * Patroli.gs - modul Patroli (Security).
 *
 * v6.2: sheet "Titik Patroli" (titik ronda, diisi manual admin, pola sama dengan
 * Data Master Lokasi) dan "Data Patroli" (log tiap kali satpam checkpoint).
 * Proses berat (upload foto) dilakukan DI LUAR LockService, sama seperti pola
 * v6.1 di simpanAbsen/ajukanIjin - hanya appendRow yang wajib atomik.
 *
 * v6.3: kolom "Kode QR" untuk identifikasi titik via scan barcode (dicocokkan
 * sisi frontend dari daftar yang sudah dikirim ke client, backend tetap yang
 * menegakkan validasi jarak GPS).
 */

var SHEET_TITIK_PATROLI = 'Titik Patroli';
var SHEET_DATA_PATROLI = 'Data Patroli';

function getOrBuatSheetTitikPatroli_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TITIK_PATROLI);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TITIK_PATROLI);
    sheet.appendRow(['Nama Pos', 'Latitude', 'Longitude', 'Radius', 'Satuan Radius', 'Kode QR']);
  }
  return sheet;
}

function getOrBuatSheetDataPatroli_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_DATA_PATROLI);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DATA_PATROLI);
    sheet.appendRow(['Nama', 'Nama Pos', 'Tanggal', 'Waktu', 'Foto', 'Latitude', 'Longitude', 'Akurasi', 'Catatan']);
  }
  return sheet;
}

// Daftar titik patroli yang valid saja (pola sama seperti getDaftarLokasiValid_)
// v6.3: sertakan kodeQr (kolom F "Kode QR") untuk dicocokkan sisi frontend saat scan barcode.
function getDaftarTitikPatroliValid_() {
  var sheet = getOrBuatSheetTitikPatroli_();
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][0] || '').trim();
    var lat = parseFloat(String(data[i][1]).replace(',', '.'));
    var lng = parseFloat(String(data[i][2]).replace(',', '.'));
    var radius = parseFloat(data[i][3]) || 50;
    var satuan = String(data[i][4] || 'Meter').toLowerCase();
    var kodeQr = String(data[i][5] || '').trim();

    if (!nama || isNaN(lat) || isNaN(lng)) continue;
    if (Math.abs(lat) > 90 || Math.abs(lng) > 180) continue;

    if (satuan.indexOf('km') !== -1) radius = radius * 1000;

    out.push({ nama: nama, lat: lat, lng: lng, radius: radius, kodeQr: kodeQr });
  }
  return out;
}

function getDaftarTitikPatroli() {
  return getDaftarTitikPatroliValid_();
}

function simpanPatroli(nama, pin, namaPos, fotoBase64, mimeType, lat, lng, akurasi, catatan) {
  var karyawan = cariKaryawanByNamaPin_(nama, pin);
  if (!karyawan) return { sukses: false, pesan: 'Nama atau PIN tidak valid. Silakan masuk ulang.' };

  if (karyawan.bagianKerja !== 'Security') {
    return { sukses: false, pesan: 'Fitur Patroli hanya untuk staff Security.' };
  }

  lat = parseFloat(lat);
  lng = parseFloat(lng);
  if (isNaN(lat) || isNaN(lng)) return { sukses: false, pesan: 'Lokasi GPS tidak valid.' };

  if (!fotoBase64) return { sukses: false, pesan: 'Foto bukti patroli wajib dilampirkan.' };

  var titikList = getDaftarTitikPatroliValid_();
  if (titikList.length === 0) {
    return { sukses: false, pesan: 'Belum ada titik patroli terdaftar. Hubungi admin untuk mengisi sheet "Titik Patroli".' };
  }

  var titik = null;
  for (var i = 0; i < titikList.length; i++) {
    if (titikList[i].nama === namaPos) { titik = titikList[i]; break; }
  }
  if (!titik) return { sukses: false, pesan: 'Titik pos tidak dikenal.' };

  var jarak = hitungJarakMeter_(lat, lng, titik.lat, titik.lng);
  if (jarak > titik.radius) {
    return {
      sukses: false,
      pesan: 'Anda berada ' + Math.round(jarak) + ' meter dari pos "' + titik.nama +
        '" (maksimal ' + titik.radius + ' meter). Lapor patroli ditolak.'
    };
  }

  // ---- Upload foto (proses lambat), DI LUAR lock ----
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('FOLDER_ID');
  var fileUrl = 'Folder Drive belum diatur';

  if (folderId) {
    try {
      var folder = DriveApp.getFolderById(folderId);
      var now = new Date();
      var namaFile = 'PATROLI_' + karyawan.nama + '_' + namaPos + '_' +
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.jpg';
      var blob = Utilities.newBlob(Utilities.base64Decode(fotoBase64), mimeType || 'image/jpeg', namaFile);
      var file = folder.createFile(blob);
      fileUrl = file.getUrl();
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (eShare) {
        Logger.log('Gagal set share publik: ' + eShare.message);
      }
    } catch (eUp) {
      fileUrl = 'Gagal upload: ' + eUp.message;
    }
  }

  // ---- Bagian kritis (wajib atomik): tulis baris baru ----
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (eLock) {
    return { sukses: false, pesan: 'Server sibuk, coba lagi sebentar.' };
  }
  try {
    var waktuSekarang = new Date();
    var sheet = getOrBuatSheetDataPatroli_();
    sheet.appendRow([
      karyawan.nama,
      namaPos,
      Utilities.formatDate(waktuSekarang, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      Utilities.formatDate(waktuSekarang, Session.getScriptTimeZone(), 'HH:mm:ss'),
      fileUrl,
      lat,
      lng,
      (akurasi || '-') + ' meter',
      String(catatan || '').trim()
    ]);

    return { sukses: true, pesan: 'Patroli di pos "' + namaPos + '" berhasil dicatat.' };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  } finally {
    lock.releaseLock();
  }
}


// ============ DASHBOARD KPI PATROLI (ADMIN) ============
// Dipakai oleh KEDUA tingkat admin (internal & client).

function hitungKPIPatroli_(bulan, tahun) {
  var sheetPatroli = getOrBuatSheetDataPatroli_();
  var dataPatroli = sheetPatroli.getDataRange().getValues();
  // kolom Tanggal ada di index 2: Nama(0), Nama Pos(1), Tanggal(2), ...
  return hitungKPIOperasional_(bulan, tahun, 'Security', dataPatroli, 2);
}

function ambilDashboardPatroli(password, bulan, tahun) {
  var cekLogin = loginAdmin(password);
  if (!cekLogin.sukses) return cekLogin;
  try {
    return { sukses: true, kpi: hitungKPIPatroli_(bulan, tahun) };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  }
}
