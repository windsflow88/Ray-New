/**
 * Cleaning.gs - modul Laporan Kegiatan (Cleaning Service).
 *
 * v6.2: sheet "Master Tugas Kebersihan" (daftar tugas, diisi manual admin) dan
 * "Data Laporan Kegiatan" (log laporan tugas selesai, dengan foto sebelum & sesudah).
 * v6.3: setiap tugas sekarang punya Latitude/Longitude/Radius/Kode QR sendiri
 * (sama seperti Titik Patroli) - lapor kegiatan sekarang WAJIB scan barcode di
 * lokasi tugas & divalidasi geofencing seperti Patroli.
 */

var SHEET_MASTER_TUGAS = 'Master Tugas Kebersihan';
var SHEET_LAPORAN_KEGIATAN = 'Data Laporan Kegiatan';

function getOrBuatSheetMasterTugas_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MASTER_TUGAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MASTER_TUGAS);
    sheet.appendRow(['Nama Tugas', 'Area/Lokasi', 'Latitude', 'Longitude', 'Radius', 'Satuan Radius', 'Kode QR', 'Keterangan']);
  }
  return sheet;
}

function getOrBuatSheetLaporanKegiatan_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_LAPORAN_KEGIATAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LAPORAN_KEGIATAN);
    sheet.appendRow(['Nama', 'Tugas', 'Area', 'Tanggal', 'Waktu', 'Latitude', 'Longitude', 'Akurasi', 'Foto Sebelum', 'Foto Sesudah', 'Catatan']);
  }
  return sheet;
}

// v6.3: kolom Latitude(2)/Longitude(3)/Radius(4)/Satuan(5)/Kode QR(6) opsional - kalau lat/lng kosong,
// tugas itu dianggap tidak butuh validasi lokasi (dilewati saat cek geofencing di simpanLaporanKegiatan).
function getDaftarTugasKebersihan() {
  var sheet = getOrBuatSheetMasterTugas_();
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var namaTugas = String(data[i][0] || '').trim();
    var area = String(data[i][1] || '').trim();
    if (!namaTugas) continue;

    var lat = parseFloat(String(data[i][2]).replace(',', '.'));
    var lng = parseFloat(String(data[i][3]).replace(',', '.'));
    var radius = parseFloat(data[i][4]) || 50;
    var satuan = String(data[i][5] || 'Meter').toLowerCase();
    var kodeQr = String(data[i][6] || '').trim();
    if (satuan.indexOf('km') !== -1) radius = radius * 1000;

    out.push({
      namaTugas: namaTugas,
      area: area,
      lat: isNaN(lat) ? null : lat,
      lng: isNaN(lng) ? null : lng,
      radius: radius,
      kodeQr: kodeQr
    });
  }
  return out;
}

function simpanLaporanKegiatan(nama, pin, tugas, area, lat, lng, akurasi, fotoSebelumBase64, fotoSesudahBase64, mimeType, catatan) {
  var karyawan = cariKaryawanByNamaPin_(nama, pin);
  if (!karyawan) return { sukses: false, pesan: 'Nama atau PIN tidak valid. Silakan masuk ulang.' };

  if (karyawan.bagianKerja !== 'Cleaning Service') {
    return { sukses: false, pesan: 'Fitur Laporan Kegiatan hanya untuk staff Cleaning Service.' };
  }

  tugas = String(tugas || '').trim();
  if (!tugas) return { sukses: false, pesan: 'Tugas harus dipilih/discan dulu.' };
  if (!fotoSebelumBase64 || !fotoSesudahBase64) {
    return { sukses: false, pesan: 'Foto sebelum dan sesudah wajib dilampirkan.' };
  }

  var tugasList = getDaftarTugasKebersihan();
  var tugasObj = null;
  for (var i = 0; i < tugasList.length; i++) {
    if (tugasList[i].namaTugas === tugas) { tugasObj = tugasList[i]; break; }
  }
  if (!tugasObj) return { sukses: false, pesan: 'Tugas tidak dikenal.' };
  if (!area) area = tugasObj.area;

  lat = parseFloat(lat);
  lng = parseFloat(lng);

  // Geofencing cuma ditegakkan kalau tugas ini punya koordinat tersimpan (kolom Latitude/Longitude diisi).
  if (tugasObj.lat !== null && tugasObj.lng !== null) {
    if (isNaN(lat) || isNaN(lng)) return { sukses: false, pesan: 'Lokasi GPS tidak valid.' };

    var jarak = hitungJarakMeter_(lat, lng, tugasObj.lat, tugasObj.lng);
    if (jarak > tugasObj.radius) {
      return {
        sukses: false,
        pesan: 'Anda berada ' + Math.round(jarak) + ' meter dari lokasi tugas "' + tugasObj.namaTugas +
          '" (maksimal ' + tugasObj.radius + ' meter). Laporan ditolak.'
      };
    }
  }

  // ---- Upload foto (proses lambat), DI LUAR lock ----
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('FOLDER_ID');
  var urlSebelum = 'Folder Drive belum diatur';
  var urlSesudah = 'Folder Drive belum diatur';

  if (folderId) {
    try {
      var folder = DriveApp.getFolderById(folderId);
      var now = new Date();
      var stempel = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');

      var blobSebelum = Utilities.newBlob(Utilities.base64Decode(fotoSebelumBase64), mimeType || 'image/jpeg',
        'KEGIATAN_SEBELUM_' + karyawan.nama + '_' + stempel + '.jpg');
      var fileSebelum = folder.createFile(blobSebelum);
      urlSebelum = fileSebelum.getUrl();

      var blobSesudah = Utilities.newBlob(Utilities.base64Decode(fotoSesudahBase64), mimeType || 'image/jpeg',
        'KEGIATAN_SESUDAH_' + karyawan.nama + '_' + stempel + '.jpg');
      var fileSesudah = folder.createFile(blobSesudah);
      urlSesudah = fileSesudah.getUrl();

      try {
        fileSebelum.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        fileSesudah.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (eShare) {
        Logger.log('Gagal set share publik: ' + eShare.message);
      }
    } catch (eUp) {
      urlSebelum = 'Gagal upload: ' + eUp.message;
      urlSesudah = 'Gagal upload: ' + eUp.message;
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
    var sheet = getOrBuatSheetLaporanKegiatan_();
    sheet.appendRow([
      karyawan.nama,
      tugas,
      String(area || '').trim(),
      Utilities.formatDate(waktuSekarang, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      Utilities.formatDate(waktuSekarang, Session.getScriptTimeZone(), 'HH:mm:ss'),
      isNaN(lat) ? '' : lat,
      isNaN(lng) ? '' : lng,
      (akurasi || '-') + ' meter',
      urlSebelum,
      urlSesudah,
      String(catatan || '').trim()
    ]);

    return { sukses: true, pesan: 'Laporan kegiatan "' + tugas + '" berhasil dikirim.' };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  } finally {
    lock.releaseLock();
  }
}


// ============ DASHBOARD KPI CLEANING (ADMIN) ============
// Dipakai oleh KEDUA tingkat admin (internal & client).

function hitungKPICleaning_(bulan, tahun) {
  var sheetKegiatan = getOrBuatSheetLaporanKegiatan_();
  var dataKegiatan = sheetKegiatan.getDataRange().getValues();
  // kolom Tanggal ada di index 3: Nama(0), Tugas(1), Area(2), Tanggal(3), ...
  return hitungKPIOperasional_(bulan, tahun, 'Cleaning Service', dataKegiatan, 3);
}

function ambilDashboardCleaning(password, bulan, tahun) {
  var cekLogin = loginAdmin(password);
  if (!cekLogin.sukses) return cekLogin;
  try {
    return { sukses: true, kpi: hitungKPICleaning_(bulan, tahun) };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  }
}
