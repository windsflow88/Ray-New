/**
 * Ijin.gs - form ijin/cuti karyawan, dan panel admin untuk
 * menyetujui/menolak pengajuan (hanya peran 'internal').
 */

var SHEET_IJIN = 'Pengajuan Ijin';

function getOrBuatSheetIjin_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_IJIN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_IJIN);
    sheet.appendRow(['Nama', 'Jenis Ijin', 'Tanggal Ijin', 'Keterangan', 'Link Bukti', 'Tanggal Pengajuan', 'Status']);
  }
  return sheet;
}

// v6.1: upload bukti foto/surat ke Drive dipindah KELUAR lock (proses lambat),
// cuma appendRow yang tetap di dalam lock.
function ajukanIjin(nama, pin, jenisIjin, tanggalIjin, keterangan, buktiBase64, mimeType) {
  var karyawan = cariKaryawanByNamaPin_(nama, pin);
  if (!karyawan) return { sukses: false, pesan: 'Sesi tidak valid, silakan masuk ulang.' };

  jenisIjin = String(jenisIjin || '').trim();
  tanggalIjin = String(tanggalIjin || '').trim();
  keterangan = String(keterangan || '').trim();

  if (!jenisIjin) return { sukses: false, pesan: 'Jenis ijin harus dipilih.' };
  if (!tanggalIjin) return { sukses: false, pesan: 'Tanggal ijin harus diisi.' };
  if (!buktiBase64) return { sukses: false, pesan: 'Foto/surat bukti wajib dilampirkan.' };

  // ---- Upload foto (proses lambat), DI LUAR lock ----
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('FOLDER_ID');
  var linkBukti = 'Folder Drive belum diatur';

  if (folderId) {
    try {
      var folder = DriveApp.getFolderById(folderId);
      var now = new Date();
      var namaFile = 'IJIN_' + karyawan.nama + '_' + jenisIjin + '_' +
        Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.jpg';
      var blob = Utilities.newBlob(Utilities.base64Decode(buktiBase64), mimeType || 'image/jpeg', namaFile);
      var file = folder.createFile(blob);
      linkBukti = file.getUrl();
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (eShare) {
        Logger.log('Gagal set share publik: ' + eShare.message);
      }
    } catch (eUp) {
      linkBukti = 'Gagal upload: ' + eUp.message;
    }
  }

  // ---- Bagian kritis (wajib atomik): tulis baris baru, sesingkat mungkin di dalam lock ----
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (eLock) {
    return { sukses: false, pesan: 'Server sibuk, coba lagi sebentar.' };
  }
  try {
    var sheet = getOrBuatSheetIjin_();
    sheet.appendRow([
      karyawan.nama,
      jenisIjin,
      tanggalIjin,
      keterangan,
      linkBukti,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      'Menunggu Persetujuan'
    ]);

    return { sukses: true, pesan: 'Pengajuan ijin berhasil dikirim. Menunggu persetujuan admin.' };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  } finally {
    lock.releaseLock();
  }
}

function getDaftarIjinAdmin(password) {
  var cekLogin = loginAdmin(password);
  if (!cekLogin.sukses) return cekLogin;
  if (cekLogin.peran !== 'internal') return { sukses: false, pesan: 'Fitur ini hanya untuk Admin PT. Ray Mitra Perkasa.' };

  var sheet = getOrBuatSheetIjin_();
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][6] || '').trim();
    if (status === 'Menunggu Persetujuan') {
      hasil.push({
        baris: i + 1,
        nama: data[i][0],
        jenisIjin: data[i][1],
        tanggalIjin: data[i][2],
        keterangan: data[i][3],
        linkBukti: data[i][4],
        tanggalPengajuan: data[i][5]
      });
    }
  }
  return { sukses: true, daftar: hasil };
}

function setujuiIjin(password, baris) {
  var cekLogin = loginAdmin(password);
  if (!cekLogin.sukses) return cekLogin;
  if (cekLogin.peran !== 'internal') return { sukses: false, pesan: 'Fitur ini hanya untuk Admin PT. Ray Mitra Perkasa.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (eLock) {
    return { sukses: false, pesan: 'Server sibuk, coba lagi sebentar.' };
  }
  try {
    var sheet = getOrBuatSheetIjin_();
    var baris_ = parseInt(baris, 10);
    var status = String(sheet.getRange(baris_, 7).getValue()).trim();
    if (status !== 'Menunggu Persetujuan') {
      return { sukses: false, pesan: 'Pengajuan ini sudah diproses.' };
    }
    sheet.getRange(baris_, 7).setValue('Disetujui');
    return { sukses: true, pesan: 'Pengajuan ijin disetujui.' };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  } finally {
    lock.releaseLock();
  }
}

function tolakIjin(password, baris) {
  var cekLogin = loginAdmin(password);
  if (!cekLogin.sukses) return cekLogin;
  if (cekLogin.peran !== 'internal') return { sukses: false, pesan: 'Fitur ini hanya untuk Admin PT. Ray Mitra Perkasa.' };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (eLock) {
    return { sukses: false, pesan: 'Server sibuk, coba lagi sebentar.' };
  }
  try {
    var sheet = getOrBuatSheetIjin_();
    var baris_ = parseInt(baris, 10);
    var status = String(sheet.getRange(baris_, 7).getValue()).trim();
    if (status !== 'Menunggu Persetujuan') {
      return { sukses: false, pesan: 'Pengajuan ini sudah diproses.' };
    }
    sheet.getRange(baris_, 7).setValue('Ditolak');
    return { sukses: true, pesan: 'Pengajuan ijin ditolak.' };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  } finally {
    lock.releaseLock();
  }
}
