/**
 * Pendaftaran.gs - registrasi mandiri karyawan baru, dan panel admin
 * untuk menyetujui/menolak pendaftaran (hanya peran 'internal').
 */

var SHEET_PENDAFTARAN = 'Pendaftaran Baru';

function getOrBuatSheetPendaftaran_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_PENDAFTARAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PENDAFTARAN);
    sheet.appendRow(['Nama', 'Bagian Kerja', 'PIN', 'Tanggal Daftar', 'Status']);
  }
  return sheet;
}

function daftarKaryawanBaru(nama, bagianKerja, pin) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (eLock) {
    return { sukses: false, pesan: 'Server sibuk, coba lagi sebentar.' };
  }
  try {
    nama = String(nama || '').trim();
    bagianKerja = String(bagianKerja || '').trim();
    pin = String(pin || '').trim();

    if (!nama) throw new Error('Nama tidak boleh kosong.');
    if (DAFTAR_BAGIAN_KERJA.indexOf(bagianKerja) === -1) throw new Error('Bagian kerja tidak valid.');
    if (!/^[0-9]{4}$/.test(pin)) throw new Error('PIN harus 4 digit angka.');

    var sheetKaryawan = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
    var dataKaryawan = sheetKaryawan.getDataRange().getValues();
    for (var i = 1; i < dataKaryawan.length; i++) {
      if (String(dataKaryawan[i][0]).trim() === nama) {
        throw new Error('Nama "' + nama + '" sudah terdaftar. Gunakan nama lain (mis. tambahkan angka di belakang).');
      }
    }

    var sheetDaftar = getOrBuatSheetPendaftaran_();
    var dataDaftar = sheetDaftar.getDataRange().getValues();
    for (var j = 1; j < dataDaftar.length; j++) {
      if (String(dataDaftar[j][0]).trim() === nama) {
        throw new Error('Nama "' + nama + '" sudah pernah didaftarkan dan sedang menunggu persetujuan.');
      }
    }

    sheetDaftar.appendRow([
      nama,
      bagianKerja,
      pin,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      'Menunggu Persetujuan'
    ]);

    return { sukses: true, pesan: 'Pendaftaran berhasil dikirim. Tunggu persetujuan admin sebelum bisa login.' };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  } finally {
    lock.releaseLock();
  }
}


// ============ PANEL ADMIN: PERSETUJUAN PENDAFTARAN ============
// Hanya peran 'internal' (PT. Ray) yang boleh memproses persetujuan pendaftaran.

function getDaftarPendaftaranBaru(password) {
  var cekLogin = loginAdmin(password);
  if (!cekLogin.sukses) return cekLogin;
  if (cekLogin.peran !== 'internal') return { sukses: false, pesan: 'Fitur ini hanya untuk Admin PT. Ray Mitra Perkasa.' };

  var sheet = getOrBuatSheetPendaftaran_();
  var data = sheet.getDataRange().getValues();
  var hasil = [];
  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][4] || '').trim();
    if (status === 'Menunggu Persetujuan') {
      hasil.push({ baris: i + 1, nama: data[i][0], bagianKerja: data[i][1], pin: data[i][2], tanggalDaftar: data[i][3] });
    }
  }
  return { sukses: true, daftar: hasil };
}

function setujuiPendaftaran(password, baris) {
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
    var sheetDaftar = getOrBuatSheetPendaftaran_();
    var baris_ = parseInt(baris, 10);
    var data = sheetDaftar.getRange(baris_, 1, 1, 5).getValues()[0];
    var nama = String(data[0]).trim();
    var bagianKerja = String(data[1]).trim();
    var pin = String(data[2]).trim();
    var status = String(data[4]).trim();

    if (status !== 'Menunggu Persetujuan') {
      return { sukses: false, pesan: 'Pendaftaran ini sudah diproses sebelumnya.' };
    }

    var sheetKaryawan = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
    var dataKaryawan = sheetKaryawan.getDataRange().getValues();
    for (var i = 1; i < dataKaryawan.length; i++) {
      if (String(dataKaryawan[i][0]).trim() === nama) {
        return { sukses: false, pesan: 'Nama "' + nama + '" sudah ada di Data Master Karyawan. Tolak pendaftaran ini kalau memang duplikat.' };
      }
    }

    sheetKaryawan.appendRow([nama, '', '', bagianKerja, pin]);
    sheetDaftar.getRange(baris_, 5).setValue('Disetujui');

    return { sukses: true, pesan: nama + ' berhasil disetujui dan sekarang bisa login.' };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  } finally {
    lock.releaseLock();
  }
}

function tolakPendaftaran(password, baris) {
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
    var sheetDaftar = getOrBuatSheetPendaftaran_();
    var baris_ = parseInt(baris, 10);
    var status = String(sheetDaftar.getRange(baris_, 5).getValue()).trim();
    if (status !== 'Menunggu Persetujuan') {
      return { sukses: false, pesan: 'Pendaftaran ini sudah diproses.' };
    }
    sheetDaftar.getRange(baris_, 5).setValue('Ditolak');
    return { sukses: true, pesan: 'Pendaftaran ditolak.' };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  } finally {
    lock.releaseLock();
  }
}
