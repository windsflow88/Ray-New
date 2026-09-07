/**
 * Absensi.gs - simpan absen (clock in/out) dan riwayat absensi karyawan.
 *
 * v6.1: validasi + proses berat (upload foto, reverse geocoding) dilakukan DI LUAR LockService.
 * Hanya cek-status-ulang + appendRow (cepat) yang dijalankan di dalam lock, supaya saat banyak
 * karyawan absen bersamaan (mis. jam masuk shift), antrean tidak macet menunggu upload/network.
 */

function simpanAbsen(nama, pin, jenisAbsen, fotoBase64, mimeType, lat, lng, akurasi) {
  var karyawan = cariKaryawanByNamaPin_(nama, pin);
  if (!karyawan) return { sukses: false, pesan: 'Nama atau PIN tidak valid. Silakan masuk ulang.' };

  lat = parseFloat(lat);
  lng = parseFloat(lng);
  if (isNaN(lat) || isNaN(lng)) return { sukses: false, pesan: 'Lokasi GPS tidak valid.' };

  var lokasiList = getDaftarLokasiValid_();
  if (lokasiList.length === 0) {
    return { sukses: false, pesan: 'Konfigurasi lokasi kerja bermasalah (data lat/lng di sheet "Data Master Lokasi" tidak valid). Hubungi admin.' };
  }

  var terdekat = null;
  var jarakTerdekat = Infinity;
  lokasiList.forEach(function (l) {
    var jarak = hitungJarakMeter_(lat, lng, l.lat, l.lng);
    if (jarak < jarakTerdekat) {
      jarakTerdekat = jarak;
      terdekat = l;
    }
  });

  if (jarakTerdekat > terdekat.radius) {
    return {
      sukses: false,
      pesan: 'Anda berada ' + Math.round(jarakTerdekat) + ' meter dari lokasi "' + terdekat.nama +
        '" (maksimal ' + terdekat.radius + ' meter). Absen ditolak.'
    };
  }

  // Cek awal (cepat, di luar lock) supaya tidak buang waktu upload foto kalau memang sudah tidak valid.
  // Cek final yang mengikat tetap dilakukan ulang di dalam lock di bawah.
  var statusAwal = getStatusAbsenHariIni_(karyawan.nama);
  if (jenisAbsen === 'CLOCK IN' && statusAwal.sudahClockIn) {
    return { sukses: false, pesan: 'Anda sudah CLOCK IN hari ini.' };
  }
  if (jenisAbsen === 'CLOCK OUT' && !statusAwal.sudahClockIn) {
    return { sukses: false, pesan: 'Anda belum CLOCK IN hari ini.' };
  }
  if (jenisAbsen === 'CLOCK OUT' && statusAwal.sudahClockOut) {
    return { sukses: false, pesan: 'Anda sudah CLOCK OUT hari ini.' };
  }

  // ---- Proses berat (upload foto + reverse geocoding), DI LUAR lock ----
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('FOLDER_ID');
  var fileUrl = 'Folder Drive belum diatur';

  if (folderId) {
    try {
      var folder = DriveApp.getFolderById(folderId);
      var now = new Date();
      var namaFile = karyawan.nama + '_' + jenisAbsen + '_' +
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
      // Catatan keamanan: pesan error sengaja TIDAK menyertakan folderId lagi (lihat SECURITY.md)
      // supaya ID folder Drive privat tidak bocor ke response publik saat terjadi error upload.
      fileUrl = 'Gagal upload: [' + eUp.name + '] ' + eUp.message;
    }
  }

  var alamat = reverseGeocode_(lat, lng);
  var linkMaps = 'https://www.google.com/maps?q=' + lat + ',' + lng;

  // ---- Bagian kritis (wajib atomik): cek ulang status + tulis baris baru ----
  // Sesingkat mungkin di dalam lock: cuma satu sheet-read + satu appendRow.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (eLock) {
    return { sukses: false, pesan: 'Server sibuk, coba lagi sebentar.' };
  }

  try {
    var statusUlang = getStatusAbsenHariIni_(karyawan.nama);
    if (jenisAbsen === 'CLOCK IN' && statusUlang.sudahClockIn) {
      throw new Error('Anda sudah CLOCK IN hari ini.');
    }
    if (jenisAbsen === 'CLOCK OUT' && !statusUlang.sudahClockIn) {
      throw new Error('Anda belum CLOCK IN hari ini.');
    }
    if (jenisAbsen === 'CLOCK OUT' && statusUlang.sudahClockOut) {
      throw new Error('Anda sudah CLOCK OUT hari ini.');
    }

    var waktuSekarang = new Date();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ABSENSI);
    sheet.appendRow([
      karyawan.nama,
      jenisAbsen,
      "'" + Utilities.formatDate(waktuSekarang, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      Utilities.formatDate(waktuSekarang, Session.getScriptTimeZone(), 'HH:mm:ss'),
      fileUrl,
      lat,
      lng,
      (akurasi || '-') + ' meter',
      alamat.kelurahan,
      alamat.kecamatan,
      alamat.kabupaten,
      linkMaps,
      karyawan.bagianKerja
    ]);

    return { sukses: true, pesan: 'Absen ' + jenisAbsen + ' berhasil dicatat untuk ' + karyawan.nama + '.' };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  } finally {
    lock.releaseLock();
  }
}


// ============ RIWAYAT ABSENSI (N hari terakhir) ============

function getRiwayatAbsensi(nama, pin, jumlahHari) {
  var karyawan = cariKaryawanByNamaPin_(nama, pin);
  if (!karyawan) return { sukses: false, pesan: 'Sesi tidak valid, silakan masuk ulang.' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ABSENSI);
  var data = sheet.getDataRange().getValues();
  var batasHari = jumlahHari || 7;
  var batasWaktu = new Date().getTime() - batasHari * 24 * 60 * 60 * 1000;

  var perTanggal = {};

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() !== karyawan.nama) continue;

    var tglBaris = data[i][2];
    var tglObj = (tglBaris instanceof Date) ? tglBaris : parseTanggalDDMMYYYY_(String(tglBaris).trim());
    if (!tglObj || tglObj.getTime() < batasWaktu) continue;

    var tglStr = Utilities.formatDate(tglObj, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var jamBaris = data[i][3];
    var jamStr = (jamBaris instanceof Date)
      ? Utilities.formatDate(jamBaris, Session.getScriptTimeZone(), 'HH:mm')
      : String(jamBaris).trim().substring(0, 5);

    if (!perTanggal[tglStr]) perTanggal[tglStr] = { tanggal: tglStr, urutan: tglObj.getTime(), jamMasuk: '', jamPulang: '' };
    if (data[i][1] === 'CLOCK IN') perTanggal[tglStr].jamMasuk = jamStr;
    if (data[i][1] === 'CLOCK OUT') perTanggal[tglStr].jamPulang = jamStr;
  }

  var hasil = [];
  for (var key in perTanggal) hasil.push(perTanggal[key]);
  hasil.sort(function (a, b) { return b.urutan - a.urutan; });
  hasil.forEach(function (r) { delete r.urutan; });

  return { sukses: true, riwayat: hasil };
}
