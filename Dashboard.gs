/**
 * Dashboard.gs - Dashboard KPI Disiplin (absensi, telat, ijin) untuk admin,
 * dan fungsi generik hitungKPIOperasional_ yang dipakai bareng oleh
 * Patroli.gs & Cleaning.gs untuk menghitung KPI kerajinan lapor.
 *
 * Dipakai oleh KEDUA tingkat admin (internal & client) - tidak ada data
 * sensitif approval di sini, hanya statistik.
 */

function hitungKPI_(bulan, tahun) {
  var sekarang = new Date();
  bulan = parseInt(bulan, 10);
  tahun = parseInt(tahun, 10);
  if (isNaN(bulan) || bulan < 1 || bulan > 12) bulan = sekarang.getMonth() + 1;
  if (isNaN(tahun)) tahun = sekarang.getFullYear();

  var sheetKaryawan = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
  var dataKaryawan = sheetKaryawan ? sheetKaryawan.getDataRange().getValues() : [];
  var daftarKaryawan = [];
  for (var k = 1; k < dataKaryawan.length; k++) {
    var namaK = String(dataKaryawan[k][0] || '').trim();
    if (!namaK) continue;
    daftarKaryawan.push({ nama: namaK, bagianKerja: String(dataKaryawan[k][3] || '').trim() });
  }

  var jamStandarMap = getJamStandarMap_();

  var sheetAbsensi = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ABSENSI);
  var dataAbsensi = sheetAbsensi ? sheetAbsensi.getDataRange().getValues() : [];

  var sheetIjin = getOrBuatSheetIjin_();
  var dataIjin = sheetIjin.getDataRange().getValues();

  var statPerKaryawan = {};
  daftarKaryawan.forEach(function (kry) {
    statPerKaryawan[kry.nama] = {
      nama: kry.nama,
      bagianKerja: kry.bagianKerja,
      hadir: 0,
      telat: 0,
      ijin: 0,
      adaJamStandar: !!jamStandarMap[kry.bagianKerja]
    };
  });

  var tanggalHadirSet = {};

  for (var i = 1; i < dataAbsensi.length; i++) {
    var namaBaris = String(dataAbsensi[i][0] || '').trim();
    if (!statPerKaryawan[namaBaris]) continue;
    if (dataAbsensi[i][1] !== 'CLOCK IN') continue;

    var tglBaris = dataAbsensi[i][2];
    var tglObj = (tglBaris instanceof Date) ? tglBaris : parseTanggalDDMMYYYY_(String(tglBaris).trim());
    if (!tglObj) continue;
    if ((tglObj.getMonth() + 1) !== bulan || tglObj.getFullYear() !== tahun) continue;

    var tglKey = namaBaris + '|' + Utilities.formatDate(tglObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!tanggalHadirSet[tglKey]) {
      tanggalHadirSet[tglKey] = true;
      statPerKaryawan[namaBaris].hadir++;
    }

    var standar = jamStandarMap[statPerKaryawan[namaBaris].bagianKerja];
    if (standar) {
      var menitAktual = menitDariJam_(formatJamSel_(dataAbsensi[i][3]));
      if (menitAktual !== null && menitAktual > (standar.menit + standar.toleransi)) {
        statPerKaryawan[namaBaris].telat++;
      }
    }
  }

  for (var j = 1; j < dataIjin.length; j++) {
    var namaIjin = String(dataIjin[j][0] || '').trim();
    if (!statPerKaryawan[namaIjin]) continue;
    if (String(dataIjin[j][6] || '').trim() !== 'Disetujui') continue;

    var tglIjinObj = parseTanggalFleksibel_(dataIjin[j][2] instanceof Date
      ? Utilities.formatDate(dataIjin[j][2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(dataIjin[j][2]).trim());
    if (!tglIjinObj) continue;
    if ((tglIjinObj.getMonth() + 1) !== bulan || tglIjinObj.getFullYear() !== tahun) continue;

    statPerKaryawan[namaIjin].ijin++;
  }

  var akhirBulan = new Date(tahun, bulan, 0);
  var batasHari = akhirBulan;
  var iniBulanBerjalan = (tahun === sekarang.getFullYear() && bulan === (sekarang.getMonth() + 1));
  if (iniBulanBerjalan && sekarang < akhirBulan) batasHari = sekarang;

  var hariKerja = 0;
  for (var t = new Date(tahun, bulan - 1, 1); t <= batasHari; t.setDate(t.getDate() + 1)) {
    var hari = t.getDay();
    if (hari !== 0 && hari !== 6) hariKerja++;
  }
  if (hariKerja < 1) hariKerja = 1;

  var perKaryawan = [];
  var totalHadir = 0, totalTelat = 0, totalIjin = 0, totalPersen = 0;
  for (var nm in statPerKaryawan) {
    var s = statPerKaryawan[nm];
    s.persenKehadiran = Math.min(100, Math.round((s.hadir / hariKerja) * 100));
    perKaryawan.push(s);
    totalHadir += s.hadir;
    totalTelat += s.telat;
    totalIjin += s.ijin;
    totalPersen += s.persenKehadiran;
  }
  perKaryawan.sort(function (a, b) { return b.telat - a.telat || a.nama.localeCompare(b.nama); });

  var ringkasan = {
    bulan: bulan,
    tahun: tahun,
    hariKerja: hariKerja,
    totalKaryawan: perKaryawan.length,
    totalHadir: totalHadir,
    totalTelat: totalTelat,
    totalIjin: totalIjin,
    rataRataKehadiran: perKaryawan.length ? Math.round(totalPersen / perKaryawan.length) : 0
  };

  var namaBulanId = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  var grafik = [];
  for (var g = 5; g >= 0; g--) {
    var tglGeser = new Date(tahun, bulan - 1 - g, 1);
    var b = tglGeser.getMonth() + 1, y = tglGeser.getFullYear();
    var hadirBln = 0, telatBln = 0, ijinBln = 0;

    for (var ii = 1; ii < dataAbsensi.length; ii++) {
      var nmB = String(dataAbsensi[ii][0] || '').trim();
      if (!statPerKaryawan[nmB]) continue;
      if (dataAbsensi[ii][1] !== 'CLOCK IN') continue;
      var tglBObj = (dataAbsensi[ii][2] instanceof Date) ? dataAbsensi[ii][2] : parseTanggalDDMMYYYY_(String(dataAbsensi[ii][2]).trim());
      if (!tglBObj) continue;
      if ((tglBObj.getMonth() + 1) !== b || tglBObj.getFullYear() !== y) continue;
      hadirBln++;
      var standarB = jamStandarMap[statPerKaryawan[nmB].bagianKerja];
      if (standarB) {
        var menitB = menitDariJam_(formatJamSel_(dataAbsensi[ii][3]));
        if (menitB !== null && menitB > (standarB.menit + standarB.toleransi)) telatBln++;
      }
    }
    for (var jj = 1; jj < dataIjin.length; jj++) {
      var nmIj = String(dataIjin[jj][0] || '').trim();
      if (!statPerKaryawan[nmIj]) continue;
      if (String(dataIjin[jj][6] || '').trim() !== 'Disetujui') continue;
      var tglIjObj = parseTanggalFleksibel_(dataIjin[jj][2] instanceof Date
        ? Utilities.formatDate(dataIjin[jj][2], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(dataIjin[jj][2]).trim());
      if (!tglIjObj) continue;
      if ((tglIjObj.getMonth() + 1) !== b || tglIjObj.getFullYear() !== y) continue;
      ijinBln++;
    }

    grafik.push({ label: namaBulanId[b - 1] + ' ' + String(y).substring(2), hadir: hadirBln, telat: telatBln, ijin: ijinBln });
  }

  return { ringkasan: ringkasan, perKaryawan: perKaryawan, grafik: grafik };
}

function ambilDashboardKPI(password, bulan, tahun) {
  var cekLogin = loginAdmin(password);
  if (!cekLogin.sukses) return cekLogin;
  try {
    return { sukses: true, kpi: hitungKPI_(bulan, tahun) };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  }
}


// ============ KPI GENERIK OPERASIONAL (dipakai Patroli & Cleaning) ============
// Basis kerajinan: jumlah laporan masuk per bulan dibanding hari kerja Senin-Jumat berjalan
// (pola sama seperti hitungKPI_ untuk absensi).

function hariKerjaBerjalan_(bulan, tahun) {
  var sekarang = new Date();
  var akhirBulan = new Date(tahun, bulan, 0);
  var batasHari = akhirBulan;
  var iniBulanBerjalan = (tahun === sekarang.getFullYear() && bulan === (sekarang.getMonth() + 1));
  if (iniBulanBerjalan && sekarang < akhirBulan) batasHari = sekarang;

  var hariKerja = 0;
  for (var t = new Date(tahun, bulan - 1, 1); t <= batasHari; t.setDate(t.getDate() + 1)) {
    var hari = t.getDay();
    if (hari !== 0 && hari !== 6) hariKerja++;
  }
  return hariKerja < 1 ? 1 : hariKerja;
}

// v6.4: sekarang juga melacak "terakhirLapor" (tanggal+jam laporan paling baru DARI SEMUA WAKTU,
// bukan cuma bulan yang sedang dilihat admin) - dipakai di dashboard & dikirim ke chatbot AI supaya
// bisa jawab pertanyaan seperti "kapan terakhir si Fulan patroli?".
function hitungKPIOperasional_(bulan, tahun, bagianKerjaTarget, sheetData, kolomNamaData) {
  var sekarang = new Date();
  bulan = parseInt(bulan, 10);
  tahun = parseInt(tahun, 10);
  if (isNaN(bulan) || bulan < 1 || bulan > 12) bulan = sekarang.getMonth() + 1;
  if (isNaN(tahun)) tahun = sekarang.getFullYear();

  var sheetKaryawan = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_KARYAWAN);
  var dataKaryawan = sheetKaryawan ? sheetKaryawan.getDataRange().getValues() : [];
  var daftarStaf = [];
  for (var k = 1; k < dataKaryawan.length; k++) {
    var namaK = String(dataKaryawan[k][0] || '').trim();
    var bagianK = String(dataKaryawan[k][3] || '').trim();
    if (!namaK || bagianK !== bagianKerjaTarget) continue;
    daftarStaf.push(namaK);
  }

  var hariKerja = hariKerjaBerjalan_(bulan, tahun);
  var kolomWaktu = kolomNamaData + 1;

  var statPerStaf = {};
  daftarStaf.forEach(function (nm) { statPerStaf[nm] = { nama: nm, jumlahLapor: 0, terakhirTs_: 0, terakhirLapor: '-' }; });

  for (var i = 1; i < sheetData.length; i++) {
    var nm = String(sheetData[i][0] || '').trim();
    if (!statPerStaf[nm]) continue;
    var tglBaris = sheetData[i][kolomNamaData];
    var tglObj = (tglBaris instanceof Date) ? tglBaris : parseTanggalDDMMYYYY_(String(tglBaris).trim());
    if (!tglObj) continue;

    var jamStr = formatJamSel_(sheetData[i][kolomWaktu]);
    var menit = menitDariJam_(jamStr);
    var tsIni = tglObj.getTime() + (menit !== null ? menit * 60000 : 0);
    if (tsIni > statPerStaf[nm].terakhirTs_) {
      statPerStaf[nm].terakhirTs_ = tsIni;
      statPerStaf[nm].terakhirLapor = Utilities.formatDate(tglObj, Session.getScriptTimeZone(), 'dd/MM/yyyy') + (jamStr ? ' ' + jamStr : '');
    }

    if ((tglObj.getMonth() + 1) !== bulan || tglObj.getFullYear() !== tahun) continue;
    statPerStaf[nm].jumlahLapor++;
  }

  var perStaf = [];
  for (var nm2 in statPerStaf) {
    var s = statPerStaf[nm2];
    s.rataRataPerHari = Math.round((s.jumlahLapor / hariKerja) * 10) / 10;
    delete s.terakhirTs_;
    perStaf.push(s);
  }
  perStaf.sort(function (a, b) { return b.jumlahLapor - a.jumlahLapor || a.nama.localeCompare(b.nama); });

  return { bulan: bulan, tahun: tahun, hariKerja: hariKerja, totalStaf: perStaf.length, perStaf: perStaf };
}
