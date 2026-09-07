/**
 * SISTEM ABSENSI RAY - API BACKEND (v6.4, untuk GitHub Pages)
 * ------------------------------------------------------------
 * Tampilan (HTML/CSS/JS/ikon) di-host terpisah di GitHub Pages.
 * File-file .gs di project ini HANYA melayani data (JSON) lewat doPost,
 * dipanggil via fetch() dari luar.
 *
 * Struktur file (dipecah dari 1 file besar, semua tetap 1 global scope
 * seperti biasa di Apps Script - urutan file tidak masalah untuk fungsi,
 * hanya untuk "var" di top-level yang dieksekusi saat load):
 *   - Main.gs        : konstanta, menu, doGet/doPost (router), jsonResponse_
 *   - Utils.gs       : helper umum (jarak GPS, parsing tanggal/jam, geocoding)
 *   - Auth.gs        : loginAdmin (2 tingkat: internal & client)
 *   - Karyawan.gs    : data master karyawan & lokasi, login karyawan
 *   - Absensi.gs     : simpanAbsen, riwayat absensi
 *   - Pendaftaran.gs : registrasi mandiri karyawan + approval admin
 *   - Ijin.gs        : pengajuan ijin/cuti + approval admin
 *   - JamKerja.gs    : sheet "Jam Kerja Standar" (dipakai hitung telat)
 *   - Dashboard.gs   : KPI Disiplin (absensi) - dipakai admin & chatbot AI
 *   - Patroli.gs     : modul Patroli (Security)
 *   - Cleaning.gs    : modul Laporan Kegiatan (Cleaning Service)
 *   - ChatbotAI.gs   : chatbot AI read-only (OpenAI, model di OPENAI_MODEL)
 *
 * v6.4: hitungKPIOperasional_ juga lacak "terakhirLapor" (tanggal+jam laporan
 * terakhir dari semua waktu) per staf; chatbot AI sekarang ikut dikasih
 * konteks data Patroli & Laporan Kegiatan.
 *
 * Deployment: Deploy > New deployment > Web app
 * Execute as: Me
 * Who has access: Anyone
 */

var SHEET_ABSENSI = 'Data Absensi';
var SHEET_LOKASI = 'Data Master Lokasi';
var SHEET_KARYAWAN = 'Data Master Karyawan';

// v6.2: Bagian Kerja sekarang dibatasi ke daftar tetap ini (dulu bebas teks) supaya frontend bisa
// otomatis menampilkan menu Patroli (Security) / Laporan Kegiatan (Cleaning Service) sesuai divisi.
// Karyawan lama yang Bagian Kerja-nya belum cocok salah satu ini harus diedit manual oleh admin
// langsung di sheet "Data Master Karyawan" supaya menu barunya muncul.
var DAFTAR_BAGIAN_KERJA = ['Security', 'Cleaning Service', 'Administrasi', 'Lainnya'];


// ============ MENU & PENGATURAN (tetap di spreadsheet) ============

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Menu Absensi RAY')
    .addItem('Pengaturan API Key & Folder Foto', 'aturPengaturan')
    .addToUi();
}

function aturPengaturan() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var resFolder = ui.prompt(
    'Folder Foto Absensi',
    'Masukkan ID Folder Google Drive untuk menyimpan foto absensi:\n(Ambil dari URL folder Drive)',
    ui.ButtonSet.OK_CANCEL
  );
  if (resFolder.getSelectedButton() !== ui.Button.OK) return;
  props.setProperty('FOLDER_ID', resFolder.getResponseText().trim());

  var resKey = ui.prompt(
    'Google Maps API Key',
    'Masukkan API Key Google Maps (harus sudah aktifkan Geocoding API):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resKey.getSelectedButton() !== ui.Button.OK) return;
  props.setProperty('GOOGLE_MAPS_API_KEY', resKey.getResponseText().trim());

  var resAdmin = ui.prompt(
    'Password Admin (PT. Ray Mitra Perkasa)',
    'Masukkan password untuk login panel admin INTERNAL (akses penuh, termasuk persetujuan pendaftaran & ijin):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resAdmin.getSelectedButton() !== ui.Button.OK) return;
  props.setProperty('ADMIN_PASSWORD', resAdmin.getResponseText().trim());

  var resClientAdmin = ui.prompt(
    'Password Admin Client (opsional)',
    'Masukkan password untuk login panel admin CLIENT (dashboard KPI/Patroli/Cleaning/AI saja, TANPA persetujuan pendaftaran & ijin)\n(kosongkan/Batal kalau belum perlu, bisa diisi belakangan lewat menu ini lagi):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resClientAdmin.getSelectedButton() === ui.Button.OK && resClientAdmin.getResponseText().trim()) {
    props.setProperty('CLIENT_ADMIN_PASSWORD', resClientAdmin.getResponseText().trim());
  }

  var resOpenAi = ui.prompt(
    'OpenAI API Key (opsional)',
    'Masukkan API Key OpenAI (model GPT-5.6 Luna) untuk mengaktifkan chatbot AI di panel admin\n(kosongkan/Batal kalau belum punya, bisa diisi belakangan lewat menu ini lagi):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resOpenAi.getSelectedButton() === ui.Button.OK && resOpenAi.getResponseText().trim()) {
    props.setProperty('OPENAI_API_KEY', resOpenAi.getResponseText().trim());
  }

  ui.alert('Sukses', 'Pengaturan berhasil disimpan.', ui.ButtonSet.OK);
}


// ============ ENTRY POINT API ============

function doGet(e) {
  return jsonResponse_({ ok: true, pesan: 'Absensi RAY API aktif. Gunakan POST untuk mengambil/mengirim data.' });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (errParse) {
    return jsonResponse_({ sukses: false, pesan: 'Payload tidak valid.' });
  }

  var action = body.action;
  try {
    switch (action) {
      case 'daftarKaryawan':
        return jsonResponse_({ sukses: true, daftar: getDaftarKaryawanPublik() });
      case 'daftarLokasi':
        return jsonResponse_({ sukses: true, daftar: getDaftarLokasiValid_() });
      case 'login':
        return jsonResponse_(loginKaryawan(body.nama, body.pin));
      case 'statusUlang':
        return jsonResponse_(getStatusUlang(body.nama, body.pin));
      case 'riwayatAbsensi':
        return jsonResponse_(getRiwayatAbsensi(body.nama, body.pin, 7));
      case 'daftarBaru':
        return jsonResponse_(daftarKaryawanBaru(body.nama, body.bagianKerja, body.pin));
      case 'loginAdmin':
        return jsonResponse_(loginAdmin(body.password));
      case 'daftarPendaftaranBaru':
        return jsonResponse_(getDaftarPendaftaranBaru(body.password));
      case 'setujuiPendaftaran':
        return jsonResponse_(setujuiPendaftaran(body.password, body.baris));
      case 'tolakPendaftaran':
        return jsonResponse_(tolakPendaftaran(body.password, body.baris));
      case 'ajukanIjin':
        return jsonResponse_(ajukanIjin(body.nama, body.pin, body.jenisIjin, body.tanggalIjin, body.keterangan, body.buktiBase64, body.mimeType));
      case 'daftarIjinAdmin':
        return jsonResponse_(getDaftarIjinAdmin(body.password));
      case 'setujuiIjin':
        return jsonResponse_(setujuiIjin(body.password, body.baris));
      case 'tolakIjin':
        return jsonResponse_(tolakIjin(body.password, body.baris));
      case 'ambilDashboardKPI':
        return jsonResponse_(ambilDashboardKPI(body.password, body.bulan, body.tahun));
      case 'ambilDashboardPatroli':
        return jsonResponse_(ambilDashboardPatroli(body.password, body.bulan, body.tahun));
      case 'ambilDashboardCleaning':
        return jsonResponse_(ambilDashboardCleaning(body.password, body.bulan, body.tahun));
      case 'tanyaAdminAI':
        return jsonResponse_(tanyaAdminAI(body.password, body.pertanyaan, body.bulan, body.tahun));
      case 'simpanAbsen':
        return jsonResponse_(simpanAbsen(
          body.nama, body.pin, body.jenisAbsen, body.fotoBase64,
          body.mimeType, body.lat, body.lng, body.akurasi
        ));
      case 'daftarTitikPatroli':
        return jsonResponse_({ sukses: true, daftar: getDaftarTitikPatroli() });
      case 'simpanPatroli':
        return jsonResponse_(simpanPatroli(
          body.nama, body.pin, body.namaPos, body.fotoBase64,
          body.mimeType, body.lat, body.lng, body.akurasi, body.catatan
        ));
      case 'daftarTugasKebersihan':
        return jsonResponse_({ sukses: true, daftar: getDaftarTugasKebersihan() });
      case 'simpanLaporanKegiatan':
        return jsonResponse_(simpanLaporanKegiatan(
          body.nama, body.pin, body.tugas, body.area, body.lat, body.lng, body.akurasi,
          body.fotoSebelumBase64, body.fotoSesudahBase64, body.mimeType, body.catatan
        ));
      default:
        return jsonResponse_({ sukses: false, pesan: 'Aksi tidak dikenal: ' + action });
    }
  } catch (errAksi) {
    return jsonResponse_({ sukses: false, pesan: errAksi.message });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
