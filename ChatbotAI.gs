/**
 * ChatbotAI.gs - chatbot AI (admin, READ-ONLY, tanya-jawab seputar KPI).
 * Pakai OpenAI GPT-5.6 Luna (model murah/cepat, cocok untuk tugas ringan
 * seperti ini) via Chat Completions API.
 * Ganti OPENAI_MODEL di sini kalau suatu saat nama model berubah/pensiun
 * (sama seperti kasus Gemini sebelumnya).
 */

var OPENAI_MODEL = 'gpt-5.6-luna';

function tanyaAdminAI(password, pertanyaan, bulan, tahun) {
  var cekLogin = loginAdmin(password);
  if (!cekLogin.sukses) return cekLogin;

  pertanyaan = String(pertanyaan || '').trim();
  if (!pertanyaan) return { sukses: false, pesan: 'Pertanyaan tidak boleh kosong.' };

  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    return { sukses: false, pesan: 'API Key AI belum diatur. Jalankan menu "Pengaturan API Key & Folder Foto" di spreadsheet dulu, lalu isi kolom OpenAI API Key.' };
  }

  try {
    var kpi = hitungKPI_(bulan, tahun);
    var kpiPatroli = hitungKPIPatroli_(bulan, tahun);
    var kpiCleaning = hitungKPICleaning_(bulan, tahun);
    var konteks = bangunKonteksKPI_(kpi) + '\n\n' +
      bangunKonteksOperasional_(kpiPatroli, 'PATROLI (SECURITY)') + '\n\n' +
      bangunKonteksOperasional_(kpiCleaning, 'LAPORAN KEGIATAN (CLEANING SERVICE)');

    var promptLengkap =
      'Kamu adalah asisten AI untuk admin sistem operasional "Absensi RAY" - mencakup absensi karyawan, ' +
      'patroli satpam (Security), dan laporan kegiatan cleaning service. ' +
      'Jawab HANYA berdasarkan data yang diberikan di bawah, dalam Bahasa Indonesia, singkat dan jelas (maks beberapa paragraf pendek atau daftar poin). ' +
      'Kamu bersifat read-only: JANGAN PERNAH berpura-pura melakukan aksi (approve/tolak ijin, ubah data, dsb), cukup jawab pertanyaan/analisa dari data yang ada. ' +
      'Kalau data tidak cukup untuk menjawab dengan yakin, katakan terus terang alih-alih mengarang.\n\n' +
      '=== DATA (bulan ' + kpi.ringkasan.bulan + '/' + kpi.ringkasan.tahun + ') ===\n' +
      konteks + '\n\n' +
      '=== PERTANYAAN ADMIN ===\n' + pertanyaan;

    var url = 'https://api.openai.com/v1/chat/completions';
    var payload = {
      model: OPENAI_MODEL,
      messages: [{ role: 'user', content: promptLengkap }],
      reasoning_effort: 'low'
    };
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var kode = resp.getResponseCode();
    var hasil = JSON.parse(resp.getContentText());

    if (kode !== 200) {
      var pesanErr = (hasil.error && hasil.error.message) ? hasil.error.message : ('HTTP ' + kode);
      return { sukses: false, pesan: 'AI API error: ' + pesanErr };
    }

    var teksJawaban = hasil.choices && hasil.choices[0] && hasil.choices[0].message &&
      hasil.choices[0].message.content;

    if (!teksJawaban) {
      return { sukses: false, pesan: 'AI tidak memberikan jawaban (respons kosong atau diblokir filter keamanan).' };
    }

    return { sukses: true, jawaban: teksJawaban.trim() };
  } catch (err) {
    return { sukses: false, pesan: err.message };
  }
}

function bangunKonteksKPI_(kpi) {
  var baris = [];
  baris.push('Ringkasan: ' + kpi.ringkasan.totalKaryawan + ' karyawan, hari kerja berjalan ' + kpi.ringkasan.hariKerja +
    ' hari, total hadir ' + kpi.ringkasan.totalHadir + ', total telat ' + kpi.ringkasan.totalTelat +
    ', total ijin disetujui ' + kpi.ringkasan.totalIjin + ', rata-rata kehadiran ' + kpi.ringkasan.rataRataKehadiran + '%.');

  baris.push('\nDetail per karyawan (nama | bagian kerja | hadir | telat | ijin | % kehadiran):');
  kpi.perKaryawan.forEach(function (k) {
    baris.push(k.nama + ' | ' + k.bagianKerja + ' | ' + k.hadir + ' | ' + k.telat + ' | ' + k.ijin + ' | ' + k.persenKehadiran + '%' +
      (k.adaJamStandar ? '' : ' (jam standar belum diatur untuk bagian ini, telat tidak terhitung)'));
  });

  baris.push('\nTren 6 bulan terakhir, format "bulan: hadir/telat/ijin":');
  kpi.grafik.forEach(function (g) {
    baris.push(g.label + ': ' + g.hadir + '/' + g.telat + '/' + g.ijin);
  });

  var lokasiResmi = getDaftarLokasiValid_();
  if (lokasiResmi.length) {
    baris.push('\nDaftar lokasi kerja resmi terdaftar: ' +
      lokasiResmi.map(function (l) { return l.nama + ' (radius ' + Math.round(l.radius) + 'm)'; }).join(', ') +
      '. Catatan: setiap absen di sistem ini SUDAH divalidasi berada dalam radius salah satu lokasi resmi di atas (absen di luar radius otomatis ditolak server), jadi "lokasi mencurigakan" di sini lebih berarti pola aneh seperti koordinat GPS identik dipakai beberapa karyawan berbeda di hari sama (indikasi titip absen), bukan absen di luar area kantor.');
  }

  return baris.join('\n');
}

// v6.4: konteks Patroli/Laporan Kegiatan untuk chatbot AI - sebelumnya AI cuma diberi data absensi,
// jadi tidak bisa jawab pertanyaan seperti "kapan terakhir si Fulan patroli?". judul contoh: "PATROLI (SECURITY)".
function bangunKonteksOperasional_(kpi, judul) {
  var baris = [];
  baris.push('=== DATA ' + judul + ' (bulan ' + kpi.bulan + '/' + kpi.tahun + ') ===');
  baris.push('Total staf: ' + kpi.totalStaf + ', hari kerja berjalan bulan ini: ' + kpi.hariKerja + '.');
  if (!kpi.perStaf.length) {
    baris.push('Belum ada staf terdaftar untuk divisi ini.');
  } else {
    baris.push('Detail per staf (nama | jumlah lapor bulan ini | rata-rata lapor per hari | tanggal+jam laporan TERAKHIR dari semua waktu):');
    kpi.perStaf.forEach(function (s) {
      baris.push(s.nama + ' | ' + s.jumlahLapor + ' | ' + s.rataRataPerHari + ' | ' + s.terakhirLapor);
    });
  }
  return baris.join('\n');
}
