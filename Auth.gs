/**
 * Auth.gs - login admin.
 *
 * v6.3: dua tingkat admin. 'internal' (ADMIN_PASSWORD, PT. Ray) akses penuh.
 * 'client' (CLIENT_ADMIN_PASSWORD, opsional) cuma boleh lihat dashboard
 * KPI/Patroli/Cleaning/AI, TIDAK boleh proses persetujuan pendaftaran/ijin
 * (ditegakkan di masing-masing fungsi lewat cek cekLogin.peran, bukan cuma
 * disembunyikan di frontend).
 */
function loginAdmin(password) {
  var props = PropertiesService.getScriptProperties();
  var passAdmin = props.getProperty('ADMIN_PASSWORD');
  var passClient = props.getProperty('CLIENT_ADMIN_PASSWORD');
  var passInput = String(password || '').trim();

  if (!passAdmin) return { sukses: false, pesan: 'Password admin belum diatur. Jalankan menu Pengaturan dulu.' };
  if (passInput === passAdmin) return { sukses: true, peran: 'internal' };
  if (passClient && passInput === passClient) return { sukses: true, peran: 'client' };
  return { sukses: false, pesan: 'Password admin salah.' };
}
