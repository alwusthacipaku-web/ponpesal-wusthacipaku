/**
 * Code.gs — Website Pondok Pesantren Al-Wustha (Google Apps Script)
 *
 * LANGKAH PASANG
 * 1. Buat Google Sheets baru, salin ID-nya (bagian di antara /d/ dan /edit pada URL).
 * 2. script.google.com > New project. Tempel file ini sebagai Code.gs,
 *    lalu tambah file HTML bernama "index" dan tempel isi index.html.
 * 3. Isi SPREADSHEET_ID di bawah.
 * 4. Jalankan fungsi setup() satu kali (klik Run, izinkan akses). Ini membuat
 *    semua sheet dan mengisi kode akses awal.
 * 5. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone.
 *    Setiap kali ubah kode, buat versi deployment baru.
 *
 * KODE AKSES AWAL (juga ada di sheet "Kode_Akses", bebas diubah/ditambah):
 *   Admin        : ADM-WSTH-7K4Q
 *   Guru/Mudaris : GRU-WSTH-3M9P
 *   Santri       : STR-WSTH-5X2R
 *   Orang Tua    : ORT-WSTH-8D6T
 */

const SPREADSHEET_ID = 'ISI_ID_SPREADSHEET_ANDA';
const FOLDER_BERKAS = 'Berkas PSB Al-Wustha'; // folder Drive dibuat otomatis
const SESSION_DETIK = 6 * 60 * 60;            // login berlaku 6 jam

const SHEETS = {
  Kode_Akses: ['Kode', 'Peran', 'Nama', 'Aktif'],
  Pendaftaran: ['ID', 'Waktu', 'Jenjang', 'Nama', 'Jenis Kelamin', 'Tempat Lahir', 'Tanggal Lahir',
                'Asal Sekolah', 'Alamat', 'WhatsApp', 'Akta', 'KK', 'Ijazah', 'Pas Foto', 'Status', 'Catatan'],
  Pengumuman: ['Waktu', 'Judul', 'Isi', 'Penulis']
};

const KODE_AWAL = [
  ['ADM-WSTH-7K4Q', 'admin',  'Administrator',  'YA'],
  ['GRU-WSTH-3M9P', 'guru',   'Guru / Mudaris', 'YA'],
  ['STR-WSTH-5X2R', 'santri', 'Santri',         'YA'],
  ['ORT-WSTH-8D6T', 'ortu',   'Orang Tua',      'YA']
];

/* ---------- Halaman ---------- */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Pondok Pesantren Al-Wustha')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Setup ---------- */
function setup() {
  Object.keys(SHEETS).forEach(function (n) { sheet_(n); });
}

function sheet_(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(SHEETS[name]);
    sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold').setBackground('#0b4d24').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    if (name === 'Kode_Akses') {
      sh.getRange(2, 1, KODE_AWAL.length, 4).setValues(KODE_AWAL);
    }
  }
  return sh;
}

/* ---------- Login & sesi ---------- */
function login(role, code) {
  try {
    code = String(code || '').trim().toUpperCase();
    role = String(role || '').trim();
    if (!code) return { ok: false, message: 'Masukkan kode akses.' };

    const rows = sheet_('Kode_Akses').getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (String(r[0]).trim().toUpperCase() === code &&
          String(r[1]).trim() === role &&
          String(r[3]).trim().toUpperCase() !== 'TIDAK') {
        const token = Utilities.getUuid();
        const sesi = { role: role, nama: String(r[2] || '') };
        CacheService.getScriptCache().put('s_' + token, JSON.stringify(sesi), SESSION_DETIK);
        return { ok: true, token: token, role: role, nama: sesi.nama };
      }
    }
    return { ok: false, message: 'Kode tidak valid atau tidak sesuai dengan peran yang dipilih.' };
  } catch (e) {
    return { ok: false, message: 'Kesalahan server: ' + e.message };
  }
}

function logout(token) {
  CacheService.getScriptCache().remove('s_' + token);
  return true;
}

function sesi_(token) {
  const v = token ? CacheService.getScriptCache().get('s_' + token) : null;
  return v ? JSON.parse(v) : null;
}

function cekSesi(token) {
  const s = sesi_(token);
  return s ? { ok: true, role: s.role, nama: s.nama } : { ok: false };
}

/* ---------- PSB ---------- */
function norm_(n) {
  n = String(n || '').replace(/\D/g, '');
  if (n.indexOf('62') === 0) n = '0' + n.slice(2);
  return n;
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_BERKAS);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_BERKAS);
}

function simpanBerkas_(f, id, label) {
  if (!f || !f.data) return '';
  const bytes = Utilities.base64Decode(f.data);
  if (bytes.length > 3 * 1024 * 1024) throw new Error('Ukuran berkas ' + label + ' maksimal 3 MB.');
  const blob = Utilities.newBlob(bytes, f.type || 'application/octet-stream', id + '_' + label + '_' + (f.name || ''));
  return folder_().createFile(blob).getUrl();
}

function submitPSB(d) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const wajib = ['jenjang', 'nama', 'jk', 'tempat', 'tgl', 'asal', 'alamat', 'wa'];
    for (let i = 0; i < wajib.length; i++) {
      if (!String(d[wajib[i]] || '').trim()) return { ok: false, message: 'Mohon lengkapi semua kolom bertanda *.' };
    }
    const wa = norm_(d.wa);
    if (wa.length < 9 || wa.length > 15) return { ok: false, message: 'Nomor WhatsApp tidak valid.' };

    const sh = sheet_('Pendaftaran');
    const tahun = new Date().getFullYear();
    const urut = String(sh.getLastRow()).padStart(4, '0'); // baris 1 = header
    const id = 'PSB-' + tahun + '-' + urut;

    const b = d.berkas || {};
    const akta = simpanBerkas_(b.akta, id, 'Akta');
    const kk = simpanBerkas_(b.kk, id, 'KK');
    const ijazah = simpanBerkas_(b.ijazah, id, 'Ijazah');
    const foto = simpanBerkas_(b.foto, id, 'Foto');

    sh.appendRow([id, new Date(), d.jenjang, d.nama.trim(), d.jk, d.tempat.trim(), d.tgl,
                  d.asal.trim(), d.alamat.trim(), "'" + wa, akta, kk, ijazah, foto, 'Menunggu Verifikasi', '']);
    return { ok: true, id: id, message: 'Pendaftaran berhasil. Simpan Nomor Pendaftaran Anda: ' + id };
  } catch (e) {
    return { ok: false, message: 'Gagal menyimpan: ' + e.message };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function cekStatus(id, wa) {
  id = String(id || '').trim().toUpperCase();
  wa = norm_(wa);
  if (!id || !wa) return { ok: false, message: 'Isi Nomor Pendaftaran dan No. WhatsApp.' };
  const rows = sheet_('Pendaftaran').getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toUpperCase() === id && norm_(rows[i][9]) === wa) {
      return { ok: true, nama: rows[i][3], jenjang: rows[i][2], status: rows[i][14], catatan: rows[i][15] || '' };
    }
  }
  return { ok: false, message: 'Data tidak ditemukan. Periksa kembali nomor pendaftaran dan WhatsApp.' };
}

/* ---------- Admin ---------- */
function admin_(token) {
  const s = sesi_(token);
  if (!s || s.role !== 'admin') throw new Error('Akses ditolak. Silakan masuk sebagai admin.');
  return s;
}

function getPendaftar(token) {
  admin_(token);
  const rows = sheet_('Pendaftaran').getDataRange().getValues();
  const out = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    out.push({
      id: r[0], waktu: Utilities.formatDate(new Date(r[1]), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm'),
      jenjang: r[2], nama: r[3], jk: r[4], wa: String(r[9]), asal: r[7],
      akta: r[10], kk: r[11], ijazah: r[12], foto: r[13], status: r[14], catatan: r[15]
    });
  }
  return out;
}

function updateStatus(token, id, status, catatan) {
  admin_(token);
  const boleh = ['Menunggu Verifikasi', 'Diverifikasi', 'Diterima', 'Ditolak'];
  if (boleh.indexOf(status) < 0) return { ok: false, message: 'Status tidak valid.' };
  const sh = sheet_('Pendaftaran');
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) {
      sh.getRange(i + 1, 15).setValue(status);
      sh.getRange(i + 1, 16).setValue(catatan || '');
      return { ok: true };
    }
  }
  return { ok: false, message: 'ID tidak ditemukan.' };
}

/* ---------- Pengumuman ---------- */
function getPengumuman(token) {
  if (!sesi_(token)) throw new Error('Sesi berakhir, silakan masuk kembali.');
  const rows = sheet_('Pengumuman').getDataRange().getValues();
  const out = [];
  for (let i = rows.length - 1; i >= 1 && out.length < 20; i--) {
    out.push({
      waktu: Utilities.formatDate(new Date(rows[i][0]), 'Asia/Jakarta', 'dd MMM yyyy'),
      judul: rows[i][1], isi: rows[i][2], penulis: rows[i][3]
    });
  }
  return out;
}

function addPengumuman(token, judul, isi) {
  const s = admin_(token);
  judul = String(judul || '').trim();
  isi = String(isi || '').trim();
  if (!judul || !isi) return { ok: false, message: 'Judul dan isi wajib diisi.' };
  sheet_('Pengumuman').appendRow([new Date(), judul, isi, s.nama]);
  return { ok: true };
}
