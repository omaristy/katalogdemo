/**
 * BACKEND KATALOG - Google Apps Script
 * Pasang di: Google Sheet > Ekstensi > Apps Script
 * Lalu Deploy > New deployment > Web app
 *   Execute as: Me | Who has access: Anyone
 *
 * Nilai di CONFIG hanya dipakai sebagai nilai AWAL. Setelah itu semuanya
 * (nama toko, WA, logo, tema, password) diatur di panel admin atau langsung
 * di tab "Pengaturan" pada Google Sheet.
 */
const CONFIG = {
  ADMIN_PASSWORD: '1111',   // password awal (bisa diganti di admin / di Sheet)
  WA_NUMBER: '6281234567890',
  STORE_NAME: 'Toko Saya',
  SHEET_NAME: 'Produk',
  FOLDER_NAME: 'Foto Katalog'
};
const HEAD = ['id', 'nama', 'kategori', 'harga', 'satuan', 'gambar', 'deskripsi', 'aktif', 'harga_diskon', 'stok', 'keterangan'];
const SKEYS = ['store', 'wa', 'logo', 'tema', 'password', 'jam', 'zona', 'tutup', 'logobentuk', 'logowarna', 'logoisi'];

// ---- Cache (dipecah per 30.000 karakter agar aman dari batas 100 KB) ----
const TTL = 21600;   // 6 jam; cache otomatis dibangun ulang setiap ada perubahan data
function cget_(k) {
  const c = CacheService.getScriptCache(), n = c.get(k + '_n'); if (!n) return null;
  const keys = []; for (let i = 0; i < +n; i++) keys.push(k + '_' + i);
  const m = c.getAll(keys); let s = '';
  for (const key of keys) { if (m[key] === undefined) return null; s += m[key]; }
  return s;
}
function cput_(k, txt) {
  const parts = []; for (let i = 0; i < txt.length; i += 30000) parts.push(txt.slice(i, i + 30000));
  const o = {}; parts.forEach((p, i) => o[k + '_' + i] = p); o[k + '_n'] = String(parts.length);
  try { CacheService.getScriptCache().putAll(o, TTL); } catch (e) {}
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!s) { s = ss.insertSheet(CONFIG.SHEET_NAME); s.setFrozenRows(1); }
  const c = CacheService.getScriptCache();
  if (!c.get('hdr')) {                                        // cek header cukup sekali per beberapa jam
    const cur = s.getRange(1, 1, 1, HEAD.length).getValues()[0];
    if (cur.join('|') !== HEAD.join('|')) s.getRange(1, 1, 1, HEAD.length).setValues([HEAD]);
    c.put('hdr', '1', TTL);
  }
  return s;
}
function rows_() {
  const hit = cget_('rows');                                  // daftar produk di-cache -> login admin cepat
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const v = sheet_().getDataRange().getValues();
  const out = v.slice(1).filter(r => r[0] !== '').map(r => {
    const o = {}; HEAD.forEach((h, i) => o[h] = r[i] === undefined ? '' : r[i]);
    o.harga = Number(o.harga) || 0;
    o.harga_diskon = Number(o.harga_diskon) || 0;
    const sv = String(o.stok).trim();                       // kosong = stok tidak dibatasi
    o.stok = (sv === '' || isNaN(Number(sv))) ? null : Math.max(0, Math.floor(Number(sv)));
    return o;
  });
  cput_('rows', JSON.stringify(out));
  return out;
}
function folder_() {
  const it = DriveApp.getFoldersByName(CONFIG.FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.FOLDER_NAME);
}

// ---- Pengaturan (tab "Pengaturan": kolom kunci | nilai) ----
function writeSettings_(s, o) {
  s.getRange('B:B').setNumberFormat('@');
  const v = [['kunci', 'nilai']].concat(SKEYS.map(k => [k, o[k] === undefined ? '' : o[k]]));
  s.getRange(1, 1, v.length, 2).setValues(v);
}
function settings_() {
  const hit = cget_('set');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const def = { store: CONFIG.STORE_NAME, wa: CONFIG.WA_NUMBER, logo: '', tema: 'oranye', password: CONFIG.ADMIN_PASSWORD, jam: '', zona: 'Asia/Jakarta', tutup: '', logobentuk: 'bulat', logowarna: 'putih', logoisi: 'putih' };
  let s = ss.getSheetByName('Pengaturan');
  if (!s) { s = ss.insertSheet('Pengaturan'); writeSettings_(s, def); return def; }
  const o = Object.assign({}, def);
  const rows = s.getDataRange().getValues().slice(1);
  rows.forEach(r => { if (r[0] && String(r[1]) !== '') o[r[0]] = String(r[1]); });
  const have = rows.map(r => r[0]);
  if (SKEYS.some(k => have.indexOf(k) < 0)) writeSettings_(s, o);   // sheet lama: tambahkan baris pengaturan yang belum ada
  cput_('set', JSON.stringify(o));
  return o;
}
function saveSettingsObj_(o) {
  settings_();
  writeSettings_(SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengaturan'), o);
  CacheService.getScriptCache().remove('set_n');   // paksa baca ulang
}
function waFix_(v) {
  let n = String(v).replace(/\D/g, '');
  if (n.startsWith('0')) n = '62' + n.slice(1);
  else if (n.startsWith('8')) n = '62' + n;
  return n;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Katalog publik: hanya produk aktif. Password TIDAK pernah dikirim.
function buildPub_() {
  try { CacheService.getScriptCache().remove('rows_n'); } catch (e) {}   // data berubah -> baca ulang dari Sheet
  const list = rows_().filter(p => String(p.aktif).toLowerCase() !== 'tidak');
  const st = settings_();
  const txt = JSON.stringify({ ok: true, store: st.store, wa: st.wa, logo: st.logo, tema: st.tema, jam: st.jam, zona: st.zona, tutup: st.tutup, logobentuk: st.logobentuk, logowarna: st.logowarna, logoisi: st.logoisi, products: list });
  cput_('pub', txt);
  return txt;
}
function doGet() {
  const txt = cget_('pub') || buildPub_();      // hampir selalu dari cache -> sangat cepat
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

// Panel admin: wajib password
function handle_(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    const st = settings_();
    if (String(b.password) !== st.password) return out_({ ok: false, error: 'Password salah' });
    const s = sheet_();

    if (b.action === 'ping') return out_({ ok: true, tema: st.tema });   // cek password saja (sangat cepat)
    if (b.action === 'list') return out_({ ok: true, products: rows_(), tema: st.tema });

    if (b.action === 'getSettings') {
      const o = Object.assign({}, st); delete o.password;
      return out_(Object.assign({ ok: true }, o));
    }
    if (b.action === 'saveSettings') {
      const wa = waFix_(b.wa), store = String(b.store || '').trim();
      if (!store || wa.length < 9) return out_({ ok: false, error: 'Nama atau nomor tidak valid' });
      st.store = store; st.wa = wa;
      st.logo = /^https:\/\//.test(String(b.logo || '')) ? String(b.logo) : '';
      st.tema = /^[a-z0-9]{3,12}$/.test(String(b.tema || '')) ? String(b.tema) : 'oranye';
      // jam operasional, format: 1=08:00-17:00;2=08:00-17:00;...;0=tutup  (0=Minggu ... 6=Sabtu)
      const J = '(tutup|\\d\\d:\\d\\d-\\d\\d:\\d\\d)';
      st.jam = new RegExp('^[0-6]=' + J + '(;[0-6]=' + J + '){0,6}$').test(String(b.jam || '')) ? String(b.jam) : '';
      st.zona = ['Asia/Jakarta', 'Asia/Makassar', 'Asia/Jayapura'].indexOf(b.zona) >= 0 ? b.zona : 'Asia/Jakarta';
      st.tutup = b.tutup === 'ya' ? 'ya' : '';
      st.logobentuk = b.logobentuk === 'asli' ? 'asli' : 'bulat';
      const WARNA = ['putih', 'hitam', 'merah', 'kuning', 'biru', 'oranye', 'hijau', 'ungu', 'pink', 'tanpa'];
      st.logowarna = WARNA.indexOf(b.logowarna) >= 0 ? b.logowarna : 'putih';   // bingkai
      st.logoisi = WARNA.indexOf(b.logoisi) >= 0 ? b.logoisi : 'putih';         // isi lingkaran
      saveSettingsObj_(st);
      return out_({ ok: true, store: st.store, wa: st.wa, logo: st.logo, tema: st.tema });
    }
    if (b.action === 'changePassword') {
      const np = String(b.newPassword || '');
      if (np.length < 6) return out_({ ok: false, error: 'Password minimal 6 karakter' });
      st.password = np; saveSettingsObj_(st);
      return out_({ ok: true });
    }
    if (b.action === 'upload') {
      const EXT = { 'image/png': '.png', 'image/webp': '.webp', 'image/jpeg': '.jpg' };   // logo transparan = png/webp
      const mime = EXT[b.mime] ? b.mime : 'image/jpeg';
      const nm = (String(b.name || 'foto').replace(/[^\w\- ]/g, '') || 'foto') + '_' + Date.now() + EXT[mime];
      const f = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(b.data), mime, nm));
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return out_({ ok: true, url: 'https://drive.google.com/file/d/' + f.getId() + '/view' });
    }
    if (b.action === 'save') {
      const d = b.data;
      const sv = String(d.stok === undefined ? '' : d.stok).trim();
      d.stok = /^\d+$/.test(sv) ? Number(sv) : '';          // kosong = tidak dibatasi
      d.keterangan = String(d.keterangan || '').trim().slice(0, 80);
      const rec = HEAD.map(h => d[h] === undefined ? '' : d[h]);
      if (d.id) {
        const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
        const i = ids.indexOf(String(d.id));
        if (i < 1) return out_({ ok: false, error: 'Produk tidak ditemukan' });
        s.getRange(i + 1, 1, 1, HEAD.length).setValues([rec]);
      } else {
        rec[0] = 'P' + Date.now();
        s.appendRow(rec);
      }
      return out_({ ok: true });
    }
    if (b.action === 'delete') {
      const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
      const i = ids.indexOf(String(b.id));
      if (i < 1) return out_({ ok: false, error: 'Produk tidak ditemukan' });
      s.deleteRow(i + 1);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'Aksi tidak dikenal' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

// Hapus cache katalog setiap kali admin mengubah data
function doPost(e) {
  const res = handle_(e);
  try {
    const b = JSON.parse(e.postData.contents);
    if (['save', 'delete', 'saveSettings'].indexOf(b.action) >= 0 && res.getContent().indexOf('"ok":true') >= 0)
      buildPub_();     // bangun ulang cache sekarang, jadi pembeli berikutnya tetap cepat
  } catch (err) {}
  return res;
}
// Otomatis dipanggil saat Anda mengedit Sheet secara manual (mis. ubah stok) -> katalog langsung ikut berubah
function onEdit() { try { CacheService.getScriptCache().removeAll(['hdr', 'set_n']); buildPub_(); } catch (e) {} }

/** JALANKAN SEKALI dari editor (pilih "setup" lalu Jalankan): minta izin & siapkan sheet + folder foto. */
function setup() {
  sheet_(); settings_();
  const f = folder_();
  const t = f.createFile('tes.txt', 'ok');
  t.setTrashed(true);
  Logger.log('BERHASIL. Folder foto: ' + f.getUrl());
}
