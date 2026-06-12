// ============================================================
// ABSENSI PEKERJA — Google Apps Script (GAS)  v2
// ============================================================
// Setup:
//   1. Buka https://script.google.com → New Project
//   2. Paste seluruh kode ini, klik Save
//   3. Deploy → New Deployment → Web App
//      - Execute as : Me
//      - Who has access : Anyone (anonymous)
//   4. Copy URL deployment → paste ke API_URL di index.html
//
// Sheet yang dibutuhkan (otomatis dibuat jika belum ada):
//   - Karyawan  : ID | Nama | Jabatan | Tim | Aktif
//   - Absen     : Tanggal | EmpID | Status | Keterangan | Nama
//   - Tim       : ID | Nama
// ============================================================

const SS      = SpreadsheetApp.getActiveSpreadsheet();
const SH_KAR  = 'Karyawan';
const SH_ABS  = 'Absen';
const SH_TIM  = 'Tim';

// ── Utility ──────────────────────────────────────────────────
function ok(data)  {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', data }))
    .setMimeType(ContentService.MimeType.JSON);
}
function err(msg)  {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name, headers) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff');
  }
  return sh;
}

// Pastikan sheet Absen punya kolom Nama (migrasi otomatis)
function ensureAbsenHasNama() {
  const sh = getOrCreateSheet(SH_ABS, ['Tanggal','EmpID','Status','Keterangan','Nama']);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (!headers.includes('Nama')) {
    // Tambahkan kolom Nama di akhir
    const col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue('Nama').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  }
  return sh;
}

function sheetToObjects(sh) {
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
}

// Normalisasi ID: strip leading zeros agar 00070108 == 70108
function normalizeEmpId(id) {
  const s = String(id || '').trim();
  // Pure numeric → strip leading zeros
  if (/^\d+$/.test(s)) return String(parseInt(s, 10));
  // Alphanumeric (MS1, MS2) → uppercase
  return s.toUpperCase();
}

// Format tanggal jadi YYYY-MM-DD (handle Date object dari Sheets)
function fmtDateKey(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(val).trim().slice(0, 10);
}

// Generate ID unik
function generateId(prefix) {
  return prefix + '-' + new Date().getTime().toString(36).toUpperCase();
}

// ── doGet — router utama (GET dengan ?action=... atau tanpa action = ambil semua data) ──
function doGet(e) {
  try {
    const action = e && e.parameter && e.parameter.action;

    // Kalau ada action → routing ke handler yang sesuai
    if (action) {
      let data = {};
      try { data = JSON.parse((e.parameter.data) || '{}'); } catch(ex) { data = {}; }

      switch (action) {
        case 'saveAbsen':      return actionSaveAbsen(data);
        case 'deleteAbsen':    return actionDeleteAbsen(data);
        case 'addKaryawan':    return actionAddKaryawan(data);
        case 'toggleKaryawan': return actionToggleKaryawan(data);
        case 'addTim':         return actionAddTim(data);
        case 'deleteTim':      return actionDeleteTim(data);
        default:               return err('Action tidak dikenal: ' + action);
      }
    }

    // Tanpa action → kembalikan semua data
    return actionGetAll();

  } catch(ex) {
    return err('doGet error: ' + ex.message);
  }
}

// doPost — jaga kompatibilitas jika suatu saat dikirim via POST
function doPost(e) {
  try {
    let action = '', data = {};
    if (e.postData && e.postData.contents) {
      // Coba parse sebagai JSON
      try {
        const body = JSON.parse(e.postData.contents);
        action = body.action || '';
        data   = body.data   || {};
      } catch(ex) {
        // Coba parse sebagai form-encoded
        const params = e.parameter || {};
        action = params.action || '';
        try { data = JSON.parse(params.data || '{}'); } catch(ex2) { data = {}; }
      }
    } else {
      const params = e.parameter || {};
      action = params.action || '';
      try { data = JSON.parse(params.data || '{}'); } catch(ex) { data = {}; }
    }

    switch (action) {
      case 'saveAbsen':      return actionSaveAbsen(data);
      case 'deleteAbsen':    return actionDeleteAbsen(data);
      case 'addKaryawan':    return actionAddKaryawan(data);
      case 'toggleKaryawan': return actionToggleKaryawan(data);
      case 'addTim':         return actionAddTim(data);
      case 'deleteTim':      return actionDeleteTim(data);
      default:               return err('Action tidak dikenal: ' + action);
    }
  } catch(ex) {
    return err('doPost error: ' + ex.message);
  }
}

// ============================================================
// ACTION: GET ALL DATA
// ============================================================
function actionGetAll() {
  const shKar = getOrCreateSheet(SH_KAR, ['ID','Nama','Jabatan','Tim','Aktif']);
  const shAbs = ensureAbsenHasNama();
  const shTim = getOrCreateSheet(SH_TIM, ['ID','Nama']);

  const karyawan = sheetToObjects(shKar).map(k => ({
    ID:      String(k['ID']      || '').trim(),
    Nama:    String(k['Nama']    || '').trim(),
    Jabatan: String(k['Jabatan'] || '').trim(),
    Tim:     String(k['Tim']     || '').trim(),
    Aktif:   k['Aktif'] === true || String(k['Aktif']).toUpperCase() === 'TRUE'
  })).filter(k => k.ID && k.Nama);

  const absen = sheetToObjects(shAbs).map(a => ({
    Tanggal:    fmtDateKey(a['Tanggal']),
    EmpID:      String(a['EmpID']      || '').trim(),
    Status:     String(a['Status']     || '').trim(),
    Keterangan: String(a['Keterangan'] || '').trim(),
    Nama:       String(a['Nama']       || '').trim()
  })).filter(a => a.Tanggal && a.EmpID && a.Status);

  const tim = sheetToObjects(shTim).map(t => ({
    ID:   String(t['ID']   || '').trim(),
    Nama: String(t['Nama'] || '').trim()
  })).filter(t => t.Nama);

  return ok({ karyawan, absen, tim });
}

// ============================================================
// ACTION: SAVE ABSEN
// ============================================================
function actionSaveAbsen(data) {
  const { tanggal, empId, nama, status, ket } = data;
  if (!tanggal || !empId || !status) return err('Field tanggal, empId, status wajib diisi');

  const sh      = ensureAbsenHasNama();
  const rows    = sh.getDataRange().getValues();
  const headers = rows[0];

  const tIdx = headers.indexOf('Tanggal');
  const eIdx = headers.indexOf('EmpID');
  const sIdx = headers.indexOf('Status');
  const kIdx = headers.indexOf('Keterangan');
  const nIdx = headers.indexOf('Nama');

  // Normalisasi empId untuk perbandingan
  const empIdNorm = normalizeEmpId(empId);

  // Cari baris existing — bandingkan normalized ID
  let foundRow = -1;
  for (let i = 1; i < rows.length; i++) {
    const rowTgl = fmtDateKey(rows[i][tIdx]);
    const rowEmp = normalizeEmpId(rows[i][eIdx]);
    if (rowTgl === tanggal && rowEmp === empIdNorm) {
      foundRow = i + 1; // 1-based untuk getRange
      break;
    }
  }

  // Resolve nama: pakai yang dikirim, atau lookup dari sheet Karyawan
  const resolvedNama = resolveNama(empId, nama);

  if (foundRow > 0) {
    // UPDATE baris yang ada
    sh.getRange(foundRow, sIdx + 1).setValue(status);
    sh.getRange(foundRow, kIdx + 1).setValue(ket || '');
    if (nIdx >= 0) sh.getRange(foundRow, nIdx + 1).setValue(resolvedNama);
  } else {
    // INSERT baris baru — bangun row sesuai urutan header
    const newRow = headers.map(h => {
      if (h === 'Tanggal')    return tanggal;
      if (h === 'EmpID')      return empId;   // simpan ID asli dari request
      if (h === 'Status')     return status;
      if (h === 'Keterangan') return ket || '';
      if (h === 'Nama')       return resolvedNama;
      return '';
    });
    sh.appendRow(newRow);
  }

  return ok('saved');
}

// Resolve nama dari sheet Karyawan berdasarkan empId
function resolveNama(empId, namaFromRequest) {
  if (namaFromRequest && String(namaFromRequest).trim()) return String(namaFromRequest).trim();
  try {
    const shKar  = getOrCreateSheet(SH_KAR, ['ID','Nama','Jabatan','Tim','Aktif']);
    const rows   = shKar.getDataRange().getValues();
    const headers = rows[0];
    const idIdx  = headers.indexOf('ID');
    const namaIdx = headers.indexOf('Nama');
    const empNorm = normalizeEmpId(empId);
    for (let i = 1; i < rows.length; i++) {
      if (normalizeEmpId(rows[i][idIdx]) === empNorm) return String(rows[i][namaIdx] || '').trim();
    }
  } catch(ex) {}
  return '';
}

// ============================================================
// ACTION: DELETE ABSEN
// ============================================================
function actionDeleteAbsen(data) {
  const { tanggal, empId } = data;
  if (!tanggal || !empId) return err('tanggal dan empId wajib diisi');

  const sh      = ensureAbsenHasNama();
  const rows    = sh.getDataRange().getValues();
  const headers = rows[0];
  const tIdx    = headers.indexOf('Tanggal');
  const eIdx    = headers.indexOf('EmpID');
  const empNorm = normalizeEmpId(empId);

  for (let i = rows.length - 1; i >= 1; i--) {
    const rowTgl = fmtDateKey(rows[i][tIdx]);
    const rowEmp = normalizeEmpId(rows[i][eIdx]);
    if (rowTgl === tanggal && rowEmp === empNorm) {
      sh.deleteRow(i + 1);
      return ok('deleted');
    }
  }
  return err('Data absen tidak ditemukan');
}

// ============================================================
// ACTION: ADD KARYAWAN
// ============================================================
function actionAddKaryawan(data) {
  const { id, nama, jabatan, tim } = data;
  if (!id || !nama || !jabatan || !tim) return err('Semua field karyawan wajib diisi');

  const sh   = getOrCreateSheet(SH_KAR, ['ID','Nama','Jabatan','Tim','Aktif']);
  const rows = sh.getDataRange().getValues();

  const existing = rows.slice(1).find(r => String(r[0]).trim() === String(id).trim());
  if (existing) return err(`ID karyawan "${id}" sudah digunakan`);

  sh.appendRow([id.trim(), nama.trim(), jabatan.trim(), tim.trim(), true]);
  return ok('added');
}

// ============================================================
// ACTION: TOGGLE KARYAWAN
// ============================================================
function actionToggleKaryawan(data) {
  const { id } = data;
  if (!id) return err('ID karyawan wajib diisi');

  const sh      = getOrCreateSheet(SH_KAR, ['ID','Nama','Jabatan','Tim','Aktif']);
  const rows    = sh.getDataRange().getValues();
  const headers = rows[0];
  const idIdx   = headers.indexOf('ID');
  const aktIdx  = headers.indexOf('Aktif');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idIdx]).trim() === String(id).trim()) {
      const current = rows[i][aktIdx];
      const newVal  = !(current === true || String(current).toUpperCase() === 'TRUE');
      sh.getRange(i + 1, aktIdx + 1).setValue(newVal);
      return ok({ toggled: true, aktif: newVal });
    }
  }
  return err('Karyawan tidak ditemukan');
}

// ============================================================
// ACTION: ADD TIM
// ============================================================
function actionAddTim(data) {
  const { nama } = data;
  if (!nama) return err('Nama tim wajib diisi');

  const sh   = getOrCreateSheet(SH_TIM, ['ID','Nama']);
  const rows = sh.getDataRange().getValues();

  const dup = rows.slice(1).find(r =>
    String(r[1]).trim().toLowerCase() === nama.trim().toLowerCase()
  );
  if (dup) return err(`Tim "${nama}" sudah ada`);

  const newId = generateId('TIM');
  sh.appendRow([newId, nama.trim()]);
  return ok({ added: true, id: newId });
}

// ============================================================
// ACTION: DELETE TIM
// ============================================================
function actionDeleteTim(data) {
  const { id } = data;
  if (!id) return err('ID tim wajib diisi');

  const shTim   = getOrCreateSheet(SH_TIM, ['ID','Nama']);
  const rows    = shTim.getDataRange().getValues();
  const headers = rows[0];
  const idIdx   = headers.indexOf('ID');

  const timRow = rows.slice(1).find(r => String(r[idIdx]).trim() === String(id).trim());
  if (!timRow) return err('Tim tidak ditemukan');

  const timNama = timRow[1];
  const shKar   = getOrCreateSheet(SH_KAR, ['ID','Nama','Jabatan','Tim','Aktif']);
  const karRows = shKar.getDataRange().getValues();
  const timIdx  = karRows[0].indexOf('Tim');
  const terpakai = karRows.slice(1).some(r =>
    String(r[timIdx]).trim().toLowerCase() === String(timNama).trim().toLowerCase()
  );
  if (terpakai) return err(`Tim "${timNama}" masih memiliki anggota. Pindahkan dulu karyawannya.`);

  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idIdx]).trim() === String(id).trim()) {
      shTim.deleteRow(i + 1);
      return ok('deleted');
    }
  }
  return err('Tim tidak ditemukan');
}

// ============================================================
// FUNGSI SETUP & TEST (jalankan manual dari GAS editor)
// ============================================================

/**
 * Jalankan SATU KALI dari GAS editor untuk migrasi:
 * menambahkan kolom Nama ke sheet Absen yang sudah ada.
 */
function migrasiTambahKolomNama() {
  ensureAbsenHasNama();
  // Isi kolom Nama dari data Karyawan untuk baris yang sudah ada
  const shAbs  = SS.getSheetByName(SH_ABS);
  const shKar  = SS.getSheetByName(SH_KAR);
  if (!shAbs || !shKar) { Logger.log('Sheet tidak ditemukan'); return; }

  const absData = shAbs.getDataRange().getValues();
  const karData = shKar.getDataRange().getValues();
  const absHeaders = absData[0];
  const karHeaders = karData[0];

  const eIdx = absHeaders.indexOf('EmpID');
  const nIdx = absHeaders.indexOf('Nama');
  const karIdIdx   = karHeaders.indexOf('ID');
  const karNamaIdx = karHeaders.indexOf('Nama');

  if (nIdx < 0) { Logger.log('Kolom Nama tidak ditemukan di sheet Absen'); return; }

  // Build map karyawan: normalizedId → nama
  const karMap = {};
  karData.slice(1).forEach(r => {
    const nid = normalizeEmpId(r[karIdIdx]);
    karMap[nid] = String(r[karNamaIdx] || '').trim();
  });

  let updated = 0;
  for (let i = 1; i < absData.length; i++) {
    const currentNama = String(absData[i][nIdx] || '').trim();
    if (!currentNama) {
      const empNorm = normalizeEmpId(absData[i][eIdx]);
      const nama = karMap[empNorm] || '';
      if (nama) {
        shAbs.getRange(i + 1, nIdx + 1).setValue(nama);
        updated++;
      }
    }
  }
  Logger.log(`Migrasi selesai. ${updated} baris diupdate dengan Nama.`);
}

/**
 * Test doGet dari GAS editor.
 */
function testDoGet() {
  const result = doGet({ parameter: {} });
  Logger.log(result.getContent().substring(0, 500));
}

/**
 * Test saveAbsen dari GAS editor.
 */
function testSaveAbsen() {
  const result = doGet({
    parameter: {
      action: 'saveAbsen',
      data: JSON.stringify({
        tanggal: '2026-06-12',
        empId:   'MS9',
        nama:    'test',
        status:  'Hadir',
        ket:     'test dari GAS editor'
      })
    }
  });
  Logger.log(result.getContent());
}
