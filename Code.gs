/*
 * USER CONTROL CENTER - MANAGEMENT USER
 * Koneksi DATABASE_USER dari 2 aplikasi
 * APP 1 = Ekspedisi Material
 * APP 2 = Monitoring Material
 */

const CONFIG = {
  APP1: {
    ID: '1tdL-FHRYywfo-GDY-eIWUTpXpT156AWO4FH_iXjXUAM',
    NAME: 'EKSPEDISI MATERIAL'
  },
  APP2: {
    ID: '1OZbZ1FvCIXXQHxba2mQ90_NYuKJe5vgDZRcv4hwr3Ec',
    NAME: 'MONITORING MATERIAL'
  },
  SOURCE_SHEET: 'DATABASE_USER',
  TARGET_SHEET: 'DATABASE_USER_CONTROL'
};

function sinkronDatabaseUser() {
  const ssControl = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ssControl.getSheetByName(CONFIG.TARGET_SHEET);

  if (!targetSheet) {
    throw new Error('Sheet ' + CONFIG.TARGET_SHEET + ' tidak ditemukan.');
  }

  const dataApp1 = bacaDatabaseUser(CONFIG.APP1.ID, CONFIG.APP1.NAME);
  const dataApp2 = bacaDatabaseUser(CONFIG.APP2.ID, CONFIG.APP2.NAME);
  const semuaData = dataApp1.concat(dataApp2);

  const header = [
    'APLIKASI',
    'ID DEVICE',
    'NAMA',
    'STATUS',
    'TANGGAL DAFTAR',
    'LAST ONLINE',
    'PERANGKAT'
  ];

  const maxRows = targetSheet.getMaxRows();
  const maxCols = targetSheet.getMaxColumns();
  if (maxRows > 1) {
    targetSheet.getRange(2, 1, maxRows - 1, Math.min(maxCols, 7)).clearContent();
  }

  targetSheet.getRange(1, 1, 1, header.length).setValues([header]);

  if (semuaData.length > 0) {
    targetSheet.getRange(2, 1, semuaData.length, header.length).setValues(semuaData);
  }

  targetSheet.getRange(1, 1, 1, header.length)
    .setFontWeight('bold')
    .setBackground('#45B8C0')
    .setFontColor('#FFFFFF');

  targetSheet.autoResizeColumns(1, header.length);

  console.log('Sinkronisasi selesai. Total user ' + semuaData.length);
}

function bacaDatabaseUser(spreadsheetId, namaAplikasi) {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);

  if (!sheet) {
    throw new Error(
      'Sheet ' + CONFIG.SOURCE_SHEET + ' tidak ditemukan pada aplikasi ' + namaAplikasi + '.'
    );
  }

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 7);
  if (lastRow < 2) return [];

  // Mendukung 2 format DATABASE_USER:
  // FORMAT LAMA : A ID DEVICE | B NAMA | C STATUS | D TANGGAL DAFTAR | E LAST ONLINE | F PERANGKAT
  // FORMAT BARU : A APLIKASI | B ID DEVICE | C NAMA | D STATUS | E TANGGAL DAFTAR | F LAST ONLINE | G PERANGKAT
  const raw = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = raw[0].map(function(v) { return String(v || '').trim().toUpperCase(); });

  const idx = function(names, fallback) {
    for (let i = 0; i < names.length; i++) {
      const pos = headers.indexOf(names[i]);
      if (pos !== -1) return pos;
    }
    return fallback;
  };

  const colAplikasi = idx(['APLIKASI', 'APPLICATION'], -1);
  const colId = idx(['ID DEVICE', 'ID_DEVICE', 'DEVICE ID', 'DEVICE_ID'], colAplikasi === -1 ? 0 : 1);
  const colNama = idx(['NAMA', 'NAMA USER', 'NAMA USER'], colAplikasi === -1 ? 1 : 2);
  const colStatus = idx(['STATUS'], colAplikasi === -1 ? 2 : 3);
  const colTanggalDaftar = idx(['TANGGAL DAFTAR', 'TANGGAL_DAFTAR', 'TANGGAL REGISTRASI'], colAplikasi === -1 ? 3 : 4);
  const colLastOnline = idx(['LAST ONLINE', 'LAST_ONLINE'], colAplikasi === -1 ? 4 : 5);
  const colPerangkat = idx(['PERANGKAT', 'DEVICE', 'PLATFORM'], colAplikasi === -1 ? 5 : 6);

  const hasil = [];

  for (let r = 1; r < raw.length; r++) {
    const row = raw[r];
    const idDevice = String(row[colId] || '').trim();
    const nama = String(row[colNama] || '').trim();
    const status = String(row[colStatus] || '').trim();
    const tanggalDaftar = String(row[colTanggalDaftar] || '').trim();
    const lastOnline = String(row[colLastOnline] || '').trim();
    const perangkat = String(row[colPerangkat] || '').trim();

    if (!idDevice && !nama) continue;

    // Jika sumber memiliki kolom APLIKASI, tetap gunakan nama aplikasi dari konfigurasi
    // agar hasil sinkronisasi selalu konsisten dengan 2 aplikasi yang terdaftar.
    hasil.push([
      namaAplikasi,
      idDevice,
      nama,
      status,
      tanggalDaftar,
      lastOnline,
      perangkat
    ]);
  }

  return hasil;
}

function buatTriggerSinkronUser() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sinkronDatabaseUser') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('sinkronDatabaseUser')
    .timeBased()
    .everyMinutes(1)
    .create();

  console.log('Trigger sinkronisasi user dibuat.');
}

function ubahStatusUser(aplikasi, idDevice, statusBaru) {
  if (!aplikasi) throw new Error('APLIKASI tidak boleh kosong.');
  if (!idDevice) throw new Error('ID DEVICE tidak boleh kosong.');
  if (!statusBaru) throw new Error('STATUS tidak boleh kosong.');

  statusBaru = String(statusBaru).trim().toUpperCase();
  if (statusBaru !== 'APPROVED' && statusBaru !== 'REJECTED') {
    throw new Error('Status hanya boleh APPROVED atau REJECTED.');
  }

  let appConfig = null;
  const namaApp = String(aplikasi).trim().toUpperCase();

  if (namaApp === CONFIG.APP1.NAME) {
    appConfig = CONFIG.APP1;
  } else if (namaApp === CONFIG.APP2.NAME) {
    appConfig = CONFIG.APP2;
  } else {
    throw new Error('Aplikasi tidak dikenal: ' + aplikasi);
  }

  const ss = SpreadsheetApp.openById(appConfig.ID);
  const sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!sheet) {
    throw new Error('Sheet DATABASE_USER tidak ditemukan pada ' + appConfig.NAME);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('DATABASE_USER kosong pada ' + appConfig.NAME);
  }

  const lastCol = Math.max(sheet.getLastColumn(), 7);
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(function(v) { return String(v || '').trim().toUpperCase(); });

  const findCol = function(names, fallback) {
    for (let i = 0; i < names.length; i++) {
      const p = headers.indexOf(names[i]);
      if (p !== -1) return p;
    }
    return fallback;
  };

  const hasNewFormat = headers.indexOf('ID DEVICE') !== -1 || headers.indexOf('APLIKASI') !== -1;
  const colId = findCol(['ID DEVICE', 'ID_DEVICE', 'DEVICE ID', 'DEVICE_ID'], hasNewFormat ? 1 : 0);
  const colNama = findCol(['NAMA', 'NAMA USER'], hasNewFormat ? 2 : 1);
  const colStatus = findCol(['STATUS'], hasNewFormat ? 3 : 2);

  let ditemukan = false;
  let namaUser = '';
  let nomorBaris = 0;

  for (let i = 1; i < values.length; i++) {
    const id = String(values[i][colId] || '').trim();
    if (id === String(idDevice).trim()) {
      nomorBaris = i + 1;
      namaUser = values[i][colNama] || '';
      ditemukan = true;
      break;
    }
  }

  if (!ditemukan) {
    throw new Error(
      'ID DEVICE tidak ditemukan pada ' + appConfig.NAME + ': ' + idDevice
    );
  }

  // Update HANYA kolom STATUS.
  // TANGGAL DAFTAR dan LAST ONLINE tidak disentuh agar nilainya tetap terjaga.
  sheet.getRange(nomorBaris, colStatus + 1).setValue(statusBaru);
  SpreadsheetApp.flush();
  sinkronDatabaseUser();

  return {
    success: true,
    aplikasi: appConfig.NAME,
    idDevice: idDevice,
    nama: namaUser,
    status: statusBaru
  };
}

function approveUser(aplikasi, idDevice) {
  return ubahStatusUser(aplikasi, idDevice, 'APPROVED');
}

function rejectUser(aplikasi, idDevice) {
  return ubahStatusUser(aplikasi, idDevice, 'REJECTED');
}

function testKoneksiUser() {
  const dataApp1 = bacaDatabaseUser(CONFIG.APP1.ID, CONFIG.APP1.NAME);
  const dataApp2 = bacaDatabaseUser(CONFIG.APP2.ID, CONFIG.APP2.NAME);
  console.log('APP 1: ' + dataApp1.length + ' user');
  console.log('APP 2: ' + dataApp2.length + ' user');
  console.log('TOTAL: ' + (dataApp1.length + dataApp2.length) + ' user');
}

function testCariUser() {
  const aplikasi = CONFIG.APP1.NAME;
  const idDevice = 'MASUKKAN_ID_DEVICE_PENDING_DI_SINI';
  const ss = SpreadsheetApp.openById(CONFIG.APP1.ID);
  const sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(idDevice).trim()) {
      console.log('USER DITEMUKAN');
      console.log('Nama: ' + data[i][1]);
      console.log('Status: ' + data[i][2]);
      console.log('Perangkat: ' + data[i][5]);
      return;
    }
  }
  console.log('USER TIDAK DITEMUKAN');
}

function cekStatusUser() {
  const apps = [CONFIG.APP1, CONFIG.APP2];
  apps.forEach(function(app) {
    const ss = SpreadsheetApp.openById(app.ID);
    const sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
    if (!sheet) throw new Error('DATABASE_USER tidak ditemukan ' + app.NAME);

    const lastRow = sheet.getLastRow();
    console.log(app.NAME + ' - JUMLAH BARIS: ' + lastRow);
    if (lastRow < 2) {
      console.log(app.NAME + ' - BELUM ADA USER');
      return;
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 6).getDisplayValues();
    let approved = 0;
    let pending = 0;
    let rejected = 0;

    data.forEach(function(row) {
      const status = String(row[2]).trim().toUpperCase();
      if (status === 'APPROVED') approved++;
      if (status === 'PENDING') pending++;
      if (status === 'REJECTED') rejected++;
    });

    console.log(
      app.NAME +
      ' - APPROVED: ' + approved +
      ' | PENDING: ' + pending +
      ' | REJECTED: ' + rejected
    );
  });

  console.log('CEK STATUS SELESAI');
}

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('USER CONTROL CENTER')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getControlCenterData() {
  // Sinkronkan sumber terlebih dahulu supaya STATUS/LAST ONLINE terbaru terbaca.
  sinkronDatabaseUser();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
  if (!sheet) throw new Error('Sheet DATABASE_USER_CONTROL tidak ditemukan.');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { active: [], pending: [] };

  const data = sheet.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
  const active = [];
  const pending = [];

  data.forEach(function(row) {
    const aplikasi = String(row[0] || '').trim();
    const idDevice = String(row[1] || '').trim();
    const nama = String(row[2] || '').trim();
    const status = String(row[3] || '').trim().toUpperCase();
    const tanggalDaftar = String(row[4] || '').trim();
    const lastOnline = String(row[5] || '').trim();
    const perangkat = String(row[6] || '').trim();

    if (!idDevice && !nama) return;

    if (
      status === 'PENDING' ||
      status === 'MENUNGGU APPROVAL' ||
      status === ''
    ) {
      pending.push({
        aplikasi: aplikasi,
        idDevice: idDevice,
        nama: nama,
        perangkat: perangkat,
        tanggalDaftar: tanggalDaftar,
        status: status
      });
    }

    const onlineInfo = hitungOnline(lastOnline);
    if (status === 'APPROVED' && onlineInfo.online) {
      active.push({
        aplikasi: aplikasi,
        idDevice: idDevice,
        nama: nama,
        perangkat: perangkat,
        lastOnline: lastOnline,
        durasi: onlineInfo.durasi
      });
    }
  });

  return { active: active, pending: pending };
}

/*
 * PERBAIKAN ONLINE:
 * Menerima Date asli Google Sheets dan string seperti:
 *   Sabtu, 03 Oktober 2026 — 19:13:36
 *   Sabtu, 03 Oktober 2026 19:13:36
 *   03 Oktober 2026 19:13:36
 *   03/10/2026 19:13:36
 *   2026-10-03T19:13:36
 */
function hitungOnline(lastOnline) {
  if (lastOnline === null || lastOnline === undefined || String(lastOnline).trim() === '') {
    return { online: false, durasi: '-' };
  }

  const date = parseTanggalIndonesia(lastOnline);
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    return { online: false, durasi: '-' };
  }

  const sekarang = new Date();
  let selisih = Math.floor((sekarang.getTime() - date.getTime()) / 1000);

  // Toleransi kecil untuk perbedaan waktu/penulisan timestamp.
  // User dianggap online jika LAST ONLINE <= 3 menit.
  const online = selisih >= 0 && selisih <= 180;

  // Jika timestamp sedikit di masa depan karena perbedaan clock, tetap tampil online.
  if (selisih < 0 && selisih >= -30) {
    selisih = 0;
  }

  let detik = Math.max(0, selisih);
  const jam = Math.floor(detik / 3600);
  detik %= 3600;
  const menit = Math.floor(detik / 60);
  detik %= 60;

  const durasi =
    String(jam).padStart(2, '0') + ':' +
    String(menit).padStart(2, '0') + ':' +
    String(detik).padStart(2, '0');

  return { online: online, durasi: durasi };
}

function parseTanggalIndonesia(value) {
  if (value === null || value === undefined || value === '') return null;

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return isNaN(value.getTime()) ? null : value;
  }

  const textOriginal = String(value).trim();
  if (!textOriginal) return null;

  // Coba format Date/ISO terlebih dahulu.
  const isoTry = new Date(textOriginal);
  if (!isNaN(isoTry.getTime()) && /\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(textOriginal)) {
    return isoTry;
  }

  // Buang nama hari dan normalisasi dash/whitespace.
  let text = textOriginal
    .replace(/^[^,]+,\s*/i, '')
    .replace(/[—–-]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const bulan = {
    januari: 0,
    februari: 1,
    maret: 2,
    april: 3,
    mei: 4,
    juni: 5,
    juli: 6,
    agustus: 7,
    september: 8,
    oktober: 9,
    november: 10,
    desember: 11
  };

  // dd/mm/yyyy [hh:mm[:ss]]
  let match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    const hari = Number(match[1]);
    const bulanIndex = Number(match[2]) - 1;
    const tahun = Number(match[3]);
    const jam = Number(match[4] || 0);
    const menit = Number(match[5] || 0);
    const detik = Number(match[6] || 0);
    const result = new Date(tahun, bulanIndex, hari, jam, menit, detik);
    return isValidDateParts(result, tahun, bulanIndex, hari, jam, menit, detik) ? result : null;
  }

  // dd NamaBulan yyyy [hh:mm[:ss]]
  match = text.match(/^([0-9]{1,2})\s+([A-Za-z]+)\s+([0-9]{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (match) {
    const hari = Number(match[1]);
    const namaBulan = String(match[2]).toLowerCase();
    const tahun = Number(match[3]);
    const jam = Number(match[4] || 0);
    const menit = Number(match[5] || 0);
    const detik = Number(match[6] || 0);

    if (bulan[namaBulan] === undefined) return null;

    const bulanIndex = bulan[namaBulan];
    const result = new Date(tahun, bulanIndex, hari, jam, menit, detik);
    return isValidDateParts(result, tahun, bulanIndex, hari, jam, menit, detik) ? result : null;
  }

  return null;
}

function isValidDateParts(date, tahun, bulan, hari, jam, menit, detik) {
  return date instanceof Date &&
    !isNaN(date.getTime()) &&
    date.getFullYear() === tahun &&
    date.getMonth() === bulan &&
    date.getDate() === hari &&
    date.getHours() === jam &&
    date.getMinutes() === menit &&
    date.getSeconds() === detik;
}

function testOnlineParser() {
  const contoh = [
    'Sabtu, 03 Oktober 2026 — 19:13:36',
    'Sabtu, 03 Oktober 2026 19:13:36',
    '03 Oktober 2026 19:13:36',
    '03/10/2026 19:13:36',
    '2026-10-03T19:13:36'
  ];

  contoh.forEach(function(item) {
    const d = parseTanggalIndonesia(item);
    console.log(item + ' => ' + (d ? d.toString() : 'GAGAL'));
  });
}
