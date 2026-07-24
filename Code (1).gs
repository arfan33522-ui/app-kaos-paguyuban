/**
 * BACKEND API untuk App Data Order & Pembayaran Kaos Paguyuban.
 * Disesuaikan dengan struktur sheet "Daftar Order":
 * Kolom: B=No, C=Nama, D=Ukuran, E=Jenis Kaos, F=Status Bayar,
 *        G=Harga per pcs, H=Tagihan, I=Keterangan. Data mulai baris 3.
 *
 * CARA PASANG:
 * 1. Buka Google Sheets Abang.
 * 2. Menu Extensions -> Apps Script.
 * 3. Hapus semua isi editor, paste seluruh isi file ini.
 * 4. Deploy -> Manage deployments -> edit (pensil) -> New version -> Deploy.
 */

const SHEET_NAME = 'Daftar Order';
const SPREADSHEET_ID = '14dZ9HL_euRAwD_WV8wlx1_XxZpffOigMWD8YTISsTz0';
const START_ROW = 3; // data pertama mulai di baris ini
const START_COL = 2; // kolom B

function getSheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
}

function normalizeUkuran_(u) {
  return u === 'XXXL' ? '3XL' : u;
}
function denormalizeUkuran_(u) {
  return u === '3XL' ? 'XXXL' : u;
}
function normalizeLengan_(j) {
  return String(j || '').replace('Lengan ', '');
}
function denormalizeLengan_(l) {
  return 'Lengan ' + l;
}

function doGet(e) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) {
    return jsonOutput({ data: [] });
  }
  const numRows = lastRow - START_ROW + 1;
  const values = sheet.getRange(START_ROW, START_COL, numRows, 8).getValues();
  // urutan kolom yang diambil: No, Nama, Ukuran, JenisKaos, StatusBayar, HargaPerPcs, Tagihan, Keterangan

  const data = values
    .map(function (r, i) {
      return { r: r, rowIndex: START_ROW + i };
    })
    .filter(function (o) { return o.r[1]; }) // hanya baris yang ada nama-nya
    .map(function (o) {
      const r = o.r;
      return {
        row: o.rowIndex,
        no: r[0],
        nama: r[1],
        hp: '',
        ukuran: normalizeUkuran_(r[2]),
        lengan: normalizeLengan_(r[3]),
        harga: r[6] || r[5] || 0, // pakai Tagihan, fallback Harga per pcs
        status: r[4] || 'Belum Bayar',
        tanggal: '',
        keterangan: r[7] || '',
      };
    });
  return jsonOutput({ data: data });
}

function doPost(e) {
  const sheet = getSheet_();
  const body = JSON.parse(e.postData.contents);

  if (body.action === 'updateStatus') {
    sheet.getRange(body.row, 6).setValue(body.status); // kolom F = Status Bayar
    return jsonOutput({ ok: true });
  }

  if (body.action === 'addPeserta') {
    const lastRow = sheet.getLastRow();
    const newRowNum = lastRow + 1;
    const newNo = lastRow - START_ROW + 2;
    sheet.getRange(newRowNum, START_COL, 1, 8).setValues([[
      newNo,
      body.nama,
      denormalizeUkuran_(body.ukuran),
      denormalizeLengan_(body.lengan),
      body.status || 'Belum Bayar',
      body.harga,
      body.harga,
      body.keterangan || '',
    ]]);
    return jsonOutput({ ok: true, row: newRowNum });
  }

  return jsonOutput({ ok: false, error: 'Aksi tidak dikenali' });
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
