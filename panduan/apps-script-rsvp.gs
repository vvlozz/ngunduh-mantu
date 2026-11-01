/**
 * Apps Script untuk RSVP / Buku Tamu undangan Ngunduh Mantu.
 * Tempel di Google Sheet baru: Extensions > Apps Script.
 * Kolom di sheet (baris 1 = header): Waktu | Nama | Jumlah Tamu | Kehadiran | Ucapan
 */
const SHEET_NAME = "RSVP";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Waktu", "Nama", "Jumlah Tamu", "Kehadiran", "Ucapan"]);
  }
  return sh;
}

// Dipanggil saat tamu mengirim form
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.website) return json_({ status: "ok" }); // jebakan anti-bot
    const nama = String(d.nama || "").trim().slice(0, 100);
    if (!nama) return json_({ status: "error", message: "Nama kosong" });
    getSheet_().appendRow([
      new Date(),
      nama,
      String(d.jumlahTamu || "").slice(0, 20),
      String(d.kehadiran || "").slice(0, 30),
      String(d.ucapan || "").slice(0, 500),
    ]);
    return json_({ status: "ok" });
  } catch (err) {
    return json_({ status: "error", message: String(err) });
  }
}

// Dipanggil website untuk menampilkan Buku Tamu
function doGet() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return json_({ status: "ok", data: [] });
  const rows = sh.getRange(2, 1, last - 1, 5).getValues();
  const data = rows.map((r, i) => ({
    id: i + 2, // nomor baris sebagai ID
    nama: r[1],
    jumlahTamu: r[2],
    kehadiran: r[3],
    ucapan: r[4],
  }));
  return json_({ status: "ok", data: data });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}
