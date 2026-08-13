// 1. Fungsi MEMBACA data dari Google Sheets (GET)
function doGet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var rows = data.slice(1); 
  
  var result = rows.map(function(row, index) {
    return {
      rowIndex: index + 2, // Posisi nomor baris di Google Sheets
      namaAlat: row[0],
      merek: row[1],
      kategori: row[2],
      jumlah: row[3],
      spesifikasi: row[4],
      kondisi: row[5] || "Baik"
    };
  });
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 2. Fungsi MENAMBAH & MENGUPDATE data ke Google Sheets (POST)
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    // FITUR A: UPDATE KONDISI LANGSUNG DARI DROPDOWN TABEL
    if (data.action === "updateKondisi") {
      sheet.getRange(data.rowIndex, 6).setValue(data.kondisi); // Ubah nilai Kolom F (Kondisi)
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Kondisi berhasil diperbarui!" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // FITUR B: TAMBAH BARANG BARU
    var nextRow = sheet.getLastRow() + 1;
    sheet.appendRow([
      data.namaAlat,
      data.merek,
      data.kategori,
      data.jumlah,
      data.spesifikasi,
      data.kondisi || "Baik"
    ]);
    
    // Validasi Dropdown di Google Sheets
    var cellKondisi = sheet.getRange(nextRow, 6);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Baik', 'Rusak Ringan', 'Rusak Parah'], true)
      .setAllowInvalid(false)
      .build();
    cellKondisi.setDataValidation(rule);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}