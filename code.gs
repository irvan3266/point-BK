/**
 * Aplikasi Poin BK SMA Muhammadiyah Al-Ghifari
 * Backend Google Apps Script
 * Update: Tambah akun Waka Kesiswaan (View Only + Download)
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistem Poin BK - SMAM Al-Ghifari')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // <--- TAMBAHKAN BARIS INI
}
// 1. Setup Database Awal
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheetSiswa = ss.getSheetByName("Data_Siswa");
  if (!sheetSiswa) {
    sheetSiswa = ss.insertSheet("Data_Siswa");
    sheetSiswa.appendRow(["Nama Siswa", "Kelas", "Poin Saat Ini"]);
    sheetSiswa.getRange("A1:C1").setFontWeight("bold").setBackground("#1E3A8A").setFontColor("white");
  }
  
  let sheetRiwayat = ss.getSheetByName("Riwayat_Poin");
  if (!sheetRiwayat) {
    sheetRiwayat = ss.insertSheet("Riwayat_Poin");
    sheetRiwayat.appendRow(["Timestamp", "Guru BK", "Nama Siswa", "Kelas", "Jenis", "Tindakan", "Skor", "Poin Akhir"]);
    sheetRiwayat.getRange("A1:H1").setFontWeight("bold").setBackground("#FFB800");
  }
  return "Database berhasil disiapkan!";
}

// 2. Fungsi Autentikasi Login
// ROLE: "bk" = Guru BK (akses penuh), "waka" = Waka Kesiswaan (view only + download)
function prosesLogin(kodeAkses) {
  kodeAkses = kodeAkses.toLowerCase().trim();
  
  if (kodeAkses === "irvan" || kodeAkses === "irvan123") {
    return { success: true, nama: "Muhammad Irvan, S.Pd.", role: "bk" };
  } else if (kodeAkses === "weni" || kodeAkses === "weny" || kodeAkses === "weny123") {
    return { success: true, nama: "Weny Devitasari, S.Pd.", role: "bk" };
  } else if (kodeAkses === "waka" || kodeAkses === "waka123") {
    // ✅ Akun Waka Kesiswaan — ganti nama & kode sesuai kebutuhan
    return { success: true, nama: "Waka Kesiswaan", role: "waka" };
  } else {
    return { success: false, pesan: "Kode akses tidak valid atau tidak ditemukan!" };
  }
}

// 3. Fungsi Get Semua Siswa
function getAllSiswa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSiswa = ss.getSheetByName("Data_Siswa");
  if (!sheetSiswa) return [];

  const dataSiswa = sheetSiswa.getDataRange().getValues();
  let hasil = [];
  
  for (let i = 1; i < dataSiswa.length; i++) {
    let poin = parseInt(dataSiswa[i][2]);
    let statusSanksi = "Aman";
    
    if (poin <= 0)        statusSanksi = "DIKELUARKAN";
    else if (poin <= 4)   statusSanksi = "Panggilan IV (Kepsek)";
    else if (poin <= 19)  statusSanksi = "Panggilan III (Waka)";
    else if (poin <= 49)  statusSanksi = "Panggilan II (BK)";
    else if (poin <= 80)  statusSanksi = "Panggilan I (Wali Kelas)";

    hasil.push({
      nama  : dataSiswa[i][0],
      kelas : dataSiswa[i][1],
      poin  : poin,
      status: statusSanksi
    });
  }
  return hasil;
}

// 4. Fungsi Get Dashboard Data
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetSiswa   = ss.getSheetByName("Data_Siswa");
  const sheetRiwayat = ss.getSheetByName("Riwayat_Poin");
  
  if (!sheetSiswa || !sheetRiwayat) setupDatabase();

  const dataSiswa   = sheetSiswa.getDataRange().getValues();
  const dataRiwayat = sheetRiwayat.getDataRange().getValues();

  let totalSiswa      = dataSiswa.length > 1 ? dataSiswa.length - 1 : 0;
  let totalPelanggaran = 0;
  let totalPrestasi    = 0;
  let riwayatTerbaru   = [];

  for (let i = 1; i < dataRiwayat.length; i++) {
    if (dataRiwayat[i][4] === "Pelanggaran") totalPelanggaran++;
    if (dataRiwayat[i][4] === "Prestasi")    totalPrestasi++;
  }

  let startIdx = dataRiwayat.length - 1;
  let count    = 0;
  while (startIdx > 0 && count < 5) {
    let row = dataRiwayat[startIdx];
    riwayatTerbaru.push({
      tanggal: Utilities.formatDate(new Date(row[0]), "GMT+7", "dd/MM/yyyy HH:mm"),
      guru  : row[1],
      nama  : row[2],
      kelas : row[3],
      jenis : row[4],
      skor  : row[6]
    });
    startIdx--;
    count++;
  }

  return { totalSiswa, totalPelanggaran, totalPrestasi, riwayat: riwayatTerbaru };
}

// 5. Fungsi Simpan Siswa (Massal & Single)
function simpanSiswaData(listSiswa) {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa = ss.getSheetByName("Data_Siswa");
    const existData  = sheetSiswa.getDataRange().getValues();
    const existSet   = existData.map(row => (row[0] + "_" + row[1]).toLowerCase().trim());

    let countAdded = 0;
    for (let i = 0; i < listSiswa.length; i++) {
      let key = (listSiswa[i].nama + "_" + listSiswa[i].kelas).toLowerCase().trim();
      if (!existSet.includes(key)) {
        sheetSiswa.appendRow([listSiswa[i].nama, listSiswa[i].kelas, 100]);
        existSet.push(key);
        countAdded++;
      }
    }
    return { success: true, pesan: `${countAdded} data siswa berhasil ditambahkan!` };
  } catch(e) {
    return { success: false, pesan: "Error: " + e.toString() };
  }
}

// 6. Fungsi Hapus Siswa Total
function hapusSiswa(nama, kelas) {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa  = ss.getSheetByName("Data_Siswa");
    const sheetRiwayat = ss.getSheetByName("Riwayat_Poin");

    const dataSiswa = sheetSiswa.getDataRange().getValues();
    for (let i = 1; i < dataSiswa.length; i++) {
      if (dataSiswa[i][0].toString().toLowerCase() === nama.toLowerCase() &&
          dataSiswa[i][1].toString().toLowerCase() === kelas.toLowerCase()) {
        sheetSiswa.deleteRow(i + 1);
        break;
      }
    }

    const dataRiwayat = sheetRiwayat.getDataRange().getValues();
    for (let i = dataRiwayat.length - 1; i >= 1; i--) {
      if (dataRiwayat[i][2].toString().toLowerCase() === nama.toLowerCase() &&
          dataRiwayat[i][3].toString().toLowerCase() === kelas.toLowerCase()) {
        sheetRiwayat.deleteRow(i + 1);
      }
    }

    return { success: true, pesan: `Data siswa ${nama} berhasil dihapus permanen.` };
  } catch(e) {
    return { success: false, pesan: "Error: " + e.toString() };
  }
}

// 7. Fungsi Simpan Poin
function simpanDataPoin(data) {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa   = ss.getSheetByName("Data_Siswa");
    const sheetRiwayat = ss.getSheetByName("Riwayat_Poin");
    
    const nama  = data.nama.trim();
    const kelas = data.kelas.trim();
    const skor  = parseInt(data.skor);
    
    const dataSiswa  = sheetSiswa.getDataRange().getValues();
    let barisSiswa   = -1;
    let poinSebelumnya = 100;
    
    for (let i = 1; i < dataSiswa.length; i++) {
      if (dataSiswa[i][0].toString().toLowerCase().trim() === nama.toLowerCase() && 
          dataSiswa[i][1].toString().toLowerCase().trim() === kelas.toLowerCase()) {
        barisSiswa     = i + 1;
        poinSebelumnya = parseInt(dataSiswa[i][2]);
        break;
      }
    }
    
    let poinBaru = poinSebelumnya + skor;
    if (poinBaru > 100) poinBaru = 100;
    if (poinBaru < 0)   poinBaru = 0;
    
    if (barisSiswa > -1) {
      sheetSiswa.getRange(barisSiswa, 3).setValue(poinBaru);
    } else {
      sheetSiswa.appendRow([nama, kelas, poinBaru]);
    }
    
    sheetRiwayat.appendRow([new Date(), data.guru, nama, kelas, data.jenis, data.tindakan, skor, poinBaru]);
    
    return { success: true, poinAkhir: poinBaru, pesan: `Tersimpan! Poin akhir ${nama}: ${poinBaru}` };
  } catch(error) {
    return { success: false, pesan: "Kesalahan: " + error.toString() };
  }
}

// 8. Get Riwayat Detail Siswa
function getRiwayatSiswa(nama, kelas) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRiwayat = ss.getSheetByName("Riwayat_Poin");
  if (!sheetRiwayat) return [];

  const dataRiwayat = sheetRiwayat.getDataRange().getValues();
  let riwayat = [];
  
  for (let i = dataRiwayat.length - 1; i >= 1; i--) {
    if (dataRiwayat[i][2].toString().toLowerCase().trim() === nama.toLowerCase().trim() && 
        dataRiwayat[i][3].toString().toLowerCase().trim() === kelas.toLowerCase().trim()) {
      riwayat.push({
        id      : new Date(dataRiwayat[i][0]).getTime().toString(),
        tanggal : Utilities.formatDate(new Date(dataRiwayat[i][0]), "GMT+7", "dd/MM/yyyy HH:mm"),
        jenis   : dataRiwayat[i][4],
        tindakan: dataRiwayat[i][5],
        skor    : dataRiwayat[i][6],
        guru    : dataRiwayat[i][1]
      });
    }
  }
  return riwayat;
}

// 9. Kalkulasi Ulang Poin Helper
function kalkulasiUlangPoinSiswa(nama, kelas) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRiwayat = ss.getSheetByName("Riwayat_Poin");
  const sheetSiswa   = ss.getSheetByName("Data_Siswa");

  const newDataRiwayat = sheetRiwayat.getDataRange().getValues();
  let historySiswa = [];
  for (let i = 1; i < newDataRiwayat.length; i++) {
    if (newDataRiwayat[i][2].toString().toLowerCase().trim() === nama.toLowerCase().trim() && 
        newDataRiwayat[i][3].toString().toLowerCase().trim() === kelas.toLowerCase().trim()) {
      historySiswa.push({ skor: parseInt(newDataRiwayat[i][6]), rowIdx: i + 1 });
    }
  }

  let calculatedPoint = 100;
  for (let h of historySiswa) {
    calculatedPoint += h.skor;
    if (calculatedPoint > 100) calculatedPoint = 100;
    if (calculatedPoint < 0)   calculatedPoint = 0;
    sheetRiwayat.getRange(h.rowIdx, 8).setValue(calculatedPoint);
  }

  const dataSiswa = sheetSiswa.getDataRange().getValues();
  for (let i = 1; i < dataSiswa.length; i++) {
    if (dataSiswa[i][0].toString().toLowerCase().trim() === nama.toLowerCase().trim() && 
        dataSiswa[i][1].toString().toLowerCase().trim() === kelas.toLowerCase().trim()) {
      sheetSiswa.getRange(i + 1, 3).setValue(calculatedPoint);
      break;
    }
  }
  return calculatedPoint;
}

// 10. Hapus/Batalkan Riwayat
function batalkanRiwayat(timestampStr, nama, kelas) {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRiwayat = ss.getSheetByName("Riwayat_Poin");
    const dataRiwayat  = sheetRiwayat.getDataRange().getValues();
    let rowToDelete    = -1;
    
    for (let i = 1; i < dataRiwayat.length; i++) {
      let ts = new Date(dataRiwayat[i][0]).getTime().toString();
      if (ts === timestampStr && dataRiwayat[i][2] === nama && dataRiwayat[i][3] === kelas) {
        rowToDelete = i + 1;
        break;
      }
    }

    if (rowToDelete > -1) sheetRiwayat.deleteRow(rowToDelete);
    else return { success: false, pesan: "Data riwayat tidak ditemukan!" };

    let pBaru = kalkulasiUlangPoinSiswa(nama, kelas);
    return { success: true, pesan: "Tindakan dibatalkan. Poin dikalkulasi ulang.", poinBaru: pBaru };
  } catch(e) {
    return { success: false, pesan: "Error: " + e.toString() };
  }
}

// 11. Edit Riwayat
function editRiwayatPoin(timestampStr, nama, kelas, newJenis, newTindakan, newSkor) {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRiwayat = ss.getSheetByName("Riwayat_Poin");
    const dataRiwayat  = sheetRiwayat.getDataRange().getValues();
    let rowToEdit      = -1;
    
    for (let i = 1; i < dataRiwayat.length; i++) {
      let ts = new Date(dataRiwayat[i][0]).getTime().toString();
      if (ts === timestampStr && dataRiwayat[i][2] === nama && dataRiwayat[i][3] === kelas) {
        rowToEdit = i + 1;
        break;
      }
    }

    if (rowToEdit > -1) {
      sheetRiwayat.getRange(rowToEdit, 5).setValue(newJenis);
      sheetRiwayat.getRange(rowToEdit, 6).setValue(newTindakan);
      sheetRiwayat.getRange(rowToEdit, 7).setValue(parseInt(newSkor));
    } else {
      return { success: false, pesan: "Data riwayat tidak ditemukan!" };
    }

    let pBaru = kalkulasiUlangPoinSiswa(nama, kelas);
    return { success: true, pesan: "Tindakan berhasil diubah. Poin dikalkulasi ulang.", poinBaru: pBaru };
  } catch(e) {
    return { success: false, pesan: "Error: " + e.toString() };
  }
}

// ============================================================
// 12. ✅ FUNGSI KHUSUS WAKA KESISWAAN — Export Data Siswa
// Hanya mengembalikan data siswa + status sanksi (tanpa riwayat)
// ============================================================
function getExportData() {
  try {
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa  = ss.getSheetByName("Data_Siswa");
    if (!sheetSiswa) return { success: false, pesan: "Sheet Data_Siswa tidak ditemukan!" };

    const dataSiswa = sheetSiswa.getDataRange().getValues();
    let siswaExport = [];

    for (let i = 1; i < dataSiswa.length; i++) {
      let poin   = parseInt(dataSiswa[i][2]);
      let status = "Aman";

      if (poin <= 0)       status = "DIKELUARKAN";
      else if (poin <= 4)  status = "Panggilan IV (Kepsek)";
      else if (poin <= 19) status = "Panggilan III (Waka)";
      else if (poin <= 49) status = "Panggilan II (BK)";
      else if (poin <= 80) status = "Panggilan I (Wali Kelas)";

      siswaExport.push({
        nama  : dataSiswa[i][0],
        kelas : dataSiswa[i][1],
        poin  : poin,
        status: status
      });
    }

    return { success: true, siswa: siswaExport };
  } catch(e) {
    return { success: false, pesan: "Error: " + e.toString() };
  }
}
