// ==========================================
// BACKEND SIREKSI - GOOGLE APPS SCRIPT (FIXED DATA TYPE)
// ==========================================

function ensureSheetExists(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f4f6");
  }
  return sheet;
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SIREKSI - Rekap Skripsi Otomatis')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// SOLUSI PAMUNGKAS: Menggunakan getDisplayValues() 
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if(!sheet) return [];
  
  // Mengubah semua Tanggal dan Angka menjadi Teks String agar tidak crash saat dikirim ke web
  const data = sheet.getDataRange().getDisplayValues();
  
  if(data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, index) => { 
        obj[header.toString().trim()] = row[index]; 
    });
    return obj;
  });
}

function calculateStatus(mhs) {
  let status = "Belum Mengajukan Judul";
  let progress = 0;
  let reminder = "Aman";
  let today = new Date();

  let tPengajuan = mhs.Tgl_Pengajuan ? new Date(mhs.Tgl_Pengajuan) : null;
  let tSempro = mhs.Tgl_Sempro ? new Date(mhs.Tgl_Sempro) : null;
  let tSemhas = mhs.Tgl_Semhas ? new Date(mhs.Tgl_Semhas) : null;
  let tPDD = mhs.Tgl_PDD ? new Date(mhs.Tgl_PDD) : null;

  if (tPDD && !isNaN(tPDD)) {
    status = "Selesai Pendadaran"; progress = 100;
  } else if (tSemhas && !isNaN(tSemhas)) {
    status = "Menuju PDD"; progress = 75;
    let diff = Math.floor((today - tSemhas) / (1000 * 60 * 60 * 24));
    if (diff > 45) reminder = "Segera dijadwalkan";
  } else if (tSempro && !isNaN(tSempro)) {
    status = "Menuju Semhas"; progress = 50;
    let diff = Math.floor((today - tSempro) / (1000 * 60 * 60 * 24));
    if (diff > 90) reminder = "Perlu follow up";
  } else if (tPengajuan && !isNaN(tPengajuan)) {
    status = "Menunggu Sempro"; progress = 25;
    let diff = Math.floor((today - tPengajuan) / (1000 * 60 * 60 * 24));
    if (diff > 60) reminder = "Perlu dipantau";
  }
  return { status, progress, reminder };
}

// === CRUD DOSEN ===
function getDosen() {
  try {
    let dosenList = getSheetData("Data_Dosen");
    let mhsList = getSheetData("Data_Mahasiswa");
    dosenList.forEach(d => {
      d.Total_Bimbingan = mhsList.filter(m => m.ID_Dosen == d.ID_Dosen).length;
    });
    return dosenList;
  } catch (err) { throw new Error(err.message); }
}

function saveDosen(form) {
  try {
    const headers = ["ID_Dosen", "Nama_Dosen", "NIP", "Prodi", "Kuota"];
    const sheet = ensureSheetExists("Data_Dosen", headers);
    
    if (form.ID_Dosen && form.ID_Dosen !== "") {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == form.ID_Dosen) {
          sheet.getRange(i + 1, 2).setValue(form.Nama_Dosen);
          sheet.getRange(i + 1, 3).setValue(form.NIP);
          sheet.getRange(i + 1, 4).setValue(form.Prodi);
          sheet.getRange(i + 1, 5).setValue(form.Kuota);
          
          let sheetMhs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_Mahasiswa");
          if (sheetMhs) {
            let mhsData = sheetMhs.getDataRange().getValues();
            for (let j = 1; j < mhsData.length; j++) {
              if (mhsData[j][4] == form.ID_Dosen) sheetMhs.getRange(j + 1, 6).setValue(form.Nama_Dosen);
            }
          }
          return "Berhasil diupdate";
        }
      }
    } else {
      let newID = "DOS-" + new Date().getTime();
      sheet.appendRow([newID, form.Nama_Dosen, form.NIP, form.Prodi, form.Kuota]);
      return "Berhasil ditambahkan";
    }
  } catch (error) { throw new Error("Gagal menyimpan Dosen: " + error.message); }
}

function deleteDosen(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_Dosen");
    if(!sheet) return "Data tidak ditemukan";
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return "Berhasil dihapus";
      }
    }
  } catch (error) { throw new Error("Gagal menghapus: " + error.message); }
}

// === CRUD MAHASISWA ===
function getMahasiswa() { return getSheetData("Data_Mahasiswa"); }

function saveMahasiswa(form) {
  try {
    const headers = ["ID_Mahasiswa", "Nama", "NIM", "Angkatan", "ID_Dosen", "Nama_Dosen", "Judul", "Tgl_Pengajuan", "Tgl_Sempro", "Tgl_Semhas", "Tgl_PDD", "Status", "Progress", "Reminder"];
    const sheet = ensureSheetExists("Data_Mahasiswa", headers);
    
    let calc = calculateStatus(form);
    let dosenList = getSheetData("Data_Dosen");
    let namaDosen = dosenList.find(d => d.ID_Dosen == form.ID_Dosen)?.Nama_Dosen || "-";

    let rowData = [
      form.ID_Mahasiswa || "MHS-" + new Date().getTime(),
      form.Nama, form.NIM, form.Angkatan, form.ID_Dosen, namaDosen, form.Judul,
      form.Tgl_Pengajuan, form.Tgl_Sempro, form.Tgl_Semhas, form.Tgl_PDD,
      calc.status, calc.progress, calc.reminder
    ];

    if (form.ID_Mahasiswa && form.ID_Mahasiswa !== "") {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == form.ID_Mahasiswa) {
          sheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
          return "Berhasil diupdate";
        }
      }
    } else {
      sheet.appendRow(rowData);
      return "Berhasil ditambahkan";
    }
  } catch (error) { throw new Error("Gagal menyimpan Mahasiswa: " + error.message); }
}

function deleteMahasiswa(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data_Mahasiswa");
    if(!sheet) return "Data tidak ditemukan";
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return "Berhasil dihapus";
      }
    }
  } catch (error) { throw new Error("Gagal menghapus: " + error.message); }
}

// === DASHBOARD ===
function getDashboardData() {
  let mhs = getSheetData("Data_Mahasiswa");
  let dosen = getDosen();
  let stats = { total: mhs.length, pengajuan: 0, sempro: 0, semhas: 0, pdd: 0, belum: 0 };

  mhs.forEach(m => {
    if (m.Status === "Belum Mengajukan Judul") stats.belum++;
    else if (m.Status === "Menunggu Sempro") stats.pengajuan++;
    else if (m.Status === "Menuju Semhas") stats.sempro++;
    else if (m.Status === "Menuju PDD") stats.semhas++;
    else if (m.Status === "Selesai Pendadaran") stats.pdd++;
  });
  return { stats, dosen };
}
// ==========================================
// TAMBAHAN API ENDPOINT UNTUK VERCEL (FIXED)
// ==========================================
function doGet(e) {
  // Cek apakah ada parameter 'action' dari fetch Vercel
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var outputData = "";
    
    if (action === "getDosen") {
      outputData = JSON.stringify(getDosen());
    } else if (action === "getMahasiswa") {
      outputData = JSON.stringify(getMahasiswa());
    } else if (action === "getDashboard") {
      outputData = JSON.stringify(getDashboardData());
    }
    
    return ContentService.createTextOutput(outputData)
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Default: Merender tampilan web app jika dibuka langsung dari link GAS
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SIREKSI - Rekap Skripsi Otomatis')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
