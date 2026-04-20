// Ganti dengan ID Spreadsheet Kamu
const SS_ID = "SS_ID";

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function simpanData(data, jenis) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName(jenis);
    const tgl = new Date();
    const no = sheet.getLastRow();
    
    if (jenis === "Absensi") {
      sheet.appendRow([no, tgl, data.nama, data.status, data.ket]);
    } else {
      sheet.appendRow([no, tgl, data.nama, data.pelanggaran, data.ket]);
    }
    return "✅ Data " + jenis + " Berhasil Disimpan!";
  } catch(e) {
    return "❌ Error: " + e.toString();
  }
}

function ambilRekap() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const dataAbs = ss.getSheetByName("Absensi").getDataRange().getValues();
    const dataLap = ss.getSheetByName("Laporan").getDataRange().getValues();
    let rekap = {};

    for (let i = 1; i < dataAbs.length; i++) {
      let nama = dataAbs[i][2];
      let status = dataAbs[i][3];
      if (!nama) continue;
      if (!rekap[nama]) rekap[nama] = { izin: 0, sakit: 0, alpha: 0, laporan: 0 };
      if (status == "Izin") rekap[nama].izin++;
      if (status == "Sakit") rekap[nama].sakit++;
      if (status == "Alpha") rekap[nama].alpha++;
    }

    for (let j = 1; j < dataLap.length; j++) {
      let nama = dataLap[j][2];
      if (!nama) continue;
      if (!rekap[nama]) rekap[nama] = { izin: 0, sakit: 0, alpha: 0, laporan: 0 };
      rekap[nama].laporan++;
    }
    return rekap;
  } catch(e) {
    return { error: e.toString() };
  }
}
