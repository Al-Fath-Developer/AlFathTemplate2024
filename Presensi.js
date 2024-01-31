// ======= Template Presensi
// Mendapatkan spreadsheet dan sheet rekapan
// ini yang diubah
const SPREADSHEET_ID = "1mvaGSfS-WlA1noFGG_pe1SskxWDBnXiT3lb4pd4ub08";// ini id spreadsheet nya
const SHEET_NAME = "General";// ini nama sheet nya
const LOG = "Log"






//kalau make spreadsheet yang sesuai template, tidak perlu ubah kodingan dibawah ini

const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
const sheet = spreadsheet.getSheetByName(SHEET_NAME);
const log = spreadsheet.getSheetByName(LOG);

// Fungsi untuk merekap data secara otomatis

function temp(){
  // autoRekapData({values: ["01/02/2024", "13"]})
  Logger.log(cariPosisiWaktu("01/02/2024"))

}

function getLog(inputan){
  const timestamp= inputan[0]
  const nim = inputan[1].trim();
  
  const lr  = log.appendRow([timestamp, nim, false, "","Hadir"]).getLastRow()
  log.getRange("D"+lr).setValue(`=IF(C${lr}=TRUE;IF(D${lr}<>"";D${lr};now());"")`)
  log.getRange("C"+lr).insertCheckboxes()


}
function autoRekapData(e) {
  try {
    const inputan = e.values;
    getLog(inputan)
    
    // Mendapatkan waktu dan NIM dari inputan
    const waktu = inputan[0];
    const nim = inputan[1].trim();
  
    // Mencari posisi NIM dan waktu
    let posisiNIM = cariPosisiNIM(nim);
    const posisiWaktu = cariPosisiWaktu(waktu);
    
    // Jika NIM tidak ditemukan, tambahkan NIM baru ke sheet
    if (posisiNIM < 0) {
      posisiNIM = sheet.appendRow([nim]).getLastRow();
       
    }
    if (posisiWaktu>=0){
     sheet.getRange(posisiNIM, posisiWaktu).setValue("Hadir");
    }else{
        //kalo waktunya ga nemu, ditaro di kolom 2
       sheet.getRange(posisiNIM, 2).setValue(sheet.getRange(posisiNIM, 2).getValue().toString() + "H")
    }

  } catch (error) {
    Logger.log("Terjadi kesalahan: " + error.message);
  }
}

// Fungsi untuk mencari posisi NIM dalam sheet
function cariPosisiNIM(nim){
  // searching baris
  try {
    const lastRow = sheet.getLastRow();
  const list_nim = sheet.getRange(1,1,lastRow).getValues()
   const posisi =  list_nim.findIndex((nim_sheet)=>nim_sheet[0].toString() == nim)
  return posisi > 0? posisi + 1: posisi

  } catch (error) {
    throw new Error("Error pada fungsi cariPosisiNIM: " + error.message);
  }
}

// Fungsi untuk mengonversi format datetime
function konversiDatetime(inputDatetime) {
  try {
    // Menggunakan objek Date untuk mengonversi datetime
    return inputDatetime.slice(0,10);
  } catch (error) {
    throw new Error("Error pada fungsi konversiDatetime: " + error.message);
  }
}

// Fungsi untuk mencari posisi waktu dalam sheet
function cariPosisiWaktu(waktu){
  // searching kolom
  try {
    waktu = konversiDatetime(waktu);
    const lastColumn = sheet.getLastColumn();
    const list_waktu = sheet.getRange(2,1,1,lastColumn).getValues()[0]
    const posisi = list_waktu.findIndex((waktu_sheet)=>waktu_sheet.toLocaleString('id-ID', {day: '2-digit', month: '2-digit', year:'numeric'}) == waktu)
    return posisi>0 ? posisi + 1 : posisi
    
  
  
  } catch (error) {
    throw new Error("Error pada fungsi cariPosisiWaktu: " + error.message);
  }
}
