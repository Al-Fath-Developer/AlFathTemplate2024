// =====Template Perizinan

// Mendapatkan spreadsheet dan sheet
// ini yang diubah
const SPREADSHEET_ID = "1mvaGSfS-WlA1noFGG_pe1SskxWDBnXiT3lb4pd4ub08"; // ini id spreadsheet nya, diganti aja
const SHEET_NAME = "General"; // ini nama sheet nya



//kalau make spreadsheet yang sesuai template, tidak perlu ubah kodingan dibawah ini

// Fungsi untuk memproses data rekap otomatis
const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
const sheet = spreadsheet.getSheetByName(SHEET_NAME);

function autoRekapData(e) {
  try {
    const inputan = e.values;
    const waktuIzin = inputan[2];
    const nim = inputan[1].trim();

    let posisiNIM = cariPosisiNIM(nim);
    const posisiWaktu = cariPosisiWaktu(waktuIzin);
    
    if (posisiNIM < 0) {
      posisiNIM= sheet.appendRow([nim]).getLastRow();
       
    }
    if (posisiWaktu>0){
    sheet.getRange(posisiNIM, posisiWaktu).setValue("Izin");
    }else{
        //kalo waktunya ga nemu
       sheet.getRange(posisiNIM, 2).setValue(sheet.getRange(posisiNIM, 2).getValue().toString() + "I")
    }
  } catch (error) {
    Logger.log("Terjadi kesalahan: " + error.message);
  }
}

// Fungsi untuk mencari posisi NIM
function cariPosisiNIM(nim) {
  try {
    const lastRow = sheet.getLastRow();
      const list_nim = sheet.getRange(1,1,lastRow).getValues()
   const posisi =  list_nim.findIndex((nim_sheet)=>nim_sheet[0].toString() == nim)
  return posisi > 0? posisi + 1: posisi
   
  } catch (error) {
    throw new Error("Error pada fungsi cariPosisiNIM: " + error.message);
  }
}

// Fungsi untuk konversi datetime
function konversiDatetime(timestampString) {
  try {
    Logger.log(timestampString)
         var parts = timestampString.split('/'); // Memisahkan tanggal, bulan, dan tahun
  var day = parts[0];
  var month = parts[1];
  var year = parts[2];

  // Menambahkan leading zero jika diperlukan
  if (day.length === 1) {
    day = '0' + day;
  }
  if (month.length === 1) {
    month = '0' + month;
  }
  year = year.substring(0,4)

  // Menggabungkan kembali string dengan leading zero
  Logger.log(day + '/' + month + '/' + year)
  return day + '/' + month + '/' + year;
  
  } catch (error) {
    throw new Error("Error pada fungsi konversiDatetime: " + error.message);
  }
}

// Fungsi untuk mencari posisi waktu
function cariPosisiWaktu(waktu) {
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
