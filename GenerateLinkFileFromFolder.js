const SPREADSHEET_ID = "1bIYD7-j6ynYsHO-OCr8GxI0lDlwKlINepGi_0FToM6I" //spreadsheet untuk rekap nya ada dimana
const SHEET_REKAP_NAME = "Folder 1" // nama sheet untuk rekapnya ada dinana (pastiin kosong dulu)
const ID_FOLDER_YANG_DITELUSURI = "1EP4swrEbsB7BRh9gO2Ok2-tEuL4cFjZ6"; // folder yang mau ditelusuri isi isinya th yang mana
const INITIAL_ROW   =1 // tabel nya mau mulai dari row keberapa
// 



const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_REKAP_NAME)




function mulaiGenerateLink() {
  // Dapatkan folder berdasarkan ID
  var folder = DriveApp.getFolderById(ID_FOLDER_YANG_DITELUSURI);

  // Dapatkan semua file dalam folder
  var files = folder.getFiles();
  
  // Variabel untuk menyimpan link dan nama file
  
  let current_row = INITIAL_ROW
  header(current_row)
  current_row++
  // Loop melalui setiap file dalam folder
  while (files.hasNext()) {
    var file = files.next();
    
    // Dapatkan URL file dan nama file
    var fileUrl = file.getUrl();
    var fileName = file.getName();
    var fileType = file.getMimeType()
    var fileCreateedAt =  file.getDateCreated() 
    var fileUpdatedAt = file.getLastUpdated()
    var fileSize =   formatSize(file.getSize())
    var fileSharingType =  file.getSharingPermission()
    
    // Tambahkan link dan nama file ke dalam array
    insertToSpreadsheet(current_row, fileUrl, fileName,fileType,fileSize, fileSharingType,fileCreateedAt,fileUpdatedAt)
    current_row++
  }

 
}
function header(row){
  sheet.getRange(`A${row}:G${row}`).setValues([["Link", "Nama File", "Tipe File", "Ukuran", "Tipe Sharing", "Created at", "Updated at"]]).setBackground("grey").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true).setBorder(true,true,true,true,true,true)
  
}
function insertToSpreadsheet(row,url, name, type, size, sharing_type, created, edited){

sheet.getRange(`A${row}:G${row}`).setValues([[url, name, type, size, sharing_type, created, edited]]).setBackground("#fffffa").setFontWeight('normal').setHorizontalAlignments([['left', 'left', 'right', 'left', 'left', 'center', 'center']]).setWraps([[false, true, false, false, false,false,false]]).setBorder(true,true,true,true,true,true)
}
function formatSize(sizeInBytes) {
  if (sizeInBytes < 1024) {
    return sizeInBytes + " bytes";
  } else if (sizeInBytes < 1024 * 1024) {
    return (sizeInBytes / 1024).toFixed(2) + " KB";
  } else if (sizeInBytes < 1024 * 1024 * 1024) {
    return (sizeInBytes / (1024 * 1024)).toFixed(2) + " MB";
  } else {
    return (sizeInBytes / (1024 * 1024 * 1024)).toFixed(2) + " GB";
  }
}

// banyak dibantu Chat GPT
