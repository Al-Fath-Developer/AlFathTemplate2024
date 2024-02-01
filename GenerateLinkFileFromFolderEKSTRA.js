// rekap ekstra
const SHEET_REKAP_NAME = "Folder 1" // nama sheet untuk rekapnya ada dinana (pastiin kosong dulu)
const ID_FOLDER_YANG_DITELUSURI = "1EP4swrEbsB7BRh9gO2Ok2-tEuL4cFjZ6"; // folder yang mau ditelusuri isi isinya th yang mana
const INITIAL_ROW   =1 // tabel nya mau mulai dari row keberapa
// 



const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REKAP_NAME)

let current_row = INITIAL_ROW

function mulaiEkstraGenerate(){
  header(current_row, DriveApp.getFolderById(ID_FOLDER_YANG_DITELUSURI).getName())
  current_row = current_row + 2
  getFileLinksInFolder(ID_FOLDER_YANG_DITELUSURI)
  footer(current_row)

}

function getFileLinksInFolder(folder_id) {
  // Dapatkan folder berdasarkan ID
  var folder = DriveApp.getFolderById(folder_id);

  Logger.log("Folder On Proses: " + folder.getName()) 
  // Dapatkan semua file dalam folder
  var files = folder.getFiles();
  
  // Variabel untuk menyimpan link dan nama file
  
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
    var folderParent = getParentFolder(file)
    
    // Tambahkan link dan nama file ke dalam array
    insertToSpreadsheet(current_row, fileUrl, fileName,fileType,fileSize, fileSharingType,fileCreateedAt,fileUpdatedAt, folderParent.name, folderParent.link)
        current_row++

}
     var subfolders = folder.getFolders();
      while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    getFileLinksInFolder(subfolder.getId()); // Panggil fungsi lagi untuk subfolder ini
  }
  }

 

function header(row, nama_folder){
  const cell_pojok_kiri_atas = sheet.getRange('A'+ row)
  const waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEEE, dd MMMM yyyy HH:mm:ss")
  cell_pojok_kiri_atas.setBackground("black").setFontColor('white').setValue("Rekap til root Folder '" + nama_folder + "' Start From ± " +  waktu).setFontSize(13).setHorizontalAlignment('center').setFontWeight('bold')
  sheet.getRange(`A${row}:I${row}`).merge()
  sheet.getRange(`A${row + 1}:I${row + 1}`).setValues([["Nama File", "Link", "Tipe File", "Ukuran", "Tipe Sharing", "Created at", "Updated at", "Parent Link", "Parent Folder Name"]]).setBackground("grey").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true).setBorder(true,true,true,true,true,true)
  
}
function footer(row){
  Logger.log("beres tinggal footer di row + "+ row )
  const cell_footer = sheet.getRange('A'+ row)

  const waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEEE, dd MMMM yyyy HH:mm:ss")
  cell_footer.setBackground("black").setFontColor('white').setValue("Done. Finish Time ± " +  waktu).setFontSize(13).setHorizontalAlignment('center')
  sheet.getRange(`A${row}:I${row}`).merge()
  
  const ig = "https://www.instagram.com/tiohaidarhanif/"
  sheet.getRange('A' + (row + 1)).setValue("Penanggung Jawab Kodingan: Tio Haidar Hanif")
  sheet.getRange('A' + (row + 2)).setFormula('=hyperlink("'+ig+'")')

}
function insertToSpreadsheet(row,url, name, type, size, sharing_type, created, edited, folderParentName, folderParentLink){
  Logger.log([row,url, name, type, size, sharing_type, created, edited, folderParentName])
sheet.getRange(`A${row}:I${row}`).setValues([[name,url, type, size, sharing_type, created, edited, folderParentLink, folderParentName]]).setBackground("#fffffa").setFontWeight('normal').setHorizontalAlignments([['center', 'left', 'right', 'left', 'left', 'center', 'center', 'left', 'left']]).setWraps([[true, false, false, false, false,false,false, false, true]]).setBorder(true,true,true,true,true,true)
Logger.log("File Done: " + name)

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



function getParentFolder(file) {
  var parentFolders = file.getParents();
  
  // Periksa apakah file memiliki folder induk
  if (parentFolders.hasNext()) {
    var parentFolder = parentFolders.next();
    return {name :parentFolder.getName(), link:parentFolder.getUrl()};
  } else {
    return {name :"Folder Induk Tidak Ditemukan", link:"-"};
  }
}

// banyak dibantu Chat GPT
