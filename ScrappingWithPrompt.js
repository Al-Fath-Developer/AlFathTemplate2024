
function onOpen() {
const ui = SpreadsheetApp.getUi();
  try {
    ui.createMenu('Scrapping')
      .addItem("Mode Simple", "startGeneratePromptAllFolderInEmail")
      .addItem("Mode Per Folder", "startGeneratePromptPerFolder")
      .addToUi();
  } catch (error) {
    throw new Error("Error in onOpen function: " + error.message);
  }
}

function startGeneratePromptAllFolderInEmail() {
    const ui = SpreadsheetApp.getUi();
  try {
    SKIP_TILL = sheet.getLastRow();
    const alertMessage = "Proses Scrapping Insya Allah akan dimulai. Update bakal dilakukan dari baris ke "+SKIP_TILL+"\n Tekan OK lalu Silakan menunggu.\n\nCatatan: Proses ini mungkin memakan waktu yang cukup lama. Anda bisa melakukan kegiatan lain sambil menunggu.\n\nJika Anda mendapatkan kesalahan terkait jumlah eksekusi yang telah mencapai batas maksimum, jalankan fitur ini lagi. Lakukan langkah tersebut hingga Anda melihat tulisan 'Done' di baris terbawah.";
    const uiResponse = ui.alert("Ok Siap", alertMessage, ui.ButtonSet.OK_CANCEL);
    if (uiResponse == ui.Button.OK){

  
    const rootFolder = DriveApp.getRootFolder();
    const driveId = rootFolder.getId();
    ID_FOLDER_YANG_DITELUSURI = driveId;
    startGeneration();
    }

  } catch (error) {
    Logger.log("Error in startGeneratePromptAllFolderInEmail function: " + error.message);
    ui.alert("Maaf, terjadi kesalahan: " + error.message);
  }
}

function startGeneratePromptPerFolder() {
    const ui = SpreadsheetApp.getUi();
  try {
    const idResponse = ui.prompt("ID Folder", "Masukkan ID Folder yang ingin di-scrapping", ui.ButtonSet.OK_CANCEL);
    if (idResponse.getSelectedButton() == ui.Button.OK) {
      const folderId = idResponse.getResponseText();
      ID_FOLDER_YANG_DITELUSURI = folderId;
  
      const folderName = DriveApp.getFolderById(ID_FOLDER_YANG_DITELUSURI).getName();
      const lastRow = sheet.getLastRow();
      const skipTillResponse = ui.prompt("SKIP UNTIL", `Folder yang akan di-scrapping: ${folderName}\nMulai melanjutkan scrapping dari baris keberapa?:\nJika tidak yakin, masukkan angka ${lastRow}`, ui.ButtonSet.OK_CANCEL);
      if (skipTillResponse.getSelectedButton() == ui.Button.OK) {
        const skipTillRow = parseInt(skipTillResponse.getResponseText());
        SKIP_TILL = skipTillRow;
        startGeneration();
        ui.alert(`Folder ${folderName} selesai di-scrapping.`);
      }
    }
  } catch (error) {
    Logger.log("Error in startGeneratePromptPerFolder function: " + error.message);
    ui.alert("Maaf, terjadi kesalahan: " + error.message);
  }
}


let ID_FOLDER_YANG_DITELUSURI = ""; // ID folder yang akan ditelusuri isinya
const INITIAL_ROW = 1; // Mulai dari baris keberapa tabelnya
let SKIP_TILL = 0; // Jika terjadi error, mulai eksekusi dari baris terakhir yang tersimpan

const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
let currentRow = INITIAL_ROW;

function setMessage(pesan) {
  Logger.log(pesan);
}

function startGeneration() {
  try {
    if (SKIP_TILL < 2) {
      header(currentRow, DriveApp.getFolderById(ID_FOLDER_YANG_DITELUSURI).getName());
    }
    currentRow += 2;
    getFileLinksInFolder(ID_FOLDER_YANG_DITELUSURI);
    footer(currentRow);
    return "Done";
  } catch (error) {
    throw new Error("Error: " + error.message);
  }
}

function getFileLinksInFolder(folderId) {
  try {
    const stack = [folderId]; // Stack untuk menyimpan folder yang akan diproses
  
    while (stack.length > 0) {
      const currentFolderId = stack.pop();
      const folder = DriveApp.getFolderById(currentFolderId);
  
      setMessage("Folder yang Sedang Diproses: " + folder.getName());
      
      const files = folder.getFiles();
  
      while (files.hasNext()) {
        const file = files.next();
        if (currentRow >= SKIP_TILL) {
          const fileInfo = getFileInfo(file);
          insertToSpreadsheet(currentRow, fileInfo);
        }
        currentRow++;
      }
  
      const subfolders = folder.getFolders();
      while (subfolders.hasNext()) {
        const subfolder = subfolders.next();
        stack.push(subfolder.getId());
      }
    }
  } catch (error) {
    throw new Error("Error while getting file links: " + error.message);
  }
}

function getFileInfo(file) {
  const fileUrl = file.getUrl();
  const fileName = file.getName();
  const fileType = file.getMimeType();
  const fileCreatedAt = file.getDateCreated();
  const fileUpdatedAt = file.getLastUpdated();
  const fileSize = formatSize(file.getSize());
  const fileSharingType = file.getSharingPermission();
  const folderParent = getParentFolder(file);
  return { fileUrl, fileName, fileType, fileSize, fileSharingType, fileCreatedAt, fileUpdatedAt, folderParent };
}

function insertToSpreadsheet(row, fileInfo) {
  try {
    sheet.getRange(`A${row}:I${row}`).setValues([[fileInfo.fileName, fileInfo.fileUrl, fileInfo.fileType, fileInfo.fileSize, fileInfo.fileSharingType, fileInfo.fileCreatedAt, fileInfo.fileUpdatedAt, fileInfo.folderParent.link, fileInfo.folderParent.name]])
      .setBackground("#fffffa")
      .setFontWeight('normal')
      .setHorizontalAlignments([['center', 'left', 'right', 'left', 'left', 'center', 'center', 'left', 'left']])
      .setWraps([[true, false, false, false, false, false, false, false, true]])
      .setBorder(true, true, true, true, true, true);
    setMessage("File Selesai: " + fileInfo.fileName);
  } catch (error) {
    throw new Error("Error while inserting to spreadsheet: " + error.message);
  }
}

function formatSize(sizeInBytes) {
  try {
    if (sizeInBytes < 1024) {
      return sizeInBytes + " bytes";
    } else if (sizeInBytes < 1024 * 1024) {
      return (sizeInBytes / 1024).toFixed(2) + " KB";
    } else if (sizeInBytes < 1024 * 1024 * 1024) {
      return (sizeInBytes / (1024 * 1024)).toFixed(2) + " MB";
    } else {
      return (sizeInBytes / (1024 * 1024 * 1024)).toFixed(2) + " GB";
    }
  } catch (error) {
    throw new Error("Error while formatting size: " + error.message);
  }
}

function getParentFolder(file) {
  try {
    const parentFolders = file.getParents();
    if (parentFolders.hasNext()) {
      const parentFolder = parentFolders.next();
      return { name: parentFolder.getName(), link: parentFolder.getUrl() };
    } else {
      return { name: "Folder Induk Tidak Ditemukan", link: "-" };
    }
  } catch (error) {
    throw new Error("Error while getting parent folder: " + error.message);
  }
}

function header(row, folderName) {
  try {
    if (ID_FOLDER_YANG_DITELUSURI != DriveApp.getRootFolder().getId()){

    sheet.setName(folderName);
    }
    const topLeftCell = sheet.getRange('A' + row);
    const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEEE, dd MMMM yyyy HH:mm:ss");
    topLeftCell.setBackground("black").setFontColor('white').setValue("Scrapping Folder '" + folderName + "' Start From ± " + time).setFontSize(13).setHorizontalAlignment('center').setFontWeight('bold');
    sheet.getRange(`A${row}:I${row}`).merge();
    sheet.getRange(`A${row + 1}:I${row + 1}`).setValues([["Nama File", "Link", "Tipe File", "Ukuran", "Tipe Sharing", "Created at", "Updated at", "Parent Link", "Parent Folder Name"]]).setBackground("grey").setFontWeight("bold").setHorizontalAlignment("center").setWrap(true).setBorder(true, true, true, true, true, true);
  } catch (error) {
    throw new Error("Error while creating header: " + error.message);
  }
}

function footer(row) {
  try {
    setMessage("Proses Selesai, Tambahkan footer di row + " + row);
    const footerCell = sheet.getRange('A' + row);
    const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEEE, dd MMMM yyyy HH:mm:ss");
    footerCell.setBackground("black").setFontColor('white').setValue("Done. Finish Time ± " + time).setFontSize(13).setHorizontalAlignment('center');
    sheet.getRange(`A${row}:I${row}`).merge();
    const sosmed = "wa.me/6282135455703";
    sheet.getRange('A' + (row + 1)).setValue("Penanggung Jawab Kodingan: Tio Haidar Hanif");
    sheet.getRange('A' + (row + 2)).setFormula('=hyperlink("' + sosmed + '")');
  } catch (error) {
    throw new Error("Error while creating footer: " + error.message);
  }
}
