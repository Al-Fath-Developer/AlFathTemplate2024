/**
 * Fungsi yang dijalankan saat menu dibuka.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PDF Generator')
    .addItem('Generate PDFs from Slides', 'generatePDFsFromSlides')
    .addItem('Generate PDFs from Docs', 'generatePDFsFromDocs')
    .addItem('Create Template', 'createTemplate')

    .addToUi();
    ui.createMenu('Drive Tools')
      .addItem('Ekstrak Nama dan Deskripsi', 'extractFileInfo')
      .addItem('Update Nama dan Deskripsi', 'updateFileInfo')
      .addItem('Ekstrak Gambar ke Cell', 'insertImages')
            .addItem('Ekstrak List Link File', 'generateFileList')


      .addToUi();
}

/**
 * Memulai proses pembuatan PDF dari template Slides.
 */
function generatePDFsFromSlides() {
  generatePDFs('slides');
}

/**
 * Memulai proses pembuatan PDF dari template Docs.
 */
function generatePDFsFromDocs() {
  generatePDFs('docs');
}
/**
 * Membuat template pada baris pertama jika belum ada.
 */
function createTemplate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const firstRow = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];

  // Periksa apakah baris pertama kosong
  if (firstRow.every(cell => cell === '')) {
    // Menambahkan template pada baris pertama
    const templateData = [
      ['Template', '<<masukan id atau url template>>', 
       'Folder', '<<masukan id atau url folder hasil>>', 
       'Jumlah Kolom Data', 3, 
       'Posisi Hasil', 4, 
       'Keterangan Judul']
    ];
    sheet.getRange(1, 1, 1, templateData[0].length).setValues(templateData);
    SpreadsheetApp.getUi().alert('Template berhasil ditambahkan.');
  } else {
    // Menambahkan baris baru di atas baris pertama
    sheet.insertRowBefore(1);
    // Menambahkan template pada baris pertama
    const templateData = [
      ['Template', '<<masukan id atau url template>>', 
       'Folder Hasil', '<<masukan id atau url folder hasil>>', 
       'Jumlah Kolom Data', 3, 
       'Posisi Hasil', 4, 
       'Keterangan Judul']
    ];
    sheet.getRange(1, 1, 1, templateData[0].length).setValues(templateData);
    SpreadsheetApp.getUi().alert('Baris baru telah ditambahkan dan template berhasil ditambahkan.');
  }
}
/**
 * Memproses pembuatan PDF berdasarkan tipe template.
 * 
 * @param {string} templateType - Tipe template, bisa 'slides' atau 'docs'.
 */
function generatePDFs(templateType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const templateIdOrUrl = sheet.getRange('B1').getValue();
  const folderIdOrUrl = sheet.getRange('D1').getValue();
  const dataRange = parseInt(sheet.getRange('F1').getValue());
  const pdfColumn = parseInt(sheet.getRange('H1').getValue());
  const ketJudul = sheet.getRange('J1').getValue();
  
  const templateId = extractId(templateIdOrUrl);
  const folderId = extractId(folderIdOrUrl);
  const folder = DriveApp.getFolderById(folderId);
  const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, dataRange).getValues();
  const ui = SpreadsheetApp.getUi();
  const templateName = templateType === 'slides' ? SlidesApp.openById(templateId).getName() : DocumentApp.openById(templateId).getName();
  const estimatedTime = Math.ceil(data.length * 7); // Estimasi waktu (detik)
  
  const confirmation = ui.alert(
    'Konfirmasi',
    `Anda akan generate PDF sebanyak ${data.length} atau kurang dari itu jika sudah tergenerate sebelumnya dengan lokasi folder hasil di "${folder.getName()}" berdasarkan template "${templateName}" dengan estimasi waktu lebih ${estimatedTime} detik.
    Catatan: batas maksimal eksekusi selama sekitar 360 detik, jika estimasi eksekusinya lebih dari itu, maka Anda harus generate ulang ketika waktu eksekusi telah habis
     Apakah Anda ingin melanjutkan?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) {
    ui.alert('Terima kasih!');
    return;
  }

  const cache = CacheService.getScriptCache();
  // let index = parseInt(cache.get('currentIndex')) || 0;
  let index =  0;

  const startTime = new Date();
  let lastPdfTitle = '';

  for (; index < data.length; index++) {
    const row = data[index];
    const pdfUrlCell = sheet.getRange(3 + index, pdfColumn);
    const existingPdfUrl = pdfUrlCell.getValue();
    
    if (!existingPdfUrl) {
      try {
        if (templateType === 'slides') {
          lastPdfTitle = generatePDFfromSlides(templateId, row, ketJudul, folder, pdfUrlCell);
        } else if (templateType === 'docs') {
          lastPdfTitle = generatePDFfromDocs(templateId, row, ketJudul, folder, pdfUrlCell);
        }
      } catch (error) {
        ui.alert(`Error pada baris ${index + 3}: ${error.message}`);
        cache.put('currentIndex', index.toString());
        return;
      }
    }

    // // Periksa apakah sudah lebih dari 5 menit
    // const elapsedTime = (new Date() - startTime) / 1000;
    // if (elapsedTime > 330) { // 5 menit 30 detik
    //   cache.put('currentIndex', index.toString());
    //   ScriptApp.newTrigger('continuePDFGeneration')
    //     .timeBased()
    //     .after(1000)
    //     .create();
    //   return;
    // }
  }
  
  cache.remove('currentIndex');
  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(2);
  sheet.getRange(2, 1, 1, sheet.getMaxColumns()).setBackground('green');

  ui.alert(
    `Proses pembuatan PDF selesai! Judul PDF terakhir yang dibuat: "${lastPdfTitle}".\n\nWaktu mulai: ${startTime}\nWaktu selesai: ${endTime}\nDurasi: ${duration} detik.\n\nTerima kasih!\n\Penanggung Jawab: Tio Haidar Hanif\n*Note: Kodingan ini mayoritas di generate dari ChatGPT`
  );

  deleteTriggers(); // Hapus semua trigger setelah pekerjaan selesai
}

/**
 * Fungsi untuk melanjutkan pembuatan PDF setelah waktu eksekusi habis.
 */
function continuePDFGeneration() {
  
  const cache = CacheService.getScriptCache();
  const templateType = cache.get('templateType');
  generatePDFs(templateType);
}

/**
 * Membuat PDF dari template Slides.
 * 
 * @param {string} templateId - ID template Slides.
 * @param {Array} row - Data baris dari Google Sheets.
 * @param {string} ketJudul - Judul tambahan untuk PDF.
 * @param {Folder} folder - Folder tempat menyimpan PDF.
 * @param {Range} pdfUrlCell - Sel tempat menyimpan URL PDF.
 * @returns {string} - Judul PDF yang dibuat.
 */
function generatePDFfromSlides(templateId, row, ketJudul, folder, pdfUrlCell) {
  const newSlide = DriveApp.getFileById(templateId).makeCopy(row[0] + "_" + ketJudul);
  const newSlideId = newSlide.getId();
  const slide = SlidesApp.openById(newSlideId);
  const slides = slide.getSlides();
  
  slides.forEach(slideContent => {
    const shapes = slideContent.getShapes();
    shapes.forEach(shape => {
      const textRange = shape.getText();
      row.forEach((value, j) => {
        const placeholder = `<<${j + 1}>>`;
        textRange.replaceAllText(placeholder, value);
      });
    });
  });

  slide.saveAndClose();
  
  const pdfFile = DriveApp.createFile(newSlide.getAs('application/pdf'));
  folder.addFile(pdfFile);
  DriveApp.getRootFolder().removeFile(pdfFile);
  
  const pdfUrl = pdfFile.getUrl();
  pdfUrlCell.setValue(pdfUrl);
  
  DriveApp.getFileById(newSlideId).setTrashed(true);
  
  return newSlide.getName();
}

/**
 * Membuat PDF dari template Docs.
 * 
 * @param {string} templateId - ID template Docs.
 * @param {Array} row - Data baris dari Google Sheets.
 * @param {string} ketJudul - Judul tambahan untuk PDF.
 * @param {Folder} folder - Folder tempat menyimpan PDF.
 * @param {Range} pdfUrlCell - Sel tempat menyimpan URL PDF.
 * @returns {string} - Judul PDF yang dibuat.
 */
function generatePDFfromDocs(templateId, row, ketJudul, folder, pdfUrlCell) {
  const newDoc = DriveApp.getFileById(templateId).makeCopy(row[0] + "_" + ketJudul);
  const newDocId = newDoc.getId();
  const doc = DocumentApp.openById(newDocId);
  const body = doc.getBody();
  
  row.forEach((value, j) => {
    const placeholder = `<<${j + 1}>>`;
    body.replaceText(placeholder, value);
  });

  doc.saveAndClose();
  
  const pdfFile = DriveApp.createFile(newDoc.getAs('application/pdf'));
  folder.addFile(pdfFile);
  DriveApp.getRootFolder().removeFile(pdfFile);
  
  const pdfUrl = pdfFile.getUrl();
  pdfUrlCell.setValue(pdfUrl);
  
  DriveApp.getFileById(newDocId).setTrashed(true);
  
  return newDoc.getName();
}

/**
 * Ekstraksi ID dari URL atau mengembalikan string asli jika itu adalah ID.
 * 
 * @param {string} idOrUrl - ID atau URL untuk diekstraksi.
 * @returns {string} - ID yang diekstraksi.
 */
function extractId(idOrUrl) {
  const urlPattern = /[-\w]{25,}/;
  const match = idOrUrl.match(urlPattern);
  return match ? match[0] : idOrUrl;
}

/**
 * Hapus semua trigger yang ada.
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
}

function extractFileInfo() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Meminta input kolom file link
  var responseFileCol = ui.prompt('Masukkan nomor kolom untuk file link (misal: 1 untuk kolom A):');
  var fileCol = parseInt(responseFileCol.getResponseText());
  
  // Menentukan kolom untuk nama dan deskripsi
  var nameCol = fileCol + 1;
  var descCol = fileCol + 2;

  // Mendapatkan data dari kolom file link
  var data = sheet.getRange(2, fileCol, sheet.getLastRow() - 1, 1).getValues().filter(String);
  var fileCount = data.length;

  var confirm = ui.alert('Konfirmasi', 'Apakah Anda ingin mengambil nama dan deskripsi sebanyak ' + fileCount + ' file?', ui.ButtonSet.YES_NO);

  if (confirm == ui.Button.YES) {
    var startTime = new Date();
    for (var i = 0; i < data.length; i++) {
      var fileId = data[i][0].match(/[-\w]{25,}/);
      if (fileId) {
        var file = DriveApp.getFileById(fileId[0]);
        sheet.getRange(i + 2, nameCol).setValue(file.getName());
        sheet.getRange(i + 2, descCol).setValue(file.getDescription());
      }
    }
    var endTime = new Date();
    var duration = (endTime - startTime) / 1000;
    ui.alert('Proses selesai', 'Terima kasih! Proses selesai dalam ' + duration + ' detik.', ui.ButtonSet.OK);
  }
}

function updateFileInfo() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Meminta input kolom file link
  var responseFileCol = ui.prompt('Masukkan nomor kolom untuk file link (misal: 1 untuk kolom A):');
  var fileCol = parseInt(responseFileCol.getResponseText());
  
  // Menentukan kolom untuk nama dan deskripsi yang akan di-update
  var nameCol = fileCol + 1;
  var descCol = fileCol + 2;

  // Mendapatkan data dari kolom file link
  var data = sheet.getRange(2, fileCol, sheet.getLastRow() - 1, descCol).getValues().filter(function(row) { return row[0]; });
  var fileCount = data.length;

  var confirm = ui.alert('Konfirmasi', 'Apakah Anda ingin mengubah nama file dan deskripsi sebanyak ' + fileCount + ' file?', ui.ButtonSet.YES_NO);

  if (confirm == ui.Button.YES) {
    var startTime = new Date();
    for (var i = 0; i < data.length; i++) {
      var fileId = data[i][0].match(/[-\w]{25,}/);
      if (fileId) {
        var file = DriveApp.getFileById(fileId[0]);
        file.setName(data[i][1]); // Kolom nama
        file.setDescription(data[i][2]); // Kolom deskripsi
      }
    }
    var endTime = new Date();
    var duration = (endTime - startTime) / 1000;
    ui.alert('Proses selesai', 'Terima kasih! Proses selesai dalam ' + duration + ' detik.', ui.ButtonSet.OK);
  }
}

function insertImages() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Meminta input kolom file link
  var responseFileCol = ui.prompt('Masukkan nomor kolom untuk file link (misal: 1 untuk kolom A):');
  var fileCol = parseInt(responseFileCol.getResponseText());
  
  // Meminta input kolom untuk gambar
  var responseImgCol = ui.prompt('Masukkan nomor kolom untuk gambar (misal: 2 untuk kolom B):');
  var imgCol = parseInt(responseImgCol.getResponseText());

  // Mendapatkan data dari kolom file link
  var data = sheet.getRange(2, fileCol, sheet.getLastRow() - 1, 1).getValues().filter(String);
  var fileCount = data.length;

  var confirm = ui.alert('Konfirmasi', 'Apakah Anda ingin memasukkan gambar sebanyak ' + fileCount + ' file?', ui.ButtonSet.YES_NO);

  if (confirm == ui.Button.YES) {
    var startTime = new Date();
    for (var i = 0; i < data.length; i++) {
      var fileId = data[i][0].match(/[-\w]{25,}/);
      if (fileId) {
        try {
          var file = DriveApp.getFileById(fileId[0]);
          var permissions = file.getSharingAccess();
          var imgUrl = '';
          
          if (permissions === DriveApp.Access.ANYONE_WITH_LINK) {
            imgUrl = 'https://drive.google.com/thumbnail?authuser=0&sz=w320&id=' + fileId[0];
            var cell = sheet.getRange(i + 2, imgCol);
            cell.setFormula('=IMAGE("' + imgUrl + '")');
          } else {
            var cell = sheet.getRange(i + 2, imgCol);
            cell.setValue('atur akses file menjadi anyone with the link agar bisa menampilkan gambar');
          }
        } catch (e) {
          sheet.getRange(i + 2, imgCol).setValue('File tidak ditemukan atau Anda tidak memiliki akses');
        }
      }
    }
    var endTime = new Date();
    var duration = (endTime - startTime) / 1000;
    ui.alert('Proses selesai', 'Terima kasih! Proses selesai dalam ' + duration + ' detik.', ui.ButtonSet.OK);
  }
}

function generateFileList() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Meminta input link folder dari pengguna
  var responseFolderLink = ui.prompt('Masukkan link folder Google Drive (contoh: https://drive.google.com/drive/folders/xxxxxxxxxxxxxxxxxxxx):');
  var folderLink = responseFolderLink.getResponseText();
  
  // Mendapatkan ID folder dari link
  var folderId = getIdFromUrl(folderLink);
  if (!folderId) {
    ui.alert('Error', 'Link folder tidak valid. Pastikan Anda memasukkan link yang benar.', ui.ButtonSet.OK);
    return;
  }

  var folder = DriveApp.getFolderById(folderId);

  // Meminta input kolom untuk menyimpan hasil
  var responseResultCol = ui.prompt('Masukkan nomor kolom untuk menyimpan hasil (misal: 1 untuk kolom A):');
  var resultCol = parseInt(responseResultCol.getResponseText());

  // Menulis header
  sheet.getRange(1, resultCol).clear(); // Membersihkan kolom hasil sebelum menulis baru
  sheet.getRange(1, resultCol).setValue('Link');
  sheet.getRange(1, resultCol + 1).setValue('Judul File');
  sheet.getRange(1, resultCol + 2).setValue('Nama Folder Parent');

  var stack = []; // Stack untuk menyimpan folder yang akan dieksplorasi
  stack.push({ folder: folder, parentName: '' }); // Memasukkan folder utama ke dalam stack

  var row = 2;

  while (stack.length > 0) {
    var current = stack.pop();
    var currentFolder = current.folder;
    var parentFolderName = current.parentName;

    var files = currentFolder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      var fileId = file.getId();
      var fileTitle = file.getName();
      var fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
      sheet.getRange(row, resultCol).setValue(fileUrl);
      sheet.getRange(row, resultCol + 1).setValue(fileTitle);
      sheet.getRange(row, resultCol + 2).setValue(parentFolderName);
      row++;
    }

    var subfolders = currentFolder.getFolders();
    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
      stack.push({ folder: subfolder, parentName: currentFolder.getName() });
    }
  }

  // Memberi tahu bahwa proses selesai
  ui.alert('Proses selesai', 'List link file sudah dibuat di spreadsheet Anda.', ui.ButtonSet.OK);
}

// Fungsi untuk mendapatkan ID folder dari URL
function getIdFromUrl(url) {
  var id;
  var match = url.match(/[-\w]{25,}/);
  if (match) {
    id = match[0];
  }
  return id;
}


// ini kodingan mayoritas dibuat oleh ChatGPT, tio cuman minor aja

//contoh sheet: https://docs.google.com/spreadsheets/d/18TRhW0InKjDxEXno3NkYqPD91itMC97Pg3ViJ36HLlg/edit?gid=0#gid=0
