/**
 * Fungsi yang dijalankan saat menu dibuka.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('PDF Generator')
    .addItem('Generate PDFs from Slides', 'generatePDFsFromSlides')
    .addItem('Generate PDFs from Docs', 'generatePDFsFromDocs')
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
 * Memproses pembuatan PDF berdasarkan tipe template.
 * 
 * @param {string} templateType - Tipe template, bisa 'slides' atau 'docs'.
 */
function generatePDFs(templateType) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  const templateId = sheet.getRange('B1').getValue();
  const folderId = sheet.getRange('D1').getValue();
  const dataRange = parseInt(sheet.getRange('F1').getValue());
  const pdfColumn = parseInt(sheet.getRange('H1').getValue());
  const ketJudul = sheet.getRange('J1').getValue();
  
  const folder = DriveApp.getFolderById(folderId);
  const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, dataRange).getValues();
  const ui = SpreadsheetApp.getUi();
  const templateName = templateType === 'slides' ? SlidesApp.openById(templateId).getName() : DocumentApp.openById(templateId).getName();
  const estimatedTime = Math.ceil(data.length * 7 ); // Estimasi waktu (detik)
  
  const confirmation = ui.alert(
    'Konfirmasi',
    `Anda akan generate PDF sebanyak ${data.length} dengan lokasi folder hasil di "${folder.getName()}" berdasarkan template "${templateName}" dengan estimasi waktu ${estimatedTime} detik. Apakah Anda ingin melanjutkan?`,
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) {
    ui.alert('Terima kasih!');
    return;
  }

  const cache = CacheService.getScriptCache();
  let index = parseInt(cache.get('currentIndex')) || 0;
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
  }
  
  cache.remove('currentIndex');
  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(2);

  ui.alert(
    `Proses pembuatan PDF selesai! Judul PDF terakhir yang dibuat: "${lastPdfTitle}".\n\nWaktu mulai: ${startTime}\nWaktu selesai: ${endTime}\nDurasi: ${duration} detik.\n\nTerima kasih!\n\nCredit by [Your Name]`
  );
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
  const newSlide = DriveApp.getFileById(templateId).makeCopy(`${row[0]}_${ketJudul}`);
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
  const newDoc = DriveApp.getFileById(templateId).makeCopy(`${row[0]}_${ketJudul}`);
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

// ini kodingan mayoritas dibuat oleh ChatGPT, tio cuman minor aja

//contoh sheet: https://docs.google.com/spreadsheets/d/18TRhW0InKjDxEXno3NkYqPD91itMC97Pg3ViJ36HLlg/edit?gid=0#gid=0
