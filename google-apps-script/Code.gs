/**
 * Google Apps Script — Backend para o Guia Prático de Tutoriais
 *
 * SETUP:
 * 1. Crie uma Google Sheet com o nome "Tutoriais"
 * 2. Na primeira linha (cabeçalho), coloque: id | titulo | descricao | categoria | dataPublicacao | pdfUrl | pdfDriveId
 * 3. Abra Extensions > Apps Script e cole este código
 * 4. Crie uma pasta no Google Drive para armazenar os PDFs e copie o ID dela
 * 5. Preencha SHEET_ID e DRIVE_FOLDER_ID abaixo
 * 6. Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 * 7. Copie a URL do deploy e cole na constante APPS_SCRIPT_URL dos arquivos HTML
 */

// ── CONFIG ──
const SHEET_ID = '';        // ID da Google Sheet (da URL: docs.google.com/spreadsheets/d/SHEET_ID/edit)
const DRIVE_FOLDER_ID = ''; // ID da pasta no Google Drive para os PDFs
const SHEET_NAME = 'Tutoriais';

// ── GET: Retorna todos os tutoriais como JSON ──
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    const tutoriais = rows
      .filter(row => row[0])
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i]; });
        if (obj.dataPublicacao instanceof Date) {
          obj.dataPublicacao = Utilities.formatDate(obj.dataPublicacao, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        return obj;
      });

    return ContentService
      .createTextOutput(JSON.stringify({ tutoriais }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST: Recebe tutorial (metadados + PDF base64) e salva ──
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { titulo, descricao, categoria, dataPublicacao, pdfBase64, pdfName } = payload;

    // Salva o PDF no Google Drive
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const blob = Utilities.newBlob(
      Utilities.base64Decode(pdfBase64),
      'application/pdf',
      pdfName || `${titulo}.pdf`
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const pdfDriveId = file.getId();
    const pdfUrl = `https://drive.google.com/file/d/${pdfDriveId}/view`;

    // Insere na planilha
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    const id = new Date().getTime().toString();
    sheet.appendRow([id, titulo, descricao, categoria, dataPublicacao, pdfUrl, pdfDriveId]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, id, pdfUrl }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
