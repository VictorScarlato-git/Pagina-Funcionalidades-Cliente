/**
 * Google Apps Script — Backend para o Guia Prático de Tutoriais
 *
 * SETUP:
 * 1. Crie uma Google Sheet (qualquer nome)
 * 2. Abra Extensions > Apps Script e cole este código
 * 3. Crie uma pasta no Google Drive para armazenar os PDFs e copie o ID dela
 * 4. Preencha SHEET_ID e DRIVE_FOLDER_ID abaixo
 * 5. Deploy > New deployment > Web app > Execute as: Me > Who has access: Anyone
 * 6. Copie a URL do deploy e cole na constante APPS_SCRIPT_URL dos arquivos HTML
 *
 * A aba "Tutoriais" será criada automaticamente com os cabeçalhos corretos.
 */

// ── CONFIG ──
const SHEET_ID = '';        // ID da Google Sheet (da URL: docs.google.com/spreadsheets/d/SHEET_ID/edit)
const DRIVE_FOLDER_ID = ''; // ID da pasta no Google Drive para os PDFs
const SHEET_NAME = 'Tutoriais';
const HEADERS = ['id', 'titulo', 'descricao', 'categoria', 'dataPublicacao', 'pdfUrl', 'pdfDriveId'];

/**
 * Retorna a aba "Tutoriais". Se não existir, cria com os cabeçalhos.
 */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── GET: Retorna todos os tutoriais como JSON (suporta JSONP via ?callback=) ──
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const lastRow = sheet.getLastRow();
    var result;

    if (lastRow <= 1) {
      result = { tutoriais: [] };
    } else {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const rows = data.slice(1);

      const tutoriais = rows
        .filter(row => row[0])
        .map(row => {
          const obj = {};
          headers.forEach((h, i) => {
            const key = String(h).trim().toLowerCase();
            const val = row[i];
            if (val instanceof Date) {
              obj[key] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            } else {
              obj[key] = String(val);
            }
          });
          // Normalizar pdfUrl: se for só um ID do Drive, converter para URL completa
          if (obj.pdfurl && !String(obj.pdfurl).startsWith('http')) {
            obj.pdfurl = 'https://drive.google.com/file/d/' + obj.pdfurl + '/view';
          }
          return obj;
        });

      result = { tutoriais: tutoriais };
    }

    const json = JSON.stringify(result);
    const callback = e && e.parameter && e.parameter.callback;

    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + json + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    const errorJson = JSON.stringify({ error: err.message });
    const callback = e && e.parameter && e.parameter.callback;

    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + errorJson + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(errorJson)
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
      pdfName || titulo + '.pdf'
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const pdfDriveId = file.getId();
    const pdfUrl = 'https://drive.google.com/file/d/' + pdfDriveId + '/view';

    // Insere na planilha (cria a aba se necessário)
    const sheet = getOrCreateSheet();
    const id = new Date().getTime().toString();
    sheet.appendRow([id, titulo, descricao, categoria, dataPublicacao, pdfUrl, pdfDriveId]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, id: id, pdfUrl: pdfUrl }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
