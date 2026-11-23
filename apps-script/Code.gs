/**
 * AcordosHub — Backend Google Apps Script para sincronização bidirecional em tempo real
 * Planilha alvo: https://docs.google.com/spreadsheets/d/1FXKNHPd_YwR2C-zSDaXlFiH3EMEjM5GvDcv8ykW8gK8/edit?gid=2121317474
 *
 * COMO INSTALAR (2 minutos):
 * 1. Abra sua planilha no Google Sheets.
 * 2. Vá em: Extensões > Apps Script.
 * 3. Apague o código inicial e cole TODO este arquivo.
 * 4. Clique em "Implantar" > "Nova implantação" > Tipo: "App da Web".
 *    - Executar como: "Eu" (sua conta)
 *    - Quem pode acessar: "Qualquer pessoa"
 * 5. Copie a URL gerada (termina em /exec) e cole no botão "Conectar Planilha" do site.
 */

const CONFIG = {
  // GID da aba principal informada no link
  TARGET_GID: 2121317474,
  // Linha onde ficam os cabeçalhos (padrão = 1)
  HEADER_ROW: 1
};

function getTargetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === CONFIG.TARGET_GID) {
      return sheets[i];
    }
  }
  return ss.getActiveSheet();
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatCellValue_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return val;
}

/**
 * GET: Retorna todos os cabeçalhos e linhas da planilha em tempo real
 */
function doGet(e) {
  try {
    const sheet = getTargetSheet_();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < CONFIG.HEADER_ROW || lastCol === 0) {
      return jsonResponse_({
        ok: true,
        sheetName: sheet.getName(),
        gid: sheet.getSheetId(),
        headers: [],
        rows: [],
        updatedAt: new Date().toISOString()
      });
    }

    const values = sheet.getRange(CONFIG.HEADER_ROW, 1, lastRow - CONFIG.HEADER_ROW + 1, lastCol).getValues();
    const headers = values[0].map((h, idx) => String(h || ('Coluna ' + (idx + 1))).trim());

    const rows = [];
    for (let i = 1; i < values.length; i++) {
      const rawRow = values[i];
      const hasAnyValue = rawRow.some(cell => cell !== '' && cell !== null && cell !== undefined);
      if (!hasAnyValue) continue;

      const obj = { _rowNumber: CONFIG.HEADER_ROW + i };
      headers.forEach((header, colIdx) => {
        obj[header] = formatCellValue_(rawRow[colIdx]);
      });
      rows.push(obj);
    }

    return jsonResponse_({
      ok: true,
      sheetName: sheet.getName(),
      gid: sheet.getSheetId(),
      spreadsheetTitle: SpreadsheetApp.getActiveSpreadsheet().getName(),
      headers: headers,
      rows: rows,
      updatedAt: new Date().toISOString()
    });
  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: err.message || String(err)
    });
  }
}

/**
 * POST: Inclui ou exclui linhas na planilha
 * Payload JSON esperado:
 * - Inclusão: { action: "add", data: { "Coluna A": "valor", ... } }
 * - Exclusão: { action: "delete", rowNumbers: [5, 7] }
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    const sheet = getTargetSheet_();

    if (action === 'add') {
      const lastCol = sheet.getLastColumn();
      let headers = [];

      if (lastCol > 0 && sheet.getLastRow() >= CONFIG.HEADER_ROW) {
        headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
      } else {
        headers = Object.keys(body.data || {});
        if (headers.length > 0) {
          sheet.getRange(CONFIG.HEADER_ROW, 1, 1, headers.length).setValues([headers]);
        }
      }

      const data = body.data || {};
      const newRow = headers.map(h => (data[h] !== undefined && data[h] !== null ? data[h] : ''));
      sheet.appendRow(newRow);

      return jsonResponse_({
        ok: true,
        message: 'Acordo incluído com sucesso na planilha.',
        rowNumber: sheet.getLastRow(),
        updatedAt: new Date().toISOString()
      });
    }

    if (action === 'delete') {
      const rowNumbers = Array.isArray(body.rowNumbers)
        ? body.rowNumbers.map(Number).filter(n => Number.isInteger(n) && n > CONFIG.HEADER_ROW)
        : [];

      if (rowNumbers.length === 0) {
        return jsonResponse_({ ok: false, error: 'Nenhuma linha válida informada para exclusão.' });
      }

      // Excluir de baixo para cima para não deslocar índices
      const sortedDesc = [...new Set(rowNumbers)].sort((a, b) => b - a);
      sortedDesc.forEach(rowNum => {
        if (rowNum <= sheet.getLastRow()) {
          sheet.deleteRow(rowNum);
        }
      });

      return jsonResponse_({
        ok: true,
        deletedCount: sortedDesc.length,
        message: 'Registro(s) excluído(s) da planilha com sucesso.',
        updatedAt: new Date().toISOString()
      });
    }

    return jsonResponse_({ ok: false, error: 'Ação não reconhecida: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: err.message || String(err) });
  } finally {
    lock.releaseLock();
  }
}
