/**
 * CS Ops — Formulário de Mudanças no Time
 * Google Apps Script Web App
 *
 * Como configurar:
 *   1. Abra script.google.com
 *   2. Crie um novo projeto e cole este código
 *   3. Em "Serviços" confirme que não há dependências externas necessárias
 *   4. Vá em Implantar → Nova implantação → Tipo: Web App
 *      - Execute como: Eu (conta Google da Bianca)
 *      - Quem tem acesso: Qualquer pessoa
 *   5. Autorize as permissões (acesso ao Sheets)
 *   6. Copie a URL gerada e cole no index.html no lugar de COLE_AQUI_A_URL_DO_APPS_SCRIPT
 *
 * A planilha é criada automaticamente na primeira execução.
 * Nome: "CS Ops — Mudanças no Time"
 */

var SHEET_NAME      = "Mudanças";
var SPREADSHEET_NAME = "CS Ops — Mudanças no Time";


/**
 * Responde às requisições POST do formulário.
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var sheet = getOrCreateSheet();

    var row = [
      new Date(),                              // A: Timestamp (servidor)
      data.data_mudanca   || "",               // B: Data da mudança (informada)
      data.tipo           || "",               // C: Tipo
      data.nome           || "",               // D: Nome
      data.segmento       || "",               // E: Segmento (novo/atual)
      data.segmento_anterior || "",            // F: Segmento anterior
      data.cargo          || "",               // G: Cargo (novo/atual)
      data.cargo_anterior || "",               // H: Cargo anterior
      data.lideranca      || "",               // I: Liderança
      data.observacao     || "",               // J: Observação
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * Requisições GET retornam uma confirmação simples (útil para testar a URL).
 */
function doGet(e) {
  return HtmlService.createHtmlOutput(
    "<h2>CS Ops — Formulário de Mudanças no Time</h2>" +
    "<p>Web App ativa. Envie POST com os dados do formulário.</p>"
  );
}


/**
 * Retorna a aba da planilha, criando a planilha e o cabeçalho se necessário.
 */
function getOrCreateSheet() {
  var files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  var ss;

  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    Logger.log("Planilha criada: " + ss.getUrl());
  }

  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    // Cabeçalho
    var headers = [
      "Timestamp Envio",
      "Data da Mudança",
      "Tipo",
      "Nome",
      "Segmento (novo/atual)",
      "Segmento Anterior",
      "Cargo (novo/atual)",
      "Cargo Anterior",
      "Liderança",
      "Observação",
    ];
    sheet.appendRow(headers);

    // Formata cabeçalho: negrito + fundo roxo + texto branco
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#7668AC");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontFamily("Montserrat");
    sheet.setFrozenRows(1);

    // Larguras de coluna (px aproximado)
    sheet.setColumnWidth(1, 160);  // Timestamp
    sheet.setColumnWidth(2, 120);  // Data mudança
    sheet.setColumnWidth(3, 140);  // Tipo
    sheet.setColumnWidth(4, 200);  // Nome
    sheet.setColumnWidth(5, 180);  // Segmento novo
    sheet.setColumnWidth(6, 180);  // Segmento anterior
    sheet.setColumnWidth(7, 120);  // Cargo novo
    sheet.setColumnWidth(8, 120);  // Cargo anterior
    sheet.setColumnWidth(9, 140);  // Liderança
    sheet.setColumnWidth(10, 300); // Observação

    Logger.log("Aba criada com cabeçalho: " + SHEET_NAME);
  }

  return sheet;
}
