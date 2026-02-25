// ╔══════════════════════════════════════════════════════════╗
// ║   FAMÍLIA - BACKEND (Google Apps Script)                 ║
// ║   Cole este código no Google Apps Script e publique      ║
// ║   como Web App com acesso "Qualquer pessoa"              ║
// ╚══════════════════════════════════════════════════════════╝

// ─── CONFIGURAÇÃO ───────────────────────────────────────────
// Após criar sua planilha, cole o ID dela aqui.
// O ID fica na URL: docs.google.com/spreadsheets/d/SEU_ID_AQUI/edit
const SPREADSHEET_ID = 'COLE_O_ID_DA_SUA_PLANILHA_AQUI';

const SHEETS = {
  transactions: 'Transações',
  tasks: 'Tarefas'
};

// ─── ENTRY POINT ────────────────────────────────────────────
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json'
  };

  try {
    const params = e.parameter || {};
    const action = params.action;
    let body = {};

    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    let result;

    switch (action) {
      case 'getTransactions':   result = getTransactions();              break;
      case 'addTransaction':    result = addTransaction(body);           break;
      case 'deleteTransaction': result = deleteTransaction(body.id);     break;
      case 'getTasks':          result = getTasks();                     break;
      case 'addTask':           result = addTask(body);                  break;
      case 'updateTask':        result = updateTask(body);               break;
      case 'deleteTask':        result = deleteTask(body.id);            break;
      default:
        result = { ok: false, error: 'Ação inválida: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── SETUP (cria abas se não existirem) ─────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Transações
  let txSheet = ss.getSheetByName(SHEETS.transactions);
  if (!txSheet) {
    txSheet = ss.insertSheet(SHEETS.transactions);
    txSheet.appendRow(['id', 'type', 'desc', 'value', 'cat', 'date', 'createdAt']);
    txSheet.setFrozenRows(1);
    txSheet.getRange('1:1').setFontWeight('bold').setBackground('#1E3A5F').setFontColor('#FFFFFF');
  }

  // Tarefas
  let taskSheet = ss.getSheetByName(SHEETS.tasks);
  if (!taskSheet) {
    taskSheet = ss.insertSheet(SHEETS.tasks);
    taskSheet.appendRow(['id', 'desc', 'type', 'deadline', 'value', 'status', 'createdAt']);
    taskSheet.setFrozenRows(1);
    taskSheet.getRange('1:1').setFontWeight('bold').setBackground('#0F2438').setFontColor('#FFFFFF');
  }

  return { ok: true, message: 'Planilhas configuradas com sucesso!' };
}

// ─── TRANSACTIONS ────────────────────────────────────────────
function getTransactions() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.transactions);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, data: [] };

  const headers = rows[0];
  const data = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse(); // mais recente primeiro

  return { ok: true, data };
}

function addTransaction(body) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.transactions);
  const id = Date.now();
  sheet.appendRow([
    id,
    body.type,
    body.desc,
    parseFloat(body.value),
    body.cat,
    body.date,
    new Date().toISOString()
  ]);
  return { ok: true, id };
}

function deleteTransaction(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.transactions);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Registro não encontrado' };
}

// ─── TASKS ───────────────────────────────────────────────────
function getTasks() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.tasks);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, data: [] };

  const headers = rows[0];
  const data = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse();

  return { ok: true, data };
}

function addTask(body) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.tasks);
  const id = Date.now();
  sheet.appendRow([
    id,
    body.desc,
    body.type,
    body.deadline || '',
    body.value ? parseFloat(body.value) : '',
    body.status || 'pending',
    new Date().toISOString()
  ]);
  return { ok: true, id };
}

function updateTask(body) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.tasks);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const statusIdx = headers.indexOf('status') + 1;

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(body.id)) {
      sheet.getRange(i + 1, statusIdx).setValue(body.status);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Tarefa não encontrada' };
}

function deleteTask(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEETS.tasks);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Tarefa não encontrada' };
}
