// ╔══════════════════════════════════════════════════════════╗
// ║   FAMÍLIA APP — BACKEND                                  ║
// ║   1. Cole este código no Google Apps Script              ║
// ║   2. Salve (Ctrl+S)                                      ║
// ║   3. Implantar > Gerenciar implantações > ✏️ editar       ║
// ║   4. Versão: "Nova versão" > Implantar                   ║
// ╚══════════════════════════════════════════════════════════╝

// ★ SUBSTITUA PELO ID DA SUA PLANILHA ★
// (está na URL: docs.google.com/spreadsheets/d/SEU_ID_AQUI/edit)
var SPREADSHEET_ID = '174UmeWX3kmj9qjl7Z3I8hACNpx1pjcWYsl1_wBFzgxM';

// ════════════════════════════════════════════════════════════
//  ENTRY POINT — tudo via GET para evitar problemas de CORS
// ════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var params = e.parameter || {};
    var action = params.action || '';

    // paginaInquilino retorna HTML — tratamento especial
    if (action === 'paginaInquilino') {
      return paginaInquilino(e);
    }
    if (action === 'paginaDivida') {
      return paginaDivida(e);
    }

    var body = {};
    if (params.payload) {
      try { body = JSON.parse(decodeURIComponent(params.payload)); } catch(ex) {}
    }
    var result = doGetInternal(action, body);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGetInternal(action, body) {
  switch (action) {
    case 'getTransactions':   return getTransactions();
    case 'addTransaction':    return addTransaction(body);
    case 'deleteTransaction': return deleteTransaction(body.id);
    case 'getTasks':          return getTasks();
    case 'addTask':           return addTask(body);
    case 'updateTask':        return updateTask(body);
    case 'deleteTask':        return deleteTask(body.id);
    case 'getDocs':           return getDocs();
    case 'addDoc':            return addDoc(body);
    case 'deleteDoc':         return deleteDoc(body.id);
    case 'getRecur':          return getRecur();
    case 'addRecur':          return addRecur(body);
    case 'deleteRecur':       return deleteRecur(body.id);
    case 'uploadComprovante':          return uploadComprovante(body);
    case 'criarPlanilhaContrato':      return criarPlanilhaContrato(body);
    case 'registrarPagamentoContrato': return registrarPagamentoContrato(body);
    case 'criarFichaAluguel':          return criarFichaAluguel(body);
    case 'registrarPagamentoAluguel':  return registrarPagamentoAluguel(body);
    // CRUD genérico — dados centralizados
    case 'salvarContrato':    return salvarItemSheet('Contratos', body);
    case 'salvarContratoChunk': return salvarContratoChunk(body);
    case 'deletarContrato':   return deletarItemSheet('Contratos', body.id);
    case 'getContratos':      return getItemsSheet('Contratos');
    case 'salvarPagador':     return salvarItemSheet('Pagadores', body);
    case 'deletarPagador':    return deletarItemSheet('Pagadores', body.id);
    case 'getPagadores':      return getItemsSheet('Pagadores');
    case 'salvarAluguel':     return salvarItemSheet('Aluguéis', body);
    case 'deletarAluguel':    return deletarItemSheet('Aluguéis', body.id);
    case 'getAlugueis':       return getItemsSheet('Aluguéis');
    case 'salvarCGasto':      return salvarItemSheet('GastosCartao', body);
    case 'getCGastos':        return getItemsSheet('GastosCartao');
    case 'paginaInquilino':   return { ok: false, error: 'Use GET direto para paginaInquilino' };
    case 'salvarFcmToken':    return salvarFcmToken(body);
    case 'getValorFaturaCartao': return getValorFaturaCartao(body);
    case 'salvarRecibo':      return salvarItemSheet('Recibos', body);
    case 'getRecibos':        return getItemsSheet('Recibos');
    case 'salvarCartao':      return salvarItemSheet('Cartoes', body);
    case 'getCartoes':        return getItemsSheet('Cartoes');
    case 'deletarItemSheet':  return deletarItemSheet(body.aba, body.id);
    case 'salvarContaBanco':  return salvarItemSheet('ContasBanco', body);
    case 'getContasBanco':    return getItemsSheet('ContasBanco');
    case 'salvarExtrato':     return salvarExtratoSheet(body);
    case 'getExtratos':       return getExtratosSheet(body.contaId);
    case 'deletarExtrato':    return deletarExtratoSheet(body.contaId, body.extratoId);
    case 'salvarDivida':      return salvarDividaEstruturada(body);
    case 'getDividas':        return getDividasEstruturadas();
    case 'calcularSaldoDivida': return calcularSaldoDivida(body);
    case 'calcularLedgerDivida': return recalcularLedgerDivida(body);
    case 'deletarDivida':     return deletarDivida(body.id);
    case 'getDividasSheetUrl': return getDividasSheetUrl();
    case 'ping':              return { ok: true, msg: 'pong' };
    default:                  return { ok: false, error: 'Ação inválida: ' + action };
  }
}

function doPost(e) {
  try {
    var params = e.parameter || {};
    var action = params.action || '';
    var body = {};

    // POST body pode vir como JSON no postData
    if (e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch(ex) {}
    }
    // Também aceita payload na URL
    if (!Object.keys(body).length && params.payload) {
      body = JSON.parse(decodeURIComponent(params.payload));
    }

    var result;
    switch (action) {
      case 'uploadComprovante': result = uploadComprovante(body); break;
      default: result = doGetInternal(action, body);
    }
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok:false, error:err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ════════════════════════════════════════════════════════════
//  HELPER
// ════════════════════════════════════════════════════════════
function ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function fmtDate(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' +
      String(v.getMonth() + 1).padStart(2, '0') + '-' +
      String(v.getDate()).padStart(2, '0');
  }
  return v ? String(v) : '';
}

function sheetRows(name) {
  var sheet = ss().getSheetByName(name);
  if (!sheet) return { sheet: null, headers: [], rows: [] };
  var all = sheet.getDataRange().getValues();
  return {
    sheet:   sheet,
    headers: all[0] || [],
    rows:    all.slice(1)
  };
}

// ════════════════════════════════════════════════════════════
//  TRANSAÇÕES
// ════════════════════════════════════════════════════════════
function getTransactions() {
  var r = sheetRows('Transações');
  if (!r.sheet || !r.rows.length) return { ok: true, data: [] };
  var data = r.rows
    .map(function(row) {
      var obj = {};
      r.headers.forEach(function(h, i) {
        obj[h] = (h === 'date') ? fmtDate(row[i]) : row[i];
      });
      return obj;
    })
    .filter(function(obj) {
      if(obj.type !== 'inc' && obj.type !== 'exp') return false;
      // Excluir lançamentos automáticos da recorrência (recurId sem fromAgenda)
      var rid = String(obj.recurId || '').trim();
      var fa  = obj.fromAgenda === 'true' || obj.fromAgenda === true;
      if(rid !== '' && !fa) return false;
      return true;
    })
    .map(function(obj) {
      obj.fromAgenda = (obj.fromAgenda === 'true' || obj.fromAgenda === true);
      return obj;
    })
    .reverse();
  return { ok: true, data: data };
}

function addTransaction(body) {
  var sheet = ss().getSheetByName('Transações');
  if (!sheet) {
    sheet = ss().insertSheet('Transações');
    sheet.appendRow(['id','type','desc','value','cat','date','createdAt','fromAgenda','recurId']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#1E3A5F').setFontColor('#FFF');
  }
  // Usar o id enviado pelo cliente se disponível, para manter consistência no sync
  var id = body.id ? String(body.id) : String(Date.now());
  sheet.appendRow([id, body.type, body.desc, parseFloat(body.value), body.cat, body.date, new Date().toISOString(), body.fromAgenda ? 'true' : '', body.recurId ? String(body.recurId) : '']);
  return { ok: true, id: id };
}

function deleteTransaction(id) {
  var r = sheetRows('Transações');
  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(id)) {
      r.sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Não encontrado' };
}

// ════════════════════════════════════════════════════════════
//  TAREFAS
// ════════════════════════════════════════════════════════════
function getTasks() {
  var r = sheetRows('Tarefas');
  if (!r.sheet || !r.rows.length) return { ok: true, data: [] };
  var data = r.rows.map(function(row) {
    var obj = {};
    r.headers.forEach(function(h, i) {
      obj[h] = (h === 'deadline') ? fmtDate(row[i]) : row[i];
    });
    if (!obj.comprovUrl) obj.comprovUrl = '';
    if (!obj.cat)       obj.cat = '';
    if (!obj.recurId)   obj.recurId = '';
    return obj;
  }).reverse();
  return { ok: true, data: data };
}

function addTask(body) {
  var sheet = ss().getSheetByName('Tarefas');
  if (!sheet) {
    sheet = ss().insertSheet('Tarefas');
    sheet.appendRow(['id','desc','type','deadline','value','status','createdAt','cat','recurId']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#0F2438').setFontColor('#FFF');
  }

  // Garantir colunas cat, recurId e time existem
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('cat') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('cat');
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  if (headers.indexOf('recurId') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('recurId');
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  if (headers.indexOf('time') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('time');
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }

  // Verificar duplicata antes de inserir
  var descIdx2    = headers.indexOf('desc');
  var deadlineIdx2= headers.indexOf('deadline');
  var recurIdx2   = headers.indexOf('recurId');
  var allRows     = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow()-1, headers.length).getValues() : [];
  var newMo       = String(body.deadline || '').substring(0,7);
  var newDesc     = String(body.desc || '').trim().toLowerCase();
  var newRid      = String(body.recurId || '');

  for (var ri = 0; ri < allRows.length; ri++) {
    var rDesc = String(allRows[ri][descIdx2] || '').trim().toLowerCase();
    var rDl   = String(allRows[ri][deadlineIdx2] || '').substring(0,7);
    var rRid  = String(allRows[ri][recurIdx2]  || '');
    // Duplicata: mesmo desc + mesmo mês
    if (rDesc === newDesc && rDl === newMo) {
      return { ok: true, id: allRows[ri][0], duplicate: true };
    }
  }

  // Montar linha com posições corretas
  var id  = Date.now();
  var row = new Array(headers.length).fill('');
  var set = function(col, val){ var i=headers.indexOf(col); if(i>=0) row[i]=val; };
  set('id',        id);
  set('desc',      body.desc || '');
  set('type',      body.type || 'exp');
  set('deadline',  body.deadline || '');
  set('value',     body.value ? parseFloat(body.value) : '');
  set('status',    body.status || 'pend');
  set('createdAt', new Date().toISOString());
  set('cat',       body.cat || '');
  set('recurId',   newRid);
  set('time',      body.time || '');

  sheet.appendRow(row);
  return { ok: true, id: id };
}

function updateTask(body) {
  var r = sheetRows('Tarefas');
  var statusIdx   = r.headers.indexOf('status');
  var valueIdx    = r.headers.indexOf('value');
  var catIdx      = r.headers.indexOf('cat');
  var comprovIdx  = r.headers.indexOf('comprovUrl');
  var timeIdx     = r.headers.indexOf('time');
  var deadlineIdx = r.headers.indexOf('deadline');

  // Criar coluna comprovUrl se não existir
  if (comprovIdx === -1 && body.comprovUrl) {
    r.sheet.getRange(1, r.headers.length + 1).setValue('comprovUrl');
    comprovIdx = r.headers.length;
    r.headers.push('comprovUrl');
  }
  // Criar coluna time se não existir
  if (timeIdx === -1 && body.time !== undefined) {
    r.sheet.getRange(1, r.headers.length + 1).setValue('time');
    timeIdx = r.headers.length;
    r.headers.push('time');
  }

  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(body.id)) {
      // Atualizar status
      if (body.status !== undefined && statusIdx > -1) {
        r.sheet.getRange(i + 2, statusIdx + 1).setValue(body.status);
      }
      // Atualizar valor
      if (body.value !== undefined && body.value !== null && valueIdx > -1) {
        r.sheet.getRange(i + 2, valueIdx + 1).setValue(parseFloat(body.value));
      }
      // Atualizar categoria
      if (body.cat !== undefined && body.cat !== null && catIdx > -1) {
        r.sheet.getRange(i + 2, catIdx + 1).setValue(body.cat);
      }
      // Atualizar comprovante
      if (body.comprovUrl && comprovIdx > -1) {
        r.sheet.getRange(i + 2, comprovIdx + 1).setValue(body.comprovUrl);
      }
      // Atualizar horário (ex: consulta remarcada)
      if (body.time !== undefined && timeIdx > -1) {
        r.sheet.getRange(i + 2, timeIdx + 1).setValue(body.time);
      }
      // Atualizar data (remarcação completa)
      if (body.deadline !== undefined && body.deadline !== null && deadlineIdx > -1) {
        r.sheet.getRange(i + 2, deadlineIdx + 1).setValue(body.deadline);
      }
      return { ok: true };
    }
  }
  // Task não encontrada — tentar criar (pode ser task nova ainda não sincronizada)
  return { ok: false, error: 'Tarefa não encontrada' };
}

function deleteTask(id) {
  var r = sheetRows('Tarefas');
  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(id)) {
      r.sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Não encontrado' };
}

// ════════════════════════════════════════════════════════════
//  DOCUMENTOS
// ════════════════════════════════════════════════════════════
function getDocs() {
  var r = sheetRows('Documentos');
  if (!r.sheet || !r.rows.length) return { ok: true, data: [] };
  var data = r.rows.map(function(row) {
    var obj = {};
    r.headers.forEach(function(h, i) { obj[String(h).trim()] = row[i]; });
    return {
      id:       String(obj.id       || ''),
      name:     String(obj.name     || ''),
      url:      String(obj.url      || ''),
      type:     String(obj.type     || 'doc'),
      pasta:    String(obj.pasta    || obj.group || 'Outros'),
      subpasta: String(obj.subpasta || '')
    };
  }).filter(function(d) { return d.name && d.url; }).reverse();
  return { ok: true, data: data };
}

function addDoc(body) {
  var sheet = ss().getSheetByName('Documentos');
  if (!sheet) {
    sheet = ss().insertSheet('Documentos');
    sheet.appendRow(['id','name','url','type','pasta','subpasta','createdAt']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#1E2D40').setFontColor('#FFF');
  }
  var id = Date.now();
  sheet.appendRow([
    id,
    body.name     || '',
    body.url      || '',
    body.type     || 'doc',
    body.pasta    || 'Outros',
    body.subpasta || '',
    new Date().toISOString()
  ]);
  return { ok: true, id: id };
}

function deleteDoc(id) {
  var r = sheetRows('Documentos');
  if (!r.sheet) return { ok: false, error: 'Aba não encontrada' };
  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(id)) {
      r.sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Não encontrado' };
}

// ════════════════════════════════════════════════════════════
//  RECORRÊNCIA
// ════════════════════════════════════════════════════════════
function getRecur() {
  var r = sheetRows('Recorrentes');
  if (!r.sheet || !r.rows.length) return { ok: true, data: [] };
  var data = r.rows.map(function(row) {
    var obj = {};
    r.headers.forEach(function(h, i) { obj[String(h).trim()] = row[i]; });
    return {
      id:    String(obj.id    || ''),
      type:  String(obj.type  || 'exp'),
      desc:  String(obj.desc  || ''),
      value: parseFloat(obj.value || 0),
      cat:   String(obj.cat   || ''),
      date:  fmtDate(obj.date) || String(obj.date || '')
    };
  }).filter(function(r) { return r.desc; });
  return { ok: true, data: data };
}

function addRecur(body) {
  var ss2 = ss();
  var sheet = ss2.getSheetByName('Recorrentes');
  if (!sheet) {
    sheet = ss2.insertSheet('Recorrentes');
    sheet.appendRow(['id','type','desc','value','cat','date','updatedAt','cartaoId']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#162030').setFontColor('#FFF');
  }
  var id = String(body.id || Date.now());
  var now = new Date().toISOString();
  var cartaoId = body.cartaoId || '';

  // Verificar se já existe linha com esse id — se sim, atualizar
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      // Atualizar linha existente
      sheet.getRange(i + 1, 1, 1, 8).setValues([[
        id,
        body.type  || data[i][1],
        body.desc  || data[i][2],
        parseFloat(body.value) || 0,
        body.cat   || data[i][4],
        body.date  || data[i][5],
        now,
        cartaoId
      ]]);
      Logger.log('addRecur: atualizado id=' + id + ' valor=' + body.value + ' cat=' + body.cat);
      return { ok: true, id: id, updated: true };
    }
  }

  // Não existe — inserir nova linha
  sheet.appendRow([id, body.type, body.desc, parseFloat(body.value)||0, body.cat, body.date, now, cartaoId]);
  Logger.log('addRecur: criado id=' + id + ' valor=' + body.value);
  return { ok: true, id: id, created: true };
}

function deleteRecur(id) {
  var r = sheetRows('Recorrentes');
  if (!r.sheet) return { ok: false, error: 'Aba não encontrada' };
  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(id)) {
      r.sheet.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Não encontrado' };
}

// ════════════════════════════════════════════════════════════
//  COMPROVANTES — upload para pasta no Google Drive
// ════════════════════════════════════════════════════════════

// ★ Opcional: defina o ID de uma pasta específica no Drive para comprovantes
// Deixe vazio ('') para salvar na raiz do Drive
var COMPROVANTES_FOLDER_ID = '';

function uploadComprovante(body) {
  try {
    var folder;
    if (COMPROVANTES_FOLDER_ID) {
      folder = DriveApp.getFolderById(COMPROVANTES_FOLDER_ID);
    } else {
      // Buscar ou criar pasta "Comprovantes Família" na raiz
      var folders = DriveApp.getFoldersByName('Comprovantes Família');
      folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('Comprovantes Família');
    }

    // Decodificar base64
    var decoded = Utilities.base64Decode(body.data);
    var blob = Utilities.newBlob(decoded, body.mimeType, body.name);

    // Salvar no Drive
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    return { ok: true, url: url, id: file.getId(), name: file.getName() };

  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ═══════════════════════════════════════════════════
// FUNÇÃO DE LIMPEZA — executar UMA VEZ pelo editor
// do Apps Script para remover transações com data
// futura (geradas automaticamente pela versão antiga).
// Após executar, pode deletar esta função.
// ═══════════════════════════════════════════════════
function limparTransacoesAutomaticas() {
  var sheet = ss().getSheetByName('Transações');
  if (!sheet) { Logger.log('Aba Transações não encontrada'); return; }

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var dateIdx = headers.indexOf('date');
  var typeIdx = headers.indexOf('type');

  if (dateIdx === -1) { Logger.log('Coluna date não encontrada'); return; }

  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  var deletadas = 0;
  // Percorrer de baixo para cima
  for (var i = data.length - 1; i >= 1; i--) {
    var row  = data[i];
    var type = String(row[typeIdx] || '');
    var dateVal = row[dateIdx];

    // Converter para Date
    var d = dateVal instanceof Date ? dateVal : new Date(dateVal);
    d.setHours(0, 0, 0, 0);

    // Deletar se: data futura OU type inválido
    if (d > hoje || (type !== 'inc' && type !== 'exp')) {
      sheet.deleteRow(i + 1);
      deletadas++;
    }
  }
  Logger.log('Transações removidas: ' + deletadas);
}

// ════════════════════════════════════════════════════════════
//  E-MAIL DIÁRIO — RESUMO DE COMPROMISSOS
// ════════════════════════════════════════════════════════════

// ★ CONFIGURE SEU E-MAIL AQUI ★
var EMAIL_DESTINO  = 'armbr258@gmail.com';
var EMAILS_DESTINO = [EMAIL_DESTINO]; // Para adicionar mais: ['email1@gmail.com', 'email2@gmail.com']


// Normaliza qualquer valor de data para string 'YYYY-MM-DD'
function toDateStr(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  // Tem T (ISO): pegar só a parte da data
  if (s.indexOf('T') !== -1) s = s.split('T')[0];
  // Já é YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  // DD/MM/YYYY
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) {
    var p = s.split('/'); return p[2]+'-'+p[1]+'-'+p[0];
  }
  return s;
}

// Envia email para todos os destinatários configurados
function enviarParaTodos(assunto, corpo) {
  var emails = String(EMAIL_DESTINO).split(',').map(function(e){ return e.trim(); }).filter(function(e){ return e.indexOf('@')>0; });
  emails.forEach(function(email) {
    GmailApp.sendEmail(email, assunto, '', { htmlBody: corpo, name: 'Fluxo App' });
    Logger.log('Email enviado para: ' + email);
  });
}

function enviarResumoDiario() {
  var tz      = Session.getScriptTimeZone();
  var hoje    = new Date();
  var todayStr = Utilities.formatDate(hoje, tz, 'yyyy-MM-dd');
  var diaSem  = ['Domingo','Segunda','Terça','Quarta','Quinta','Sexta','Sábado'][hoje.getDay()];
  var diaFmt  = Utilities.formatDate(hoje, tz, 'dd/MM/yyyy');

  // ── Ler tarefas usando headers por nome ──────────────────
  var sheet = ss().getSheetByName('Tarefas');
  if (!sheet) { Logger.log('Aba Tarefas não encontrada'); return; }

  var all     = sheet.getDataRange().getValues();
  var headers = all[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var rows    = all.slice(1);

  Logger.log('Headers encontrados: ' + JSON.stringify(headers));
  Logger.log('Total de linhas: ' + rows.length);

  var iDesc     = headers.indexOf('desc');
  var iDeadline = headers.indexOf('deadline');
  var iStatus   = headers.indexOf('status');
  var iType     = headers.indexOf('type');
  var iValue    = headers.indexOf('value');
  var iCat      = headers.indexOf('cat');

  Logger.log('Índices — desc:'+iDesc+' deadline:'+iDeadline+' status:'+iStatus+' type:'+iType);

  var tarefas   = [];
  var atrasadas = [];

  rows.forEach(function(row, idx) {
    var status   = String(row[iStatus] || '').trim().toLowerCase();
    var deadline = toDateStr(row[iDeadline]);
    var desc     = String(row[iDesc]   || '').trim();

    Logger.log('Linha '+(idx+1)+': desc="'+desc+'" deadline="'+deadline+'" status="'+status+'"');

    if (status === 'done') return;
    if (!deadline || !desc) return;

    var item = {
      desc:     desc,
      type:     String(row[iType]  || '').trim(),
      value:    parseFloat(row[iValue] || 0),
      cat:      String(row[iCat]   || '').trim(),
      deadline: deadline
    };

    if (deadline === todayStr) {
      tarefas.push(item);
    } else if (deadline < todayStr) {
      atrasadas.push(item);
    }
  });

  Logger.log('Compromissos hoje: ' + tarefas.length);
  Logger.log('Atrasados: ' + atrasadas.length);

  if (!tarefas.length && !atrasadas.length) {
    Logger.log('Nenhum compromisso — e-mail não enviado.');
    return;
  }

  // ── Montar e-mail ────────────────────────────────────────
  function fmtBRLTotal(v) {
    return 'R$ ' + Number(v).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }

  function tipoDot(t) {
    var cor = t === 'inc' ? '#2ecc9a' : t === 'exp' ? '#FF6B6B' : '#5B9BD5';
    return '<div style="width:10px;height:10px;border-radius:50%;background:'+cor+';flex-shrink:0;margin-top:4px"></div>';
  }

  var totalHoje     = tarefas.reduce(function(s,t){   return s+(t.type!=='task'?Number(t.value||0):0);}, 0);
  var totalAtrasado = atrasadas.reduce(function(s,t){ return s+(t.type!=='task'?Number(t.value||0):0);}, 0);

  function linhaItem(t, corTexto) {
    var dl = t.deadline ? t.deadline.split('-').reverse().slice(0,2).join('/') : '';
    var valCor = t.type === 'inc' ? '#2ecc9a' : '#FF6B6B';
    return '<tr>'
      + '<td style="padding:12px 16px;border-bottom:1px solid #1e1e2e;vertical-align:top;width:22px">'
      +   tipoDot(t.type)
      + '</td>'
      + '<td style="padding:12px 8px 12px 0;border-bottom:1px solid #1e1e2e;vertical-align:top">'
      +   '<div style="font-size:14px;font-weight:700;color:'+(corTexto||'#eee')+'">'+t.desc+'</div>'
      +   '<div style="font-size:12px;color:#666;margin-top:3px">'+(t.cat||'')+(dl?' &nbsp;·&nbsp; Vence '+dl:'')+'</div>'
      + '</td>'
      + '<td style="padding:12px 16px 12px 8px;border-bottom:1px solid #1e1e2e;text-align:right;vertical-align:top;white-space:nowrap">'
      +   (t.value>0?'<span style="font-size:14px;font-weight:800;color:'+valCor+'">'+fmtBRLTotal(t.value)+'</span>':'')
      + '</td>'
      + '</tr>';
  }

  var secaoHoje = '';
  if (tarefas.length) {
    secaoHoje = '<div style="padding:0 0 4px 0">'
      + '<div style="padding:16px 20px 10px;background:#0d1117;border-bottom:1px solid #1e1e2e">'
      +   '<span style="font-size:11px;font-weight:800;color:#2ecc9a;letter-spacing:1.5px;text-transform:uppercase">COMPROMISSOS DE HOJE</span>'
      +   (totalHoje>0?' &nbsp;<span style="font-size:13px;font-weight:800;color:#FF6B6B;float:right">'+fmtBRLTotal(totalHoje)+'</span>':'')
      + '</div>'
      + '<table style="width:100%;border-collapse:collapse">'
      + tarefas.map(function(t){ return linhaItem(t, '#e8e8e8'); }).join('')
      + '</table></div>';
  }

  var secaoAtrasado = '';
  if (atrasadas.length) {
    secaoAtrasado = '<div style="padding:0 0 4px 0">'
      + '<div style="padding:16px 20px 10px;background:#1a0a0a;border-bottom:1px solid #2a1010;border-top:2px solid #FF6B6B">'
      +   '<span style="font-size:11px;font-weight:800;color:#FF6B6B;letter-spacing:1.5px;text-transform:uppercase">ATRASADOS</span>'
      +   (totalAtrasado>0?' &nbsp;<span style="font-size:13px;font-weight:800;color:#FF6B6B;float:right">'+fmtBRLTotal(totalAtrasado)+'</span>':'')
      + '</div>'
      + '<table style="width:100%;border-collapse:collapse">'
      + atrasadas.map(function(t){ return linhaItem(t, '#ffaaaa'); }).join('')
      + '</table></div>';
  }

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body style="margin:0;padding:20px;background:#0a0a0f;font-family:Arial,Helvetica,sans-serif">'
    + '<div style="max-width:560px;margin:0 auto;background:#13131f;border-radius:12px;overflow:hidden;border:1px solid #1e1e2e">'

    // Header
    + '<div style="padding:22px 20px;background:#0d1117;border-bottom:2px solid #2ecc9a">'
    +   '<div style="font-size:22px;font-weight:900;color:#fff;letter-spacing:-0.5px">Fluxo App</div>'
    +   '<div style="color:#666;margin-top:4px;font-size:13px">'+diaSem+', '+diaFmt+'</div>'
    + '</div>'

    + secaoHoje
    + secaoAtrasado

    // Footer
    + '<div style="padding:14px 20px;border-top:1px solid #1e1e2e;text-align:center">'
    +   '<span style="color:#333;font-size:11px">Resumo automatico diario · Fluxo App</span>'
    + '</div>'

    + '</div></body></html>';

  GmailApp.sendEmail(
    EMAIL_DESTINO,
    'Fluxo ' + diaFmt + (tarefas.length ? ' — ' + tarefas.length + ' compromisso' + (tarefas.length>1?'s':'') + (totalHoje>0?' · '+fmtBRLTotal(totalHoje):'') : '') +
    (atrasadas.length ? ' — ' + atrasadas.length + ' atrasado' + (atrasadas.length>1?'s':'') : ''),
    'Abra no Gmail para ver o resumo formatado.',
    { htmlBody: html, name: 'Fluxo App' }
  );

  Logger.log('✅ E-mail enviado para ' + EMAIL_DESTINO);

  // Enviar push notification junto com o email
  try {
    var nHoje     = tarefas.length;
    var nAtrasadas = atrasadas.length;
    var pushMsg = nHoje > 0
      ? nHoje + ' compromisso' + (nHoje > 1 ? 's' : '') + ' para hoje'
        + (nAtrasadas > 0 ? ' · ' + nAtrasadas + ' atrasado' + (nAtrasadas > 1 ? 's' : '') : '')
      : nAtrasadas > 0
        ? nAtrasadas + ' atrasado' + (nAtrasadas > 1 ? 's' : '')
        : 'Nenhum compromisso hoje 🎉';
    enviarPush('☀️ Fluxo — Resumo do dia', pushMsg);
    Logger.log('✅ Push enviado: ' + pushMsg);
  } catch(ep) {
    Logger.log('Push falhou: ' + ep.message);
  }
}

// ── Trigger diário às 8h ─────────────────────────────────
function criarTriggerDiario() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarResumoDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarResumoDiario').timeBased().everyDays(1).atHour(8).create();
  Logger.log('✅ Trigger criado — e-mail todo dia às 8h.');
}

// ── Debug ────────────────────────────────────────────────
function debugEmail() {
  Logger.log('=== DEBUG EMAIL ===');
  enviarResumoDiario();
}

// ════════════════════════════════════════════════════════════
//  ALUGUÉIS — PLANILHA GOOGLE SHEETS
// ════════════════════════════════════════════════════════════

function criarFichaAluguel(body) {
  var ssObj = ss();
  var nomeAba = 'Aluguel - ' + String(body.nome || '').substring(0, 20);

  // Criar aba se não existir
  var sheet = ssObj.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ssObj.insertSheet(nomeAba);
    // Cabeçalho
    sheet.getRange('A1').setValue('FICHA DE ALUGUEL').setFontSize(14).setFontWeight('bold');
    sheet.getRange('A2').setValue('Inquilino: ' + body.nome);
    sheet.getRange('A3').setValue('Referência: ' + (body.ref || ''));
    sheet.getRange('A4').setValue('Telefone: ' + (body.tel || ''));
    sheet.getRange('A5').setValue('Vencimento: dia ' + (body.dia || ''));
    sheet.getRange('A6').setValue('Valor base: R$ ' + parseFloat(body.valorBase || 0).toFixed(2));
    sheet.getRange('A8:E8').setValues([['Mês', 'Valor', 'Data Pagamento', 'Status', 'Código']]);
    sheet.getRange('A8:E8').setFontWeight('bold').setBackground('#1E3A5F').setFontColor('#FFFFFF');
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 150);
  }
  return { ok: true };
}

function registrarPagamentoAluguel(body) {
  var ssObj = ss();
  var nomeAba = 'Aluguel - ' + String(body.nome || '').substring(0, 20);
  var sheet = ssObj.getSheetByName(nomeAba);

  // Criar ficha se não existir
  if (!sheet) criarFichaAluguel(body);
  sheet = ssObj.getSheetByName(nomeAba);
  if (!sheet) return { ok: false, error: 'Aba não encontrada' };

  // Verificar se mês já registrado
  var dados = sheet.getDataRange().getValues();
  var mo = String(body.mesAno || '');
  var partes = mo.split('-');
  var meses = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var mesLabel = meses[parseInt(partes[1])-1] + '/' + partes[0];

  for (var i = 8; i < dados.length; i++) {
    if (String(dados[i][0]) === mesLabel) {
      // Atualizar linha existente
      sheet.getRange(i+1, 2, 1, 4).setValues([[
        parseFloat(body.valor || 0),
        body.dataPgto || '',
        'PAGO',
        body.codigo || ''
      ]]);
      sheet.getRange(i+1, 1, 1, 5).setBackground('#d4edda');
      return { ok: true };
    }
  }

  // Inserir nova linha
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, 5).setValues([[
    mesLabel,
    parseFloat(body.valor || 0),
    body.dataPgto || '',
    'PAGO',
    body.codigo || ''
  ]]);
  sheet.getRange(lastRow, 1, 1, 5).setBackground('#d4edda');

  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  LIMPEZA DE TAREFAS DUPLICADAS
//  Execute uma vez para limpar duplicatas na planilha
// ════════════════════════════════════════════════════════════
function limparTarefasDuplicadas() {
  var r = sheetRows('Tarefas');
  if (!r.sheet || !r.rows.length) { Logger.log('Sem dados'); return; }

  var headers    = r.headers;
  var descIdx    = headers.indexOf('desc');
  var deadlineIdx= headers.indexOf('deadline');
  var recurIdx   = headers.indexOf('recurId');
  var statusIdx  = headers.indexOf('status');

  var seen    = {}; // chave: desc|mes → rowIndex
  var toDelete= [];

  r.rows.forEach(function(row, i) {
    var desc     = String(row[descIdx]  || '').trim().toLowerCase();
    var deadline = String(row[deadlineIdx] || '');
    // Normalizar deadline para YYYY-MM
    var mo = deadline.length >= 7 ? deadline.substring(0,7) : deadline;
    var key = desc + '|' + mo;
    var recurId = String(row[recurIdx] || '').trim();

    if (!seen[key]) {
      seen[key] = { rowIndex: i, recurId: recurId };
    } else {
      // Duplicata — manter a que tem recurId
      if (recurId && !seen[key].recurId) {
        // A nova tem recurId, a antiga não — deletar a antiga
        toDelete.push(seen[key].rowIndex);
        seen[key] = { rowIndex: i, recurId: recurId };
      } else {
        // Deletar a nova (duplicata sem recurId, ou ambas têm)
        toDelete.push(i);
      }
    }
  });

  // Deletar de baixo para cima (para não deslocar índices)
  toDelete.sort(function(a,b){ return b-a; });
  toDelete.forEach(function(i) {
    r.sheet.deleteRow(i + 2); // +2 porque header é row 1
  });

  Logger.log('Tarefas duplicadas removidas: ' + toDelete.length);
}

// ════════════════════════════════════════════════════════════
//  CONTRATOS — GOOGLE SHEETS
// ════════════════════════════════════════════════════════════

function criarPlanilhaContrato(body) {
  var ssObj  = ss();
  var inq    = String(body.inqNome || '').substring(0, 18);
  var imovel = String(body.imovel  || '').substring(0, 10);
  var nomeAba = 'Contrato - ' + inq;

  var sheet = ssObj.getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ssObj.insertSheet(nomeAba);
    // Cabeçalho informativo
    sheet.getRange('A1').setValue('CONTROLE DE PAGAMENTOS — ' + body.inqNome).setFontSize(13).setFontWeight('bold');
    sheet.getRange('A2').setValue('Imóvel: ' + (body.imovel || '') + ' · ' + (body.end || ''));
    sheet.getRange('A3').setValue('Locador: ' + (body.locNome || ''));
    sheet.getRange('A4').setValue('Início: ' + (body.inicio || '') + ' · Valor: R$ ' + parseFloat(body.valor||0).toFixed(2) + ' · Venc. dia ' + (body.diaPgto || '10'));
    sheet.getRange('A5').setValue('');

    // Header da tabela
    var hdr = ['Mês', 'Valor (R$)', 'Data Pagamento', 'Status', 'Código Recibo', 'Observação'];
    sheet.getRange(6, 1, 1, hdr.length).setValues([hdr])
      .setFontWeight('bold')
      .setBackground('#0F2438')
      .setFontColor('#FFFFFF');

    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 200);
    sheet.setFrozenRows(6);

    // Pré-preencher meses do contrato como PENDENTE
    var dur = parseInt(body.duracao) || 12;
    var startDate = body.inicio ? new Date(body.inicio + 'T12:00:00') : new Date();
    var meses_br = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    for (var m = 0; m < dur; m++) {
      var d = new Date(startDate.getFullYear(), startDate.getMonth() + m, 1);
      var mesLabel = meses_br[d.getMonth()] + '/' + d.getFullYear();
      var row = [mesLabel, parseFloat(body.valor||0), '', 'PENDENTE', '', ''];
      sheet.getRange(7 + m, 1, 1, 6).setValues([row]);
      // Formatar coluna Mês como texto puro para evitar conversão de data
      sheet.getRange(7 + m, 1).setNumberFormat('@STRING@');
      sheet.getRange(7 + m, 4).setBackground('#FFF3CD').setFontColor('#856404');
    }
  }
  return { ok: true };
}

function registrarPagamentoContrato(body) {
  var ssObj  = ss();
  var nomeAba = 'Contrato - ' + String(body.inqNome || '').substring(0, 18);
  var sheet  = ssObj.getSheetByName(nomeAba);
  if (!sheet) {
    criarPlanilhaContrato(body);
    sheet = ssObj.getSheetByName(nomeAba);
    if (!sheet) return { ok: false, error: 'Aba não criada' };
  }

  var meses_br = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
  var partes = String(body.mesAno || '').split('-');
  var mesLabel = meses_br[parseInt(partes[1]||1)-1] + '/' + (partes[0]||'');

  var dados = sheet.getDataRange().getValues();
  for (var i = 6; i < dados.length; i++) {
    if (String(dados[i][0]) === mesLabel) {
      sheet.getRange(i+1, 2, 1, 5).setValues([[
        parseFloat(body.valor || 0),
        body.dataPgto || '',
        'PAGO',
        body.codigo   || '',
        ''
      ]]);
      sheet.getRange(i+1, 1, 1, 6).setBackground('#D4EDDA').setFontColor('#155724');
      return { ok: true };
    }
  }
  // Mês não encontrado — inserir nova linha
  var lastRow = sheet.getLastRow() + 1;
  sheet.getRange(lastRow, 1, 1, 6).setValues([[
    mesLabel, parseFloat(body.valor||0), body.dataPgto||'', 'PAGO', body.codigo||'', ''
  ]]).setBackground('#D4EDDA').setFontColor('#155724');
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  CRUD GENÉRICO — qualquer aba da planilha
// ════════════════════════════════════════════════════════════

function salvarContratoChunk(body) {
  // Recebe chunks de um contrato grande e faz merge
  Logger.log('salvarContratoChunk: id=' + body.id + ' chunk=' + body._chunk);
  if (!body || !body.id) return { ok: false, error: 'ID ausente' };
  
  var sheet = ss().getSheetByName('Contratos');
  if (!sheet) {
    sheet = ss().insertSheet('Contratos');
    sheet.appendRow(['id', 'json', 'updatedAt']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#0F2438').setFontColor('#FFF');
  }
  
  var id = String(body.id);
  var dados = sheet.getDataRange().getValues();
  var existing = null;
  var existingRow = -1;
  
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === id) {
      try { existing = JSON.parse(String(dados[i][1])); } catch(e) { existing = {}; }
      existingRow = i + 1;
      break;
    }
  }
  
  // Merge: combinar dados existentes com novo chunk
  var merged = Object.assign({}, existing || {}, body);
  delete merged._chunk; // remover flag de chunk
  
  var json = JSON.stringify(merged);
  var now  = new Date().toISOString();
  
  if (existingRow > 0) {
    sheet.getRange(existingRow, 2, 1, 2).setValues([[json, now]]);
  } else {
    sheet.appendRow([id, json, now]);
  }
  
  return { ok: true, id: id };
}

function salvarItemSheet(nomeAba, body) {
  Logger.log('salvarItemSheet: aba=' + nomeAba + ' id=' + body.id);
  if (!body || !body.id) return { ok: false, error: 'Body ou ID ausente' };
  var sheet = ss().getSheetByName(nomeAba);
  if (!sheet) {
    sheet = ss().insertSheet(nomeAba);
    sheet.appendRow(['id','json','updatedAt']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#0F2438').setFontColor('#FFF');
  }
  var dados = sheet.getDataRange().getValues();
  var id    = String(body.id || '');
  var json  = JSON.stringify(body);
  var now   = new Date().toISOString();

  // Verificar se já existe (atualizar)
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === id) {
      sheet.getRange(i+1, 2, 1, 2).setValues([[json, now]]);
      return { ok: true, updated: true };
    }
  }
  // Inserir novo
  sheet.appendRow([id, json, now]);
  return { ok: true, created: true };
}

function deletarItemSheet(nomeAba, id) {
  var sheet = ss().getSheetByName(nomeAba);
  if (!sheet) return { ok: true };
  var dados = sheet.getDataRange().getValues();
  for (var i = dados.length - 1; i >= 1; i--) {
    if (String(dados[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: true };
}

function getItemsSheet(nomeAba) {
  var sheet = ss().getSheetByName(nomeAba);
  if (!sheet) return { ok: true, data: [] };
  var dados = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < dados.length; i++) {
    if (!dados[i][1]) continue;
    try { result.push(JSON.parse(String(dados[i][1]))); } catch(e) {}
  }
  return { ok: true, data: result };
}

// ════════════════════════════════════════════════════════════
//  DÍVIDAS — EXTRATO CRONOLÓGICO de empréstimos e dívidas
//  Cria automaticamente uma planilha Google SEPARADA com:
//   • Aba "Dívidas"            — cabeçalho + saldo devedor calculado
//   • Aba "Movimentos_Dividas" — extrato completo (aportes, pagamentos,
//                                 despesas em nome de terceiro), com data
//                                 livre (inclusive retroativa)
//   • Aba "Resumo"             — total que você deve / devem a você
//
//  O saldo devedor é recalculado percorrendo CRONOLOGICAMENTE todos os
//  movimentos da dívida, aplicando juros entre cada evento — exatamente
//  como um extrato bancário manual, mas automático.
// ════════════════════════════════════════════════════════════

// Para usar uma planilha já existente, cole o ID aqui.
// Deixe em branco para o sistema criar uma planilha nova automaticamente
// na primeira vez (o ID gerado é salvo nas Propriedades do Script).
var DIVIDAS_SPREADSHEET_ID = '';

function dividasSS() {
  if (DIVIDAS_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(DIVIDAS_SPREADSHEET_ID);
  }
  var props = PropertiesService.getScriptProperties();
  var storedId = props.getProperty('DIVIDAS_SS_ID');
  if (storedId) {
    try { return SpreadsheetApp.openById(storedId); }
    catch (e) { /* planilha foi excluída — recriar abaixo */ }
  }
  var nova = SpreadsheetApp.create('Fluxo — Controle de Dívidas e Empréstimos');
  props.setProperty('DIVIDAS_SS_ID', nova.getId());
  return nova;
}

function garantirEstruturaDividas() {
  var sp = dividasSS();

  // ── Aba "Dívidas" ──────────────────────────────────────
  var sh = sp.getSheetByName('Dívidas');
  var headers = ['id','tipo','pagadorId','pagador','descricao','valorOriginal','dataOriginal',
                 'metodoJuros','taxaMensal','indexador','valorPrincipalAtual','dataBaseAtual',
                 'diasEmAtraso','saldoDevedorEstimado','status','criadoEm','atualizadoEm'];
  if (!sh) sh = sp.insertSheet('Dívidas');
  var atuais = sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0];
  var headersOk = headers.every(function(h,i){ return atuais[i]===h; });
  if (!headersOk) {
    sh.clear();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#6D28D9').setFontColor('#FFF');
    sh.setColumnWidth(4,140); sh.setColumnWidth(5,220);
    sh.getRange('F2:F2000').setNumberFormat('R$ #,##0.00');
    sh.getRange('K2:K2000').setNumberFormat('R$ #,##0.00');
    sh.getRange('N2:N2000').setNumberFormat('R$ #,##0.00');
    sh.getRange('G2:G2000').setNumberFormat('dd/mm/yyyy');
    sh.getRange('L2:L2000').setNumberFormat('dd/mm/yyyy');
    sh.getRange('P2:Q2000').setNumberFormat('dd/mm/yyyy hh:mm');
    var rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('pendente').setBackground('#FCE4E4').setFontColor('#C0392B').setRanges([sh.getRange('O2:O2000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('parcial').setBackground('#FEF6E0').setFontColor('#B7791F').setRanges([sh.getRange('O2:O2000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('pago').setBackground('#E3FCEF').setFontColor('#1E7E45').setRanges([sh.getRange('O2:O2000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('devo').setBackground('#FFF4E5').setFontColor('#C2680C').setRanges([sh.getRange('B2:B2000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('recebo').setBackground('#E5F3FF').setFontColor('#1565C0').setRanges([sh.getRange('B2:B2000')]).build()
    ];
    sh.setConditionalFormatRules(rules);
  }

  // ── Aba "Movimentos_Dividas" — extrato cronológico ────
  var shM = sp.getSheetByName('Movimentos_Dividas');
  var headersM = ['id','dividaId','pagador','tipo','data','valor','obs','jurosNoPeriodo','saldoApos'];
  if (!shM) shM = sp.insertSheet('Movimentos_Dividas');
  var atuaisM = shM.getRange(1,1,1,Math.max(1,shM.getLastColumn())).getValues()[0];
  if (!headersM.every(function(h,i){ return atuaisM[i]===h; })) {
    shM.clear();
    shM.getRange(1,1,1,headersM.length).setValues([headersM]);
    shM.setFrozenRows(1);
    shM.getRange(1,1,1,headersM.length).setFontWeight('bold').setBackground('#0F9D58').setFontColor('#FFF');
    shM.getRange('E2:E3000').setNumberFormat('dd/mm/yyyy');
    shM.getRange('F2:F3000').setNumberFormat('R$ #,##0.00');
    shM.getRange('H2:I3000').setNumberFormat('R$ #,##0.00');
    var rulesM = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('aporte').setBackground('#FFF4E5').setFontColor('#C2680C').setRanges([shM.getRange('D2:D3000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('pagamento').setBackground('#E3FCEF').setFontColor('#1E7E45').setRanges([shM.getRange('D2:D3000')]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('despesa').setBackground('#E5F3FF').setFontColor('#1565C0').setRanges([shM.getRange('D2:D3000')]).build()
    ];
    shM.setConditionalFormatRules(rulesM);
  }

  // ── Aba "Resumo" — dashboard com totais ────────────────
  var shS = sp.getSheetByName('Resumo');
  if (!shS) {
    shS = sp.insertSheet('Resumo', 0);
    shS.getRange('A1').setValue('📊 Resumo — Dívidas e Empréstimos').setFontWeight('bold').setFontSize(16);
    shS.getRange('A3').setValue('💸 Total que você DEVE (ativo)');
    shS.getRange('B3').setFormula('=SUMPRODUCT(IFERROR((Dívidas!B2:B2000="devo")*(Dívidas!O2:O2000<>"pago")*Dívidas!N2:N2000,0))');
    shS.getRange('A4').setValue('💰 Total que DEVEM a você (ativo)');
    shS.getRange('B4').setFormula('=SUMPRODUCT(IFERROR((Dívidas!B2:B2000="recebo")*(Dívidas!O2:O2000<>"pago")*Dívidas!N2:N2000,0))');
    shS.getRange('A5').setValue('⚖️ Saldo líquido (a receber - a pagar)');
    shS.getRange('B5').setFormula('=B4-B3');
    shS.getRange('B3:B5').setNumberFormat('R$ #,##0.00').setFontWeight('bold').setFontSize(13);
    shS.getRange('A3:A5').setFontWeight('600');
    shS.setColumnWidth(1,280); shS.setColumnWidth(2,160);
    shS.getRange('A1:B1').setBackground('#6D28D9').setFontColor('#FFF');
    shS.getRange('A7').setValue('Atualizado automaticamente conforme você usa o app Fluxo.').setFontColor('#888888').setFontStyle('italic');
  }

  return { dividas: sh, movimentos: shM, resumo: shS };
}

// ════════════════════════════════════════════════════════════
//  MOTOR DE CÁLCULO — percorre os movimentos em ordem
//  cronológica, capitalizando juros entre cada evento.
// ════════════════════════════════════════════════════════════

function mesesEntreDatasGS(d1Iso, d2Iso) {
  var d1 = new Date(d1Iso + 'T12:00:00');
  var d2 = new Date(d2Iso + 'T12:00:00');
  return Math.max(0, (d2 - d1) / (1000*60*60*24*30));
}

// Calcula juros de um segmento de tempo sobre um saldo, conforme método.
// Retorna {ok, saldoFinal, juros, error}
function calcularJurosSegmento(saldo, metodoJuros, taxaMensal, indexador, dataIniIso, dataFimIso) {
  var meses = mesesEntreDatasGS(dataIniIso, dataFimIso);
  if (meses <= 0 || saldo <= 0) {
    return { ok: true, saldoFinal: saldo, juros: 0 };
  }
  if (metodoJuros === 'fixo_composto') {
    var saldoFinal = saldo * Math.pow(1 + (taxaMensal/100), meses);
    return { ok: true, saldoFinal: saldoFinal, juros: saldoFinal - saldo };
  }
  if (metodoJuros === 'indexador' || metodoJuros === 'indexador_mais_taxa') {
    var fator = buscarIndexadorAcumulado(indexador, dataIniIso, dataFimIso);
    if (fator === null) {
      return { ok: false, error: 'Não foi possível obter o índice ' + indexador + ' para ' + dataIniIso + ' → ' + dataFimIso };
    }
    var saldoIdx = saldo * fator;
    var saldoFinal2 = (metodoJuros === 'indexador_mais_taxa')
      ? saldoIdx * Math.pow(1 + (taxaMensal/100), meses)
      : saldoIdx;
    return { ok: true, saldoFinal: saldoFinal2, juros: saldoFinal2 - saldo };
  }
  // fixo_simples (padrão)
  var juros = saldo * (taxaMensal/100) * meses;
  return { ok: true, saldoFinal: saldo + juros, juros: juros };
}

// Percorre TODOS os movimentos em ordem cronológica, capitalizando juros
// entre cada evento, e retorna o extrato linha a linha + saldo atual.
function recalcularLedgerDivida(body) {
  var movimentos = (body.movimentos || []).slice().sort(function(a,b){
    return a.data < b.data ? -1 : (a.data > b.data ? 1 : 0);
  });
  var metodoJuros = body.metodoJuros;
  var taxaMensal  = parseFloat(body.taxaMensal) || 0;
  var indexador   = body.indexador || '';
  var hoje        = body.hoje || Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');

  var saldo = 0;
  var dataBase = null;
  var linhas = [];

  for (var i = 0; i < movimentos.length; i++) {
    var m = movimentos[i];
    var seg = { juros: 0 };
    if (dataBase) {
      seg = calcularJurosSegmento(saldo, metodoJuros, taxaMensal, indexador, dataBase, m.data);
      if (!seg.ok) return { ok: false, error: seg.error };
      saldo = seg.saldoFinal;
    }
    if (m.tipo === 'aporte') {
      saldo += parseFloat(m.valor) || 0;
    } else {
      saldo -= parseFloat(m.valor) || 0;
    }
    dataBase = m.data;
    linhas.push({
      id: m.id, data: m.data, tipo: m.tipo, valor: parseFloat(m.valor) || 0, obs: m.obs || '',
      jurosNoPeriodo: Math.round((seg.juros || 0) * 100) / 100,
      saldoApos: Math.round(saldo * 100) / 100
    });
  }

  // Juros desde o último movimento até hoje
  var jurosFinal = 0;
  var saldoAposUltimo = saldo;
  if (dataBase && dataBase < hoje) {
    var segF = calcularJurosSegmento(saldo, metodoJuros, taxaMensal, indexador, dataBase, hoje);
    if (!segF.ok) return { ok: false, error: segF.error };
    jurosFinal = segF.juros;
    saldo = segF.saldoFinal;
  }

  return {
    ok: true,
    linhas: linhas,
    saldoAtual: Math.round(saldo * 100) / 100,
    saldoAposUltimoMovimento: Math.round(saldoAposUltimo * 100) / 100,
    jurosDesdeUltimo: Math.round(jurosFinal * 100) / 100,
    ultimaData: dataBase,
    primeiraData: movimentos.length ? movimentos[0].data : null,
    primeiroValor: movimentos.length ? movimentos[0].valor : 0
  };
}

function salvarDividaEstruturada(body) {
  var sheets = garantirEstruturaDividas();
  var sh = sheets.dividas;
  var dados = sh.getDataRange().getValues();
  var rowIdx = -1;
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(body.id)) { rowIdx = i + 1; break; }
  }

  var now = new Date();
  var hoje = Utilities.formatDate(now, 'America/Sao_Paulo', 'yyyy-MM-dd');
  var movimentos = body.movimentos || [];

  var ledger = recalcularLedgerDivida({
    movimentos: movimentos, metodoJuros: body.metodoJuros,
    taxaMensal: body.taxaMensal, indexador: body.indexador, hoje: hoje
  });

  var dataOriginalDate = ledger.ok && ledger.primeiraData ? new Date(ledger.primeiraData + 'T12:00:00') : '';
  var dataBaseDate     = ledger.ok && ledger.ultimaData   ? new Date(ledger.ultimaData   + 'T12:00:00') : dataOriginalDate;
  var valorOriginal    = ledger.ok ? (ledger.primeiroValor || 0) : (body.valorOriginal || 0);
  var saldoAtual        = ledger.ok ? ledger.saldoAtual : (body.valorPrincipalAtual || 0);
  var saldoAposUltimo   = ledger.ok ? ledger.saldoAposUltimoMovimento : saldoAtual;
  var diasEmAtraso      = (ledger.ok && ledger.ultimaData) ? Math.round((new Date(hoje+'T12:00:00') - new Date(ledger.ultimaData+'T12:00:00'))/(1000*60*60*24)) : 0;

  // Determinar status automaticamente
  var status = body.status || 'pendente';
  if (ledger.ok) {
    if (saldoAtual <= 0.01) status = 'pago';
    else if (movimentos.length > 1) status = 'parcial';
    else status = 'pendente';
  }

  var rowData = [
    String(body.id),
    body.tipo || 'devo',
    String(body.pagadorId || ''),
    body.pagadorNome || '',
    body.desc || '',
    valorOriginal,
    dataOriginalDate,
    body.metodoJuros || '',
    parseFloat(body.taxaMensal) || 0,
    body.indexador || '',
    saldoAposUltimo,
    dataBaseDate,
    diasEmAtraso,
    saldoAtual,
    status,
    body.criadoEm ? new Date(body.criadoEm) : now,
    now
  ];

  if (rowIdx > -1) {
    sh.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sh.appendRow(rowData);
  }

  // Sincronizar movimentos (substitui tudo pelo array atual enviado)
  if (body.movimentos) {
    var linhasPorId = {};
    if (ledger.ok) {
      ledger.linhas.forEach(function(l){ linhasPorId[String(l.id)] = l; });
    }
    sincronizarSubArrayDivida(sheets.movimentos, body.id, body.movimentos, function(m) {
      var l = linhasPorId[String(m.id)] || {};
      return [String(m.id), String(body.id), body.pagadorNome || '', m.tipo || '',
              m.data ? new Date(m.data + 'T12:00:00') : '', parseFloat(m.valor) || 0, m.obs || '',
              l.jurosNoPeriodo || 0, l.saldoApos || 0];
    });
  }

  return { ok: true, status: status, saldoAtual: saldoAtual };
}

function sincronizarSubArrayDivida(sheet, dividaId, novosItens, mapFn) {
  var dados = sheet.getDataRange().getValues();
  for (var i = dados.length - 1; i >= 1; i--) {
    if (String(dados[i][1]) === String(dividaId)) sheet.deleteRow(i + 1);
  }
  novosItens.forEach(function(item) { sheet.appendRow(mapFn(item)); });
}

function getDividasEstruturadas() {
  var sheets = garantirEstruturaDividas();
  var dDados = sheets.dividas.getDataRange().getValues();
  var mDados = sheets.movimentos.getDataRange().getValues();

  var movPorDivida = {};
  for (var i = 1; i < mDados.length; i++) {
    var did = String(mDados[i][1]); if (!did) continue;
    (movPorDivida[did] = movPorDivida[did] || []).push({
      id: mDados[i][0], data: fmtDate(mDados[i][4]), tipo: mDados[i][3],
      valor: parseFloat(mDados[i][5]) || 0, obs: mDados[i][6] || ''
    });
  }

  var result = [];
  for (var k = 1; k < dDados.length; k++) {
    var row = dDados[k]; if (!row[0]) continue;
    var id = String(row[0]);
    var movs = (movPorDivida[id] || []).sort(function(a,b){ return a.data < b.data ? -1 : 1; });
    result.push({
      id: id,
      tipo: row[1] || 'devo',
      pagadorId: String(row[2] || ''),
      pagadorNome: row[3] || '',
      desc: row[4] || '',
      valorOriginal: parseFloat(row[5]) || 0,
      dataOriginal: fmtDate(row[6]),
      metodoJuros: row[7] || '',
      taxaMensal: parseFloat(row[8]) || 0,
      indexador: row[9] || '',
      valorPrincipalAtual: parseFloat(row[10]) || 0,
      dataBaseAtual: fmtDate(row[11]),
      diasEmAtraso: row[12],
      saldoDevedorEstimado: (typeof row[13] === 'number') ? row[13] : null,
      status: row[14] || 'pendente',
      criadoEm: row[15] instanceof Date ? row[15].toISOString() : String(row[15] || ''),
      movimentos: movs
    });
  }
  return { ok: true, data: result };
}

function deletarDivida(id) {
  var sheets = garantirEstruturaDividas();
  var dados = sheets.dividas.getDataRange().getValues();
  for (var i = dados.length - 1; i >= 1; i--) {
    if (String(dados[i][0]) === String(id)) sheets.dividas.deleteRow(i + 1);
  }
  var d2 = sheets.movimentos.getDataRange().getValues();
  for (var j = d2.length - 1; j >= 1; j--) {
    if (String(d2[j][1]) === String(id)) sheets.movimentos.deleteRow(j + 1);
  }
  return { ok: true };
}

function getDividasSheetUrl() {
  var sp = dividasSS();
  garantirEstruturaDividas();
  var sh = sp.getSheetByName('Dívidas');
  var gid = sh ? sh.getSheetId() : 0;
  return { ok: true, url: sp.getUrl() + '#gid=' + gid };
}

// ════════════════════════════════════════════════════════════
//  ATUALIZAÇÃO AUTOMÁTICA DOS SALDOS — roda todo dia e grava o
//  saldo atualizado direto na planilha (qualquer método de
//  juros), mesmo que você nunca abra o app.
// ════════════════════════════════════════════════════════════
function atualizarSaldosDividas() {
  var sheets = garantirEstruturaDividas();
  var sh = sheets.dividas;
  var dados = sh.getDataRange().getValues();
  var mDados = sheets.movimentos.getDataRange().getValues();
  var hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');

  var movPorDivida = {};
  for (var i = 1; i < mDados.length; i++) {
    var did = String(mDados[i][1]); if (!did) continue;
    (movPorDivida[did] = movPorDivida[did] || []).push({
      id: mDados[i][0], data: fmtDate(mDados[i][4]), tipo: mDados[i][3], valor: parseFloat(mDados[i][5]) || 0
    });
  }

  var atualizados = 0;
  for (var j = 1; j < dados.length; j++) {
    var row = dados[j]; if (!row[0]) continue;
    var id = String(row[0]);
    var status = row[14];
    if (status === 'pago') continue;
    var movs = movPorDivida[id] || [];
    if (!movs.length) continue;

    var ledger = recalcularLedgerDivida({
      movimentos: movs, metodoJuros: row[7], taxaMensal: row[8], indexador: row[9], hoje: hoje
    });
    if (!ledger.ok) continue;

    var novoStatus = ledger.saldoAtual <= 0.01 ? 'pago' : (movs.length > 1 ? 'parcial' : 'pendente');
    var dias = ledger.ultimaData ? Math.round((new Date(hoje+'T12:00:00') - new Date(ledger.ultimaData+'T12:00:00'))/(1000*60*60*24)) : 0;

    sh.getRange(j+1, 11).setValue(ledger.saldoAposUltimoMovimento); // K
    sh.getRange(j+1, 13).setValue(dias);                            // M
    sh.getRange(j+1, 14).setValue(ledger.saldoAtual);               // N
    sh.getRange(j+1, 15).setValue(novoStatus);                      // O
    sh.getRange(j+1, 17).setValue(new Date());                      // Q
    atualizados++;
  }
  Logger.log('Saldos de dívidas atualizados: ' + atualizados);
  return atualizados;
}

// Execute esta função UMA VEZ para criar o gatilho diário automático
function criarTriggerAtualizarIndices() {
  ['atualizarSaldosDividasIndexadas','atualizarSaldosDividas'].forEach(function(fn){
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === fn) ScriptApp.deleteTrigger(t);
    });
  });
  ScriptApp.newTrigger('atualizarSaldosDividas')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();
  Logger.log('✅ Gatilho criado — saldos de dívidas serão atualizados todo dia às 6h');
}

// ════════════════════════════════════════════════════════════
//  FICHA PÚBLICA DA DÍVIDA / EMPRÉSTIMO — extrato cronológico
//  completo, compartilhável por link
// ════════════════════════════════════════════════════════════
function paginaDivida(e) {
  var divId = e.parameter.divId || '';
  var todos = getDividasEstruturadas();
  var d = (todos.data || []).find(function(x){ return String(x.id) === String(divId); });

  if (!d) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:40px;text-align:center;background:#0F1923;color:#EDF2F7;min-height:100vh">' +
      '<h2 style="color:#FF6B6B">Registro não encontrado</h2>' +
      '<p style="color:#8FA8C4">Verifique o link com quem compartilhou.</p></div>'
    );
  }

  var hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd');
  var ledger = recalcularLedgerDivida({
    movimentos: d.movimentos, metodoJuros: d.metodoJuros,
    taxaMensal: d.taxaMensal, indexador: d.indexador, hoje: hoje
  });
  var saldoAtual = ledger.ok ? ledger.saldoAtual : d.valorPrincipalAtual;

  var isDevo = d.tipo !== 'recebo';
  var corTema    = isDevo ? '#E0935C' : '#5CA8E0';
  var corTemaBg  = isDevo ? '#2A1F14' : '#142436';
  var tituloTipo = isDevo ? 'Empréstimo tomado' : 'Empréstimo concedido';

  function brl(v){ return 'R$ ' + Number(v||0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
  function brDate(iso){ return iso ? String(iso).split('-').reverse().join('/') : '—'; }
  var tipoIcones = { aporte:'📥', pagamento:'💵', despesa:'🧾' };
  var tipoLabels = { aporte:'Aporte/Empréstimo', pagamento:'Pagamento', despesa:'Despesa em nome' };

  var totalAportes    = (d.movimentos||[]).filter(function(m){return m.tipo==='aporte';}).reduce(function(s,m){return s+m.valor;},0);
  var totalPagamentos = (d.movimentos||[]).filter(function(m){return m.tipo!=='aporte';}).reduce(function(s,m){return s+m.valor;},0);
  var pctPago = totalAportes>0 ? Math.min(100, Math.round((totalPagamentos/totalAportes)*100)) : 0;

  var statusLabel = { pendente:'EM ABERTO', parcial:'PARCIALMENTE PAGO', pago:'PAGO INTEGRALMENTE' }[d.status] || String(d.status||'').toUpperCase();
  var statusCor   = { pendente:'#E74C3C', parcial:'#E0A30C', pago:'#2ECC9A' }[d.status] || '#888';

  var linhasMovimentos = '';
  if (ledger.ok && ledger.linhas.length) {
    ledger.linhas.forEach(function(l){
      var corValor = l.tipo === 'aporte' ? '#E0935C' : '#2ECC9A';
      var sinal = l.tipo === 'aporte' ? '+' : '−';
      linhasMovimentos +=
        '<tr>' +
          '<td style="padding:10px 8px;border-bottom:1px solid rgba(255,255,255,.06);font-size:12px">' + brDate(l.data) + '</td>' +
          '<td style="padding:10px 8px;border-bottom:1px solid rgba(255,255,255,.06);font-size:12px">' + (tipoIcones[l.tipo]||'') + ' ' + (tipoLabels[l.tipo]||l.tipo) + (l.obs?'<br><span style="color:#526680;font-size:10px">'+l.obs+'</span>':'') + '</td>' +
          '<td style="padding:10px 8px;border-bottom:1px solid rgba(255,255,255,.06);font-size:13px;text-align:right;color:'+corValor+';font-weight:700">' + sinal + ' ' + brl(l.valor) + '</td>' +
          '<td style="padding:10px 8px;border-bottom:1px solid rgba(255,255,255,.06);font-size:11px;text-align:right;color:#8FA8C4">' + (l.jurosNoPeriodo>0?brl(l.jurosNoPeriodo):'—') + '</td>' +
          '<td style="padding:10px 8px;border-bottom:1px solid rgba(255,255,255,.06);font-size:13px;text-align:right;font-weight:700">' + brl(l.saldoApos) + '</td>' +
        '</tr>';
    });
  } else {
    linhasMovimentos = '<tr><td colspan="5" style="padding:20px;text-align:center;color:#526680;font-size:13px">Nenhum movimento registrado ainda</td></tr>';
  }

  var html = '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Extrato — ' + d.desc + '</title>' +
    '<style>*{margin:0;padding:0;box-sizing:border-box}' +
    'body{font-family:Arial,Helvetica,sans-serif;background:#0F1923;display:flex;justify-content:center;padding:20px;min-height:100vh}' +
    '.card{background:#1A2A3A;border-radius:16px;overflow:hidden;width:100%;max-width:620px;border:1px solid rgba(255,255,255,.08)}' +
    '.hdr{background:#0d1117;padding:22px 20px;border-bottom:2px solid ' + corTema + '}' +
    '.hdr h1{color:#fff;font-size:18px;font-weight:900}' +
    '.hdr .sub{color:' + corTema + ';font-size:12px;font-weight:700;margin-top:4px;text-transform:uppercase;letter-spacing:.5px}' +
    '.body{padding:20px}' +
    '.row{padding:11px 0;border-bottom:1px solid rgba(255,255,255,.06);display:flex;justify-content:space-between;align-items:baseline}' +
    '.lbl{font-size:11px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;color:#526680}' +
    '.val{font-size:14px;font-weight:600;color:#EDF2F7}' +
    '.saldo-box{background:' + corTemaBg + ';border-radius:12px;padding:20px;text-align:center;margin:16px 0}' +
    '.saldo-box .lbl{color:' + corTema + '}' +
    '.saldo-box .amt{font-size:32px;font-weight:900;color:#fff;margin-top:6px}' +
    '.progress-track{height:8px;background:rgba(255,255,255,.08);border-radius:4px;margin-top:14px;overflow:hidden}' +
    '.progress-fill{height:8px;background:#2ECC9A;border-radius:4px}' +
    '.progress-lbl{font-size:11px;color:#8FA8C4;margin-top:6px;text-align:center}' +
    'table{width:100%;border-collapse:collapse;margin-top:8px}' +
    'th{font-size:10px;font-weight:800;letter-spacing:.5px;text-transform:uppercase;color:#526680;text-align:left;padding:8px;border-bottom:2px solid rgba(255,255,255,.1)}' +
    '.footer{padding:14px 20px;text-align:center;color:#526680;font-size:11px;border-top:1px solid rgba(255,255,255,.06)}' +
    '@media print{body{background:#fff}.card{border:1px solid #ddd;background:#fff}}' +
    '</style></head><body>' +
    '<div class="card">' +
      '<div class="hdr"><h1>⚡ ' + tituloTipo + '</h1><div class="sub">' + statusLabel + '</div></div>' +
      '<div class="body">' +
        '<div class="row"><span class="lbl">Pessoa</span><span class="val">' + d.pagadorNome + '</span></div>' +
        '<div class="row"><span class="lbl">Descrição</span><span class="val">' + d.desc + '</span></div>' +
        '<div class="row"><span class="lbl">Total de aportes</span><span class="val">' + brl(totalAportes) + '</span></div>' +
        '<div class="row"><span class="lbl">Total já pago/abatido</span><span class="val" style="color:#2ECC9A">' + brl(totalPagamentos) + '</span></div>' +

        '<div class="saldo-box">' +
          '<div class="lbl">Saldo devedor atual</div>' +
          '<div class="amt">' + (ledger.ok ? brl(saldoAtual) : 'Indisponível') + '</div>' +
          '<div class="progress-track"><div class="progress-fill" style="width:' + pctPago + '%"></div></div>' +
          '<div class="progress-lbl">' + pctPago + '% já abatido do total aportado</div>' +
        '</div>' +

        '<table>' +
          '<thead><tr><th>Data</th><th>Movimento</th><th style="text-align:right">Valor</th><th style="text-align:right">Juros no período</th><th style="text-align:right">Saldo</th></tr></thead>' +
          '<tbody>' + linhasMovimentos + '</tbody>' +
        '</table>' +
      '</div>' +
      '<div class="footer">Extrato gerado em ' + brDate(hoje) + ' · via Fluxo App</div>' +
    '</div>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Extrato — ' + d.desc)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ════════════════════════════════════════════════════════════
//  DÍVIDAS — cálculo isolado de UM segmento (usado pelo app
//  quando precisa só de um número rápido, sem o extrato completo)
// ════════════════════════════════════════════════════════════
var BCB_CODIGOS = {
  'IGPM':  189,
  'INPC':  188,
  'CDI':   4391,
  'SELIC': 4390
};

function calcularSaldoDivida(body) {
  var seg = calcularJurosSegmento(
    parseFloat(body.valorPrincipal) || 0, body.metodoJuros, parseFloat(body.taxaMensal) || 0,
    body.indexador || '', String(body.dataBase), String(body.hoje || Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd'))
  );
  if (!seg.ok) return { ok: false, error: seg.error };
  return {
    ok: true,
    saldoDevedor: Math.round(seg.saldoFinal * 100) / 100,
    jurosAcumulado: Math.round(seg.juros * 100) / 100
  };
}

function buscarIndexadorAcumulado(tipo, dataInicioISO, dataFimISO) {
  var codigo = BCB_CODIGOS[tipo];
  if (!codigo) return null;

  var cacheKey = 'idx_' + tipo + '_' + dataInicioISO + '_' + dataFimISO;
  var cache = CacheService.getScriptCache();
  var cached = cache.get(cacheKey);
  if (cached !== null) return parseFloat(cached);

  try {
    var dIni = isoToBr(dataInicioISO);
    var dFim = isoToBr(dataFimISO);
    var url = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.' + codigo +
              '/dados?formato=json&dataInicial=' + dIni + '&dataFinal=' + dFim;

    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      Logger.log('BCB API erro: ' + resp.getResponseCode());
      return null;
    }

    var dados = JSON.parse(resp.getContentText());
    if (!dados || !dados.length) {
      cache.put(cacheKey, '1', 21600);
      return 1;
    }

    var fator = 1;
    dados.forEach(function(item) {
      var valor = parseFloat(String(item.valor).replace(',', '.'));
      if (!isNaN(valor)) fator *= (1 + valor/100);
    });

    cache.put(cacheKey, String(fator), 21600);
    return fator;
  } catch (e) {
    Logger.log('Erro ao buscar indexador ' + tipo + ': ' + e.message);
    return null;
  }
}

function isoToBr(isoDate) {
  var p = isoDate.split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}

// ════════════════════════════════════════════════════════════
//  BANCOS — Contas e Extratos OFX
//  ContasBanco: usa salvarItemSheet/getItemsSheet (padrão)
//  Extratos:    aba própria com uma linha por extrato (JSON)
//               — pode ter centenas de transações por extrato
// ════════════════════════════════════════════════════════════

function garantirAbaExtratos() {
  var sheet = ss().getSheetByName('Extratos');
  if (!sheet) {
    sheet = ss().insertSheet('Extratos');
    sheet.getRange(1,1,1,4).setValues([['contaId','extratoId','arquivo','dados']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,4).setFontWeight('bold').setBackground('#0F9D58').setFontColor('#FFF');
    sheet.setColumnWidth(4, 400);
  }
  return sheet;
}

function salvarExtratoSheet(body) {
  var sheet = garantirAbaExtratos();
  var contaId   = String(body.contaId   || '');
  var extratoId = String(body.extratoId || '');
  var arquivo   = String(body.arquivo   || '');
  // Serializar as transações como JSON (sem as txs para economizar espaço,
  // só metadados — as txs ficam num campo separado comprimido)
  var dados = JSON.stringify({
    id:          extratoId,
    arquivo:     arquivo,
    banco:       body.banco    || '',
    conta:       body.conta    || '',
    dataIni:     body.dataIni  || '',
    dataFim:     body.dataFim  || '',
    saldoFim:    body.saldoFim || 0,
    importadoEm: body.importadoEm || new Date().toISOString(),
    txs:         body.txs || []
  });

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === contaId && String(rows[i][1]) === extratoId) {
      sheet.getRange(i+1, 4).setValue(dados);
      return { ok: true };
    }
  }
  sheet.appendRow([contaId, extratoId, arquivo, dados]);
  return { ok: true };
}

function getExtratosSheet(contaId) {
  var sheet = garantirAbaExtratos();
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(contaId)) continue;
    try {
      var obj = JSON.parse(String(rows[i][3]));
      result.push(obj);
    } catch(e) {}
  }
  return { ok: true, data: result };
}

function deletarExtratoSheet(contaId, extratoId) {
  var sheet = garantirAbaExtratos();
  var rows = sheet.getDataRange().getValues();
  for (var i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(contaId) && String(rows[i][1]) === String(extratoId)) {
      sheet.deleteRow(i+1);
      return { ok: true };
    }
  }
  return { ok: true };
}


// ════════════════════════════════════════════════════════════
//  PÁGINA PÚBLICA DO INQUILINO
//  Retorna HTML com histórico de pagamentos do contrato
// ════════════════════════════════════════════════════════════

function paginaInquilino(e) {
  var ctId  = e.parameter.ctId  || '';
  var recId = e.parameter.recId || '';

  // ── Exibir recibo avulso (sem contrato) ──────────────────
  if (recId && !ctId) {
    var recSheet = ss().getSheetByName('Recibos');
    var recibo = null;
    if (recSheet && recSheet.getLastRow() > 1) {
      var recDados = recSheet.getDataRange().getValues();
      for (var ri = 1; ri < recDados.length; ri++) {
        if (!recDados[ri][1]) continue;
        try {
          var ro = JSON.parse(String(recDados[ri][1]));
          if (String(ro.id) === String(recId)) { recibo = ro; break; }
        } catch(re2) {}
      }
    }
    if (!recibo) {
      return HtmlService.createHtmlOutput(
        '<div style="font-family:Arial;padding:40px;text-align:center;background:#0F1923;color:#EDF2F7;min-height:100vh">' +
        '<h2 style="color:#FF6B6B">Recibo não encontrado</h2>' +
        '<p style="color:#8FA8C4">Verifique o link com quem emitiu o recibo.</p></div>'
      );
    }
    var dataFmt = recibo.data ? String(recibo.data).split('-').reverse().join('/') : '';
    var valorFmt = 'R$ ' + Number(recibo.valor||0).toFixed(2).replace('.',',').replace(/\B(?=(\d{3})+(?!\d))/g,'.');
    var html = '<!DOCTYPE html><html lang="pt-BR"><head><meta charset="UTF-8">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<title>Recibo '+recibo.codigo+'</title>' +
      '<style>*{margin:0;padding:0;box-sizing:border-box}' +
      'body{font-family:Arial,sans-serif;background:#0F1923;display:flex;justify-content:center;padding:20px;min-height:100vh}' +
      '.card{background:#1A2A3A;border-radius:16px;overflow:hidden;width:100%;max-width:500px;border:1px solid rgba(255,255,255,.08)}' +
      '.hdr{background:#0d1117;padding:20px;border-bottom:2px solid #2ECC9A;display:flex;justify-content:space-between;align-items:center}' +
      '.hdr h1{color:#fff;font-size:18px;font-weight:900}' +
      '.hdr span{color:#2ECC9A;font-family:monospace;font-size:13px}' +
      '.body{padding:20px}' +
      '.row{padding:12px 0;border-bottom:1px solid rgba(255,255,255,.06);display:flex;flex-direction:column;gap:3px}' +
      '.lbl{font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:#526680}' +
      '.val{font-size:15px;font-weight:600;color:#EDF2F7}' +
      '.total{background:#0a1421;border-radius:10px;padding:20px;text-align:center;margin-top:16px}' +
      '.total .lbl{margin-bottom:6px}' +
      '.total .amt{font-size:32px;font-weight:900;color:#2ECC9A}' +
      '.footer{padding:14px 20px;text-align:center;color:#526680;font-size:11px;border-top:1px solid rgba(255,255,255,.06)}' +
      '@media print{body{background:#fff}.card{border:1px solid #ddd;background:#fff}.hdr{background:#0d1117}.row{border-bottom:1px solid #eee}.total{background:#f5f5f5}}' +
      '</style></head><body>' +
      '<div class="card">' +
      '<div class="hdr"><h1>⚡ Fluxo App</h1><span>'+recibo.codigo+'</span></div>' +
      '<div class="body">' +
      '<div class="row"><div class="lbl">Recebemos de</div><div class="val">'+recibo.pagador+(recibo.cpf?' — CPF: '+recibo.cpf:'')+'</div></div>' +
      '<div class="row"><div class="lbl">Referente a</div><div class="val">'+recibo.desc+'</div></div>' +
      '<div class="row"><div class="lbl">Data</div><div class="val">'+dataFmt+'</div></div>' +
      (recibo.obs?'<div class="row"><div class="lbl">Observação</div><div class="val">'+recibo.obs+'</div></div>':'') +
      '<div class="total"><div class="lbl">Valor pago</div><div class="amt">'+valorFmt+'</div></div>' +
      '</div>' +
      '<div class="footer">Emitido em '+recibo.emitidoEm+' · via Fluxo App</div>' +
      '</div>' +
      '<script>// Botão imprimir removido pois não necessário no mobile<\/script>' +
      '</body></html>';
    return HtmlService.createHtmlOutput(html)
      .setTitle('Recibo '+recibo.codigo)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── Exibir página do inquilino (contrato) ────────────────
  var ctSheet = ss().getSheetByName('Contratos');
  var ct = null;
  if (ctSheet) {
    var dados = ctSheet.getDataRange().getValues();
    for (var i = 1; i < dados.length; i++) {
      if (!dados[i][1]) continue;
      try {
        var obj = JSON.parse(String(dados[i][1]));
        if (String(obj.id) === ctId) { ct = obj; break; }
      } catch(e2) {}
    }
  }

  if (!ct) {
    return HtmlService.createHtmlOutput('<div style="font-family:Arial;padding:40px;text-align:center"><h2>Contrato não encontrado</h2><p>Verifique o link com seu locador.</p></div>');
  }

  // Buscar pagamentos da aba de controle
  var nomeAba = 'Contrato - ' + String(ct.inqNome || '').substring(0, 18);
  var pgSheet = ss().getSheetByName(nomeAba);
  var linhas  = [];
  if (pgSheet) {
    var pg = pgSheet.getDataRange().getValues();
    for (var j = 6; j < pg.length; j++) {
      if (!pg[j][0]) continue;
      var mesVal = pg[j][0];
      // Se Sheets converteu para Date, formatar de volta para "Mmm/AAAA"
      if (mesVal instanceof Date) {
        var meses_fix = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
        mesVal = meses_fix[mesVal.getMonth()] + '/' + mesVal.getFullYear();
      } else {
        mesVal = String(mesVal || '').trim();
      }
      linhas.push({ mes: mesVal, valor: pg[j][1], data: pg[j][2], status: pg[j][3], codigo: pg[j][4] });
    }
  }

  function fmtR(v){ return 'R$ ' + parseFloat(v||0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
  function esc(s){ return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }
  function fmtD(s){ if(!s) return '—'; var p=String(s).split('-'); return p.length===3?p[2]+'/'+p[1]+'/'+p[0]:s; }

  var totalPago = 0, totalPend = 0;
  linhas.forEach(function(l){
    if(String(l.status).toUpperCase()==='PAGO') totalPago += parseFloat(l.valor||0);
    else totalPend += parseFloat(l.valor||0);
  });

  // ── SEÇÃO 1: Resumo ──────────────────────────────────
  var htmlResumo =
    '<div class="section active" id="s-resumo">'+
    '<div class="info-card">'+
      '<div class="label">Locatário</div>'+
      '<div class="val">'+esc(ct.inqNome)+'</div>'+
      (ct.inqCpf?'<div class="sub">CPF: '+esc(ct.inqCpf)+'</div>':'')+
      (ct.inqNome2?'<div class="sub">+ '+esc(ct.inqNome2)+(ct.inqCpf2?' — CPF: '+esc(ct.inqCpf2):'')+'</div>':'')+
    '</div>'+
    '<div class="info-card">'+
      '<div class="label">Imóvel</div>'+
      '<div class="val">'+esc(ct.tipo||'Imóvel')+' — '+esc(ct.imovel)+'</div>'+
      '<div class="sub">'+esc(ct.end)+'</div>'+
    '</div>'+
    '<div class="info-card">'+
      '<div class="label">Contrato</div>'+
      '<div class="val">'+fmtR(ct.valor)+' /mês</div>'+
      '<div class="sub">Vencimento dia '+esc(ct.diaPgto)+' · Início: '+fmtD(ct.inicio)+(ct.fim?' · Término: '+fmtD(ct.fim):'')+'</div>'+
    '</div>'+
    '<div class="info-card">'+
      '<div class="label">Locador</div>'+
      '<div class="val">'+esc(ct.locNome)+'</div>'+
      (ct.locCpf?'<div class="sub">CPF: '+esc(ct.locCpf)+'</div>':'')+
    '</div>'+
    '<div class="totais">'+
      '<div class="tot pago"><div class="tv">'+fmtR(totalPago)+'</div><div class="tl">Total Pago</div></div>'+
      '<div class="tot pend"><div class="tv">'+fmtR(totalPend)+'</div><div class="tl">Pendente</div></div>'+
    '</div>'+
    '</div>';

  // ── SEÇÃO 2: Pagamentos ───────────────────────────────
  var rows = linhas.map(function(l){
    var pago = String(l.status).toUpperCase()==='PAGO';
    return '<tr class="'+(pago?'row-pago':'row-pend')+'">'+
      '<td>'+esc(l.mes)+'</td>'+
      '<td style="font-weight:700">'+fmtR(l.valor)+'</td>'+
      '<td>'+esc(fmtD(l.data))+'</td>'+
      '<td><span class="badge '+(pago?'b-pago':'b-pend')+'">'+esc(l.status||'PENDENTE')+'</span></td>'+
      '<td class="mono">'+esc(l.codigo||'—')+'</td>'+
    '</tr>';
  }).join('');

  var htmlPgto =
    '<div class="section" id="s-pgto">'+
    '<table><thead><tr><th>Mês</th><th>Valor</th><th>Pago em</th><th>Status</th><th>Recibo</th></tr></thead>'+
    '<tbody>'+(rows||'<tr><td colspan="5" style="text-align:center;padding:20px;color:#999">Nenhum lançamento</td></tr>')+'</tbody></table>'+
    '</div>';

  // ── SEÇÃO 3: Contrato completo ────────────────────────
  var durLabel = ct.duracao ? ct.duracao + ' meses' : 'Indeterminado';
  var htmlContrato =
    '<div class="section" id="s-contrato">'+
    '<div style="background:#0F2438;color:#fff;border-radius:10px;padding:20px;text-align:center;margin-bottom:16px">'+
      '<div style="font-size:22px;font-weight:900;letter-spacing:2px">⚡ FLUXO</div>'+
      '<div style="font-size:13px;color:#aaa;margin-top:4px;letter-spacing:1px">CONTRATO DE LOCAÇÃO</div>'+
    '</div>'+
    ct_secao('1. PARTES', [
      ['Locador', esc(ct.locNome)+(ct.locCpf?' — CPF: '+esc(ct.locCpf):'')+' — RG: '+esc(ct.locRg||'—')],
      (ct.locNome2?['2º Locador', esc(ct.locNome2)+(ct.locCpf2?' — CPF: '+esc(ct.locCpf2):'')]:null),
      ['Locatário', esc(ct.inqNome)+(ct.inqCpf?' — CPF: '+esc(ct.inqCpf):'')+' — RG: '+esc(ct.inqRg||'—')],
      (ct.inqNome2?['2º Locatário', esc(ct.inqNome2)+(ct.inqCpf2?' — CPF: '+esc(ct.inqCpf2):'')]:null),
      ['Endereço (locatário)', esc(ct.inqEnd||'—')]
    ])+
    ct_secao('2. IMÓVEL', [
      ['Tipo', esc(ct.tipo||'—')],
      ['Identificação', esc(ct.imovel)],
      ['Endereço', esc(ct.end||'—')],
      ['Matrícula', esc(ct.mat||'—')]
    ])+
    ct_secao('3. VIGÊNCIA E VALOR', [
      ['Início', fmtD(ct.inicio)],
      ['Término', ct.fim?fmtD(ct.fim):'Indeterminado'],
      ['Duração', durLabel],
      ['Valor do aluguel', fmtR(ct.valor)+' mensais'],
      ['Dia de vencimento', 'Todo dia '+esc(ct.diaPgto)+' de cada mês']
    ])+
    ct_secao('4. VISTORIA — ESTADO DO IMÓVEL', [
      ['Descrição geral', esc(ct.vistoria||'—')],
      ['Itens e condições', esc(ct.itens||'—')]
    ])+
    '<div class="clausulas">'+
      '<div class="cl-title">5. CLÁUSULAS GERAIS</div>'+
      '<ol>'+
        '<li>O LOCATÁRIO obriga-se a pagar o aluguel até o dia '+esc(ct.diaPgto)+' de cada mês.</li>'+
        '<li>O atraso no pagamento sujeitará o LOCATÁRIO a multa de 10% sobre o valor do aluguel, acrescida de juros de 1% ao mês.</li>'+
        '<li>O LOCATÁRIO declara ter recebido o imóvel nas condições descritas na vistoria e compromete-se a devolvê-lo nas mesmas condições.</li>'+
        '<li>São proibidas obras, reformas ou modificações no imóvel sem autorização prévia e por escrito do LOCADOR.</li>'+
        '<li>O contrato poderá ser rescindido por qualquer das partes mediante aviso prévio de 30 dias.</li>'+
        '<li>O LOCATÁRIO não poderá sublocar, ceder ou emprestar o imóvel, no todo ou em parte, sem autorização escrita do LOCADOR.</li>'+
        '<li>As despesas de condomínio, IPTU e demais encargos do imóvel são de responsabilidade definida entre as partes conforme acordado verbalmente ou em adendo a este contrato.</li>'+
        '<li>Fica eleito o foro da comarca de Chapecó/SC para dirimir quaisquer dúvidas oriundas deste contrato.</li>'+
      '</ol>'+
    '</div>'+
    '<div class="assinaturas">'+
      '<div class="ass"><div class="ass-linha"></div><div>'+esc(ct.locNome)+'</div><div style="font-size:11px;color:#999">Locador</div></div>'+
      '<div class="ass"><div class="ass-linha"></div><div>'+esc(ct.inqNome)+'</div><div style="font-size:11px;color:#999">Locatário</div></div>'+
      (ct.inqNome2?'<div class="ass"><div class="ass-linha"></div><div>'+esc(ct.inqNome2)+'</div><div style="font-size:11px;color:#999">Locatário</div></div>':'')+
    '</div>'+
    '<div style="text-align:center;margin-top:16px">'+
      '<button onclick="window.print()" style="padding:12px 28px;background:#0F2438;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer">🖨️ Imprimir / Salvar PDF</button>'+
    '</div>'+
    '</div>';

  function ct_secao(titulo, campos) {
    var linhasHtml = campos.filter(function(f){return f;}).map(function(f){
      return '<tr><td class="cl-key">'+f[0]+'</td><td class="cl-val">'+f[1]+'</td></tr>';
    }).join('');
    return '<div class="ct-secao"><div class="ct-titulo">'+titulo+'</div>'+
      '<table class="ct-table">'+linhasHtml+'</table></div>';
  }

  // ── HTML FINAL ────────────────────────────────────────
  var html = '<!DOCTYPE html><html lang="pt-BR"><head>'+
    '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'+
    '<title>Área do Inquilino — '+esc(ct.inqNome)+'</title>'+
    '<style>'+
    '*{margin:0;padding:0;box-sizing:border-box}'+
    'body{font-family:-apple-system,Arial,sans-serif;background:#f0f2f5;color:#111}'+
    '.header{background:#0F2438;color:#fff;padding:20px 16px;position:sticky;top:0;z-index:10}'+
    '.header h1{font-size:16px;font-weight:900}'+
    '.header p{font-size:12px;color:#aaa;margin-top:2px}'+
    '.tabs{display:flex;background:#1a2a3a;overflow-x:auto}'+
    '.tab{flex:1;padding:12px 8px;border:none;background:transparent;color:#aaa;font-size:12px;font-weight:700;cursor:pointer;text-align:center;white-space:nowrap;letter-spacing:.3px}'+
    '.tab.on{color:#2ecc9a;border-bottom:2px solid #2ecc9a}'+
    '.content{padding:16px;max-width:640px;margin:0 auto}'+
    '.section{display:none}.section.active{display:block}'+
    '.info-card{background:#fff;border-radius:10px;padding:14px 16px;margin-bottom:10px;box-shadow:0 1px 4px rgba(0,0,0,.06)}'+
    '.label{font-size:10px;font-weight:800;color:#999;text-transform:uppercase;letter-spacing:.8px;margin-bottom:4px}'+
    '.val{font-size:15px;font-weight:700}'+
    '.sub{font-size:12px;color:#666;margin-top:2px}'+
    '.totais{display:flex;gap:10px;margin-top:4px}'+
    '.tot{flex:1;text-align:center;padding:14px;border-radius:10px}'+
    '.tot.pago{background:#d4edda}.tot.pend{background:#fff3cd}'+
    '.tv{font-size:18px;font-weight:900}.tl{font-size:11px;color:#666;margin-top:2px}'+
    '.tot.pago .tv{color:#155724}.tot.pend .tv{color:#856404}'+
    'table{width:100%;border-collapse:collapse;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.06)}'+
    'th{padding:10px 12px;text-align:left;background:#0F2438;color:#fff;font-size:11px;font-weight:700;letter-spacing:.5px}'+
    'td{padding:10px 12px;font-size:13px;border-bottom:1px solid #f0f0f0}'+
    '.row-pago td{background:#f8fff9}.row-pend td{background:#fffdf5}'+
    '.badge{padding:3px 8px;border-radius:20px;font-size:10px;font-weight:800}'+
    '.b-pago{background:#d4edda;color:#155724}.b-pend{background:#fff3cd;color:#856404}'+
    '.mono{font-family:monospace;font-size:10px;color:#888}'+
    '.ct-secao{background:#fff;border-radius:10px;padding:16px;margin-bottom:10px;box-shadow:0 1px 4px rgba(0,0,0,.06)}'+
    '.ct-titulo{font-size:11px;font-weight:800;color:#0F2438;text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px;padding-bottom:8px;border-bottom:1px solid #eee}'+
    '.ct-table{width:100%;border-collapse:collapse}'+
    '.cl-key{font-size:11px;color:#999;padding:6px 0;width:40%;vertical-align:top}'+
    '.cl-val{font-size:13px;font-weight:600;padding:6px 0 6px 8px}'+
    '.clausulas{background:#fff;border-radius:10px;padding:16px;margin-bottom:10px;box-shadow:0 1px 4px rgba(0,0,0,.06)}'+
    '.cl-title{font-size:11px;font-weight:800;color:#0F2438;text-transform:uppercase;letter-spacing:.8px;margin-bottom:12px}'+
    '.clausulas ol{padding-left:18px}.clausulas li{font-size:13px;margin-bottom:8px;line-height:1.5;color:#333}'+
    '.assinaturas{display:flex;gap:20px;margin:20px 0;flex-wrap:wrap}'+
    '.ass{flex:1;min-width:140px;text-align:center;font-size:13px;font-weight:600}'+
    '.ass-linha{border-top:1px solid #333;margin-bottom:8px;margin-top:40px}'+
    '@media print{.header,.tabs{display:none}.section{display:block!important}#s-resumo,#s-pgto{display:none!important}#s-contrato{display:block!important}.content{padding:0;max-width:100%}}'+
    '</style></head><body>'+
    '<div class="header">'+
      '<h1>⚡ Área do Inquilino</h1>'+
      '<p>'+esc(ct.inqNome)+' · '+esc(ct.imovel)+'</p>'+
    '</div>'+
    '<div class="tabs">'+
      '<button class="tab on" onclick="showTab(\'resumo\')">📊 Resumo</button>'+
      '<button class="tab" onclick="showTab(\'pgto\')">💰 Pagamentos</button>'+
      '<button class="tab" onclick="showTab(\'contrato\')">📄 Contrato</button>'+
    '</div>'+
    '<div class="content">'+
      htmlResumo + htmlPgto + htmlContrato +
    '</div>'+
    '<script>'+
    'function showTab(t){'+
      'document.querySelectorAll(".section").forEach(function(s){s.classList.remove("active")});'+
      'document.querySelectorAll(".tab").forEach(function(b){b.classList.remove("on")});'+
      'document.getElementById("s-"+t).classList.add("active");'+
      'event.target.classList.add("on");'+
    '}'+
    '<\/script>'+
    '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Área do Inquilino — '+ct.inqNome)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ════════════════════════════════════════════════════════════
//  DEBUG — execute no editor para testar salvarItemSheet
// ════════════════════════════════════════════════════════════
function testarSalvarContrato() {
  var teste = {
    id: 'TESTE-'+Date.now(),
    inqNome: 'Inquilino Teste',
    locNome: 'Locador Teste',
    imovel:  'Sala 01',
    valor:   1500,
    inicio:  '2026-05-01',
    diaPgto: '10',
    tipo:    'comercial'
  };
  var resultado = salvarItemSheet('Contratos', teste);
  Logger.log('Resultado: ' + JSON.stringify(resultado));
  // Verificar se foi criado
  var sheet = ss().getSheetByName('Contratos');
  Logger.log('Linhas na aba Contratos: ' + (sheet ? sheet.getLastRow() : 'aba não existe'));
}

// ════════════════════════════════════════════════════════════
//  LEMBRETE DE VENCIMENTO DE ALUGUEL
//  Execute criarTriggerLembreteAluguel() UMA VEZ para ativar
//  Envia email 3 dias antes do vencimento de cada aluguel
// ════════════════════════════════════════════════════════════
function enviarLembreteAluguel() {
  var tz   = Session.getScriptTimeZone();
  var hoje = new Date();
  var diaHoje = hoje.getDate();

  // Buscar contratos da aba Contratos
  var sheet = ss().getSheetByName('Contratos');
  if (!sheet) { Logger.log('Aba Contratos não encontrada'); return; }

  var dados = sheet.getDataRange().getValues();
  var enviados = 0;

  for (var i = 1; i < dados.length; i++) {
    if (!dados[i][1]) continue;
    var ct;
    try { ct = JSON.parse(String(dados[i][1])); } catch(e) { continue; }
    if (!ct || !ct.inqNome || !ct.diaPgto || !ct.valor) continue;

    var diaPgto = parseInt(ct.diaPgto);
    var diasAte = diaPgto - diaHoje;

    // Enviar lembrete 3 dias antes e no dia do vencimento
    if (diasAte !== 3 && diasAte !== 0) continue;

    // Verificar se já foi pago este mês
    var mesAno = hoje.getFullYear() + '-' + String(hoje.getMonth() + 1).padStart(2, '0');
    var nomeAba = 'Contrato - ' + String(ct.inqNome).substring(0, 18);
    var pgSheet = ss().getSheetByName(nomeAba);
    var meses_br = ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'];
    var mesLabel = meses_br[hoje.getMonth()] + '/' + hoje.getFullYear();

    if (pgSheet) {
      var pg = pgSheet.getDataRange().getValues();
      for (var j = 6; j < pg.length; j++) {
        if (String(pg[j][0]) === mesLabel && String(pg[j][3]).toUpperCase() === 'PAGO') {
          Logger.log(ct.inqNome + ' — ' + mesLabel + ' já pago, pulando');
          continue;
        }
      }
    }

    // Montar e-mail
    var fmtV = function(v) {
      return 'R$ ' + parseFloat(v).toFixed(2).replace('.', ',');
    };
    var diaVenc = Utilities.formatDate(
      new Date(hoje.getFullYear(), hoje.getMonth(), diaPgto),
      tz, 'dd/MM/yyyy'
    );

    var assunto = diasAte === 0
      ? '⚠️ Vencimento do aluguel HOJE — ' + ct.imovel
      : '📅 Lembrete: aluguel vence em 3 dias — ' + ct.imovel;

    var corpo = '<div style="font-family:Arial,sans-serif;max-width:500px;background:#f5f5f5;padding:20px;border-radius:12px">'
      + '<div style="background:#0F2438;color:#fff;border-radius:8px;padding:16px;margin-bottom:16px;text-align:center">'
      + '<div style="font-size:22px;font-weight:900">⚡ Fluxo App</div>'
      + '<div style="color:#aaa;font-size:12px">' + (diasAte === 0 ? 'Vencimento hoje' : 'Lembrete de vencimento') + '</div>'
      + '</div>'
      + '<div style="background:#fff;border-radius:8px;padding:16px">'
      + '<p style="font-size:15px">Olá, <strong>' + ct.inqNome + '</strong>!</p>'
      + '<p style="margin:12px 0">Este é um lembrete sobre o vencimento do aluguel:</p>'
      + '<table style="width:100%;border-collapse:collapse">'
      + '<tr><td style="padding:8px;color:#666">Imóvel</td><td style="padding:8px;font-weight:700">' + ct.imovel + '</td></tr>'
      + '<tr style="background:#f9f9f9"><td style="padding:8px;color:#666">Valor</td><td style="padding:8px;font-weight:700;font-size:18px">' + fmtV(ct.valor) + '</td></tr>'
      + '<tr><td style="padding:8px;color:#666">Vencimento</td><td style="padding:8px;font-weight:700;color:' + (diasAte === 0 ? '#e74c3c' : '#e67e22') + '">' + diaVenc + '</td></tr>'
      + '</table>'
      + '<p style="margin-top:12px;font-size:12px;color:#888">Em caso de dúvidas, entre em contato com o locador.</p>'
      + '</div>'
      + '<p style="text-align:center;color:#bbb;font-size:10px;margin-top:12px">Enviado por Fluxo App</p>'
      + '</div>';

    // Enviar para o LOCADOR (aviso interno)
    if (EMAIL_DESTINO) {
      GmailApp.sendEmail(
        EMAIL_DESTINO,
        '[Fluxo] Vencimento: ' + ct.inqNome + ' — ' + diaVenc,
        '',
        { htmlBody: corpo, name: 'Fluxo App' }
      );
    }

    // Enviar para o INQUILINO se tiver email cadastrado
    var inqEmail = ct.inqEmail || '';
    if (inqEmail && inqEmail.indexOf('@') > -1) {
      var corpoInq = '<div style="font-family:Arial,sans-serif;max-width:500px;background:#f5f5f5;padding:20px;border-radius:12px">'
        + '<div style="background:#0F2438;color:#fff;border-radius:8px;padding:16px;margin-bottom:16px;text-align:center">'
        + '<div style="font-size:22px;font-weight:900">⚡ Fluxo App</div>'
        + '<div style="color:#aaa;font-size:12px">Lembrete de vencimento de aluguel</div>'
        + '</div>'
        + '<div style="background:#fff;border-radius:8px;padding:16px">'
        + '<p style="font-size:15px">Olá, <strong>' + ct.inqNome + '</strong>!</p>'
        + '<p style="margin:12px 0">Lembramos que o aluguel do imóvel abaixo vence ' + (diasAte === 0 ? '<strong>hoje</strong>' : 'em <strong>3 dias</strong>') + ':</p>'
        + '<table style="width:100%;border-collapse:collapse">'
        + '<tr><td style="padding:8px;color:#666">Imóvel</td><td style="padding:8px;font-weight:700">' + ct.imovel + '</td></tr>'
        + '<tr style="background:#f9f9f9"><td style="padding:8px;color:#666">Valor</td><td style="padding:8px;font-weight:700;font-size:18px">' + fmtV(ct.valor) + '</td></tr>'
        + '<tr><td style="padding:8px;color:#666">Vencimento</td><td style="padding:8px;font-weight:700;color:' + (diasAte === 0 ? '#e74c3c' : '#e67e22') + '">' + diaVenc + '</td></tr>'
        + '</table>'
        + '<p style="margin-top:16px;font-size:13px">Por favor, realize o pagamento conforme combinado. Em caso de dúvidas, entre em contato com o locador.</p>'
        + '<p style="font-size:12px;color:#888;margin-top:8px">Locador: <strong>' + ct.locNome + '</strong></p>'
        + '</div>'
        + '<p style="text-align:center;color:#bbb;font-size:10px;margin-top:12px">Mensagem automática gerada por Fluxo App</p>'
        + '</div>';

      GmailApp.sendEmail(
        inqEmail,
        (diasAte === 0 ? '⚠️ Vencimento do aluguel HOJE' : '📅 Lembrete: aluguel vence em 3 dias') + ' — ' + ct.imovel,
        '',
        { htmlBody: corpoInq, name: ct.locNome + ' via Fluxo App' }
      );
      Logger.log('Email enviado ao inquilino: ' + inqEmail);
    } else {
      Logger.log('Inquilino ' + ct.inqNome + ' sem email cadastrado');
    }

    // WhatsApp — gerar link wa.me para envio manual (Apps Script não envia WA diretamente)
    // Mas podemos enviar o link por email ao LOCADOR para ele encaminhar
    var inqWhats = ct.inqWhats || '';
    if (inqWhats) {
      var fone = inqWhats.replace(/\D/g, '');
      if (fone.length < 12) fone = '55' + fone; // adicionar DDI Brasil
      var meses_br2 = ['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];
      var msgWpp = encodeURIComponent(
        '📅 *Lembrete de aluguel*\n\n'
        + 'Olá, *' + ct.inqNome + '*!\n\n'
        + (diasAte === 0
            ? '⚠️ Seu aluguel vence *hoje*!'
            : '📅 Seu aluguel vence em *3 dias* (' + diaVenc + ').')
        + '\n\n'
        + '🏠 Imóvel: ' + ct.imovel + '\n'
        + '💰 Valor: ' + fmtV(ct.valor) + '\n'
        + '📅 Vencimento: dia ' + ct.diaPgto + ' de cada mês\n\n'
        + 'Por favor, realize o pagamento conforme combinado.\n\n'
        + '_Mensagem automática — Fluxo App_'
      );
      var linkWpp = 'https://wa.me/' + fone + '?text=' + msgWpp;
      Logger.log('Link WhatsApp para ' + ct.inqNome + ': ' + linkWpp);

      // Enviar link ao locador por email para ele encaminhar
      if (EMAIL_DESTINO) {
        GmailApp.sendEmail(
          EMAIL_DESTINO,
          '[Fluxo] Envie WhatsApp para ' + ct.inqNome + ' — vencimento ' + diaVenc,
          '',
          {
            htmlBody: '<div style="font-family:Arial;padding:20px">'
              + '<h3>⚡ Fluxo — Lembrete para enviar ao inquilino</h3>'
              + '<p>Clique no botão para enviar o lembrete via WhatsApp para <strong>' + ct.inqNome + '</strong>:</p>'
              + '<a href="' + linkWpp + '" style="display:inline-block;margin:12px 0;padding:12px 24px;background:#25D366;color:#fff;border-radius:8px;text-decoration:none;font-weight:bold;font-size:15px">💬 Enviar WhatsApp</a>'
              + '<p style="color:#888;font-size:12px">Ou copie o número: ' + ct.inqWhats + '</p>'
              + '</div>',
            name: 'Fluxo App'
          }
        );
        Logger.log('Email com link WhatsApp enviado ao locador');
      }
    }

    // Push para o locador junto com o email
    try {
      var pushAlug = 'Aluguel de ' + ct.inqNome + ' vence '
        + (diasAte === 0 ? 'hoje' : 'em 3 dias') + ' — ' + fmtV(ct.valor);
      enviarPush('🏠 Fluxo — Vencimento de aluguel', pushAlug);
    } catch(ep2) { Logger.log('Push aluguel falhou: ' + ep2.message); }

    enviados++;
    Logger.log('Lembrete processado para ' + ct.inqNome + ' — vence em ' + diasAte + ' dias');
  }
  Logger.log('Total de lembretes enviados: ' + enviados);
}

function criarTriggerLembreteAluguel() {
  // Remove triggers antigos
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarLembreteAluguel') ScriptApp.deleteTrigger(t);
  });
  // Criar trigger diário às 8h (junto com o resumo diário)
  ScriptApp.newTrigger('enviarLembreteAluguel').timeBased().everyDays(1).atHour(8).create();
  Logger.log('✅ Trigger de lembrete de aluguel criado — roda todo dia às 8h.');
}

// ════════════════════════════════════════════════════════════
//  FUNÇÕES DE TESTE — Execute manualmente no editor
// ════════════════════════════════════════════════════════════

// PASSO 1: Configure seu email real aqui e no topo do arquivo
// var EMAIL_DESTINO = 'seu.email@gmail.com';

function testarEmailTarefas() {
  // Simula o enviarResumoDiario independente do horário
  Logger.log('=== TESTE: Email de Tarefas ===');
  Logger.log('EMAIL_DESTINO atual: ' + EMAIL_DESTINO);
  
  if (EMAIL_DESTINO === 'SEU_EMAIL@gmail.com') {
    Logger.log('❌ ERRO: Configure EMAIL_DESTINO no topo do arquivo Code.gs!');
    return;
  }
  
  // Enviar email de teste simples
  GmailApp.sendEmail(
    EMAIL_DESTINO,
    '✅ [Fluxo] Teste — Notificação de Tarefas funcionando!',
    '',
    {
      htmlBody: '<div style="font-family:Arial;padding:20px;background:#f5f5f5;border-radius:12px">'
        + '<h2 style="color:#0F2438">⚡ Fluxo App — Teste de Notificação</h2>'
        + '<p>✅ O sistema de email de <strong>tarefas agendadas</strong> está funcionando!</p>'
        + '<p>Enviado em: ' + new Date().toLocaleString() + '</p>'
        + '<hr>'
        + '<p style="color:#888;font-size:12px">Este é um email de teste. O email real é enviado todo dia às 8h com suas tarefas do dia.</p>'
        + '</div>',
      name: 'Fluxo App'
    }
  );
  Logger.log('✅ Email de teste enviado para: ' + EMAIL_DESTINO);
}

function testarEmailInquilino() {
  // Simula o enviarLembreteAluguel para um contrato específico
  Logger.log('=== TESTE: Email ao Inquilino ===');
  
  var sheet = ss().getSheetByName('Contratos');
  if (!sheet) { Logger.log('❌ Aba Contratos não encontrada'); return; }
  
  var dados = sheet.getDataRange().getValues();
  var ct = null;
  for (var i = 1; i < dados.length; i++) {
    if (!dados[i][1]) continue;
    try { ct = JSON.parse(String(dados[i][1])); break; } catch(e) {}
  }
  
  if (!ct) { Logger.log('❌ Nenhum contrato encontrado na planilha'); return; }
  
  Logger.log('Contrato encontrado: ' + ct.inqNome);
  Logger.log('Email inquilino: ' + (ct.inqEmail || 'NÃO CADASTRADO'));
  Logger.log('Email locador (EMAIL_DESTINO): ' + EMAIL_DESTINO);
  
  if (EMAIL_DESTINO === 'SEU_EMAIL@gmail.com') {
    Logger.log('❌ Configure EMAIL_DESTINO no topo do arquivo!');
    return;
  }
  
  var corpo = '<div style="font-family:Arial;padding:20px;background:#f5f5f5;border-radius:12px">'
    + '<h2 style="color:#0F2438">⚡ Fluxo App — Teste de Lembrete</h2>'
    + '<p>✅ O sistema de <strong>lembrete de aluguel</strong> está funcionando!</p>'
    + '<p><strong>Inquilino:</strong> ' + ct.inqNome + '</p>'
    + '<p><strong>Imóvel:</strong> ' + (ct.imovel || '—') + '</p>'
    + '<p><strong>Valor:</strong> R$ ' + parseFloat(ct.valor||0).toFixed(2).replace('.', ',') + '</p>'
    + '<p><strong>Vencimento:</strong> todo dia ' + (ct.diaPgto || '—') + '</p>'
    + '<hr>'
    + '<p style="color:#888;font-size:12px">Teste enviado em: ' + new Date().toLocaleString() + '</p>'
    + '</div>';
  
  // Enviar para você (locador)
  GmailApp.sendEmail(EMAIL_DESTINO, '✅ [Fluxo] Teste — Lembrete de aluguel de ' + ct.inqNome, '', { htmlBody: corpo, name: 'Fluxo App' });
  Logger.log('✅ Email enviado ao locador: ' + EMAIL_DESTINO);
  
  // Enviar para o inquilino se tiver email
  if (ct.inqEmail && ct.inqEmail.indexOf('@') > -1) {
    GmailApp.sendEmail(ct.inqEmail, '✅ [Fluxo] Teste — Lembrete de vencimento', '', { htmlBody: corpo, name: ct.locNome + ' via Fluxo App' });
    Logger.log('✅ Email enviado ao inquilino: ' + ct.inqEmail);
  } else {
    Logger.log('⚠️ Inquilino sem email — cadastre o campo "Email do inquilino" no contrato');
  }
}

function testarTudoEmail() {
  testarEmailTarefas();
  testarEmailInquilino();
}

// ════════════════════════════════════════════════════════════
//  PUSH NOTIFICATIONS — Firebase Cloud Messaging API v1
//  Autenticação via Service Account (não usa Server Key)
// ════════════════════════════════════════════════════════════

// Chaves FCM lidas do Properties Service (nunca expostas no código)
// Para configurar, execute no console do Apps Script:
// PropertiesService.getScriptProperties().setProperties({
//   'FCM_PROJECT_ID': 'fluxo-app-46562',
//   'FCM_CLIENT_EMAIL': 'firebase-adminsdk-fbsvc@fluxo-app-46562.iam.gserviceaccount.com',
//   'FCM_PRIVATE_KEY': '-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n'
// });
var _FCM_PROPS_      = PropertiesService.getScriptProperties();
var FCM_PROJECT_ID   = _FCM_PROPS_.getProperty('FCM_PROJECT_ID')   || '';
var FCM_CLIENT_EMAIL = _FCM_PROPS_.getProperty('FCM_CLIENT_EMAIL') || '';
var FCM_PRIVATE_KEY  = _FCM_PROPS_.getProperty('FCM_PRIVATE_KEY')  || '';

// Salva o token FCM do dispositivo (chamado automaticamente pelo app)
function salvarFcmToken(body) {
  Logger.log('salvarFcmToken chamado, token: ' + String(body.token || '').substring(0,20));
  var sheet = ss().getSheetByName('FCMTokens');
  if (!sheet) {
    sheet = ss().insertSheet('FCMTokens');
    sheet.appendRow(['token', 'updatedAt']);
    sheet.setFrozenRows(1);
    Logger.log('Aba FCMTokens criada');
  }
  var token = String(body.token || '').trim();
  if (!token) { Logger.log('Token vazio!'); return { ok: false, error: 'token vazio' }; }

  // Verificar se já existe
  var dados = sheet.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0] === token) {
      sheet.getRange(i + 1, 2).setValue(new Date().toISOString());
      return { ok: true, updated: true };
    }
  }
  sheet.appendRow([token, new Date().toISOString()]);
  return { ok: true, created: true };
}

// Gera JWT para autenticação OAuth2 com Service Account
function getFcmAccessToken_() {
  var now   = Math.floor(Date.now() / 1000);
  var claim = {
    iss:   FCM_CLIENT_EMAIL,
    scope: 'https://www.googleapis.com/auth/firebase.messaging',
    aud:   'https://oauth2.googleapis.com/token',
    iat:   now,
    exp:   now + 3600
  };

  // Criar JWT manualmente (Apps Script não tem crypto nativo para RSA)
  // Usamos o serviço OAuth2 do Apps Script via ScriptApp
  var header  = Utilities.base64EncodeWebSafe(JSON.stringify({alg:'RS256',typ:'JWT'}));
  var payload = Utilities.base64EncodeWebSafe(JSON.stringify(claim));
  var toSign  = header + '.' + payload;

  // Assinar com chave privada usando Utilities.computeRsaSha256Signature
  var key  = FCM_PRIVATE_KEY;
  var sign = Utilities.computeRsaSha256Signature(toSign, key);
  var jwt  = toSign + '.' + Utilities.base64EncodeWebSafe(sign);

  // Trocar JWT por access token
  var resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method:  'post',
    payload: { grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion: jwt },
    muteHttpExceptions: true
  });
  var data = JSON.parse(resp.getContentText());
  if (!data.access_token) {
    Logger.log('Erro ao obter access token: ' + resp.getContentText());
    return null;
  }
  return data.access_token;
}

// Envia push para todos os dispositivos registrados (FCM API v1)
function enviarPush(titulo, mensagem) {
  var sheet = ss().getSheetByName('FCMTokens');
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('Nenhum token FCM cadastrado — abra o app primeiro');
    return;
  }
  var tokens = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function(r) { return String(r[0]).trim(); })
    .filter(function(t) { return t.length > 0; });

  if (!tokens.length) { Logger.log('Nenhum token válido'); return; }

  var accessToken = getFcmAccessToken_();
  if (!accessToken) { Logger.log('Falha na autenticação FCM'); return; }

  var url = 'https://fcm.googleapis.com/v1/projects/' + FCM_PROJECT_ID + '/messages:send';
  var resultados = [];

  // FCM v1 envia um token por vez
  tokens.forEach(function(token) {
    var body = JSON.stringify({
      message: {
        token: token,
        notification: { title: titulo, body: mensagem },
        webpush: {
          notification: {
            title: titulo,
            body:  mensagem,
            icon:  '/icon.svg',
            badge: '/icon.svg',
            requireInteraction: false
          },
          fcm_options: { link: 'https://seu-usuario.github.io/fluxo-app/' }
        }
      }
    });

    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + accessToken },
      payload: body,
      muteHttpExceptions: true
    });
    var result = JSON.parse(resp.getContentText());
    Logger.log('Push para ' + token.substring(0,20) + '...: ' + JSON.stringify(result));
    resultados.push(result);
  });

  return resultados;
}

// Adicionar push ao resumo diário de tarefas
function enviarPushResumoDiario() {
  // Esta função é chamada por enviarResumoDiario() com os dados já coletados
  // Não deve ser chamada diretamente — use enviarResumoDiario()
  Logger.log('Use enviarResumoDiario() para enviar push + email juntos');
  enviarResumoDiario();
}

// Testar push manualmente
function testarPush() {
  Logger.log('Enviando push de teste...');
  var result = enviarPush('⚡ Fluxo — Teste', 'Push notifications funcionando! 🎉');
  Logger.log('Resultado: ' + JSON.stringify(result));
}

// ════════════════════════════════════════════════════════════
//  VALOR VARIÁVEL — Fatura do cartão = soma dos gastos do mês
//  Chamado pelo app para obter o valor real da fatura
// ════════════════════════════════════════════════════════════
function getValorFaturaCartao(body) {
  var cartaoId = String(body.cartaoId || '');
  var mesAno   = String(body.mesAno   || '');
  if (!cartaoId || !mesAno) return { ok: false, error: 'Parâmetros inválidos' };

  var sheet = ss().getSheetByName('GastosCartao');
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, total: 0, gastos: [] };

  var dados   = sheet.getDataRange().getValues();
  var total   = 0;
  var gastos  = [];
  var porCat  = {};

  for (var i = 1; i < dados.length; i++) {
    if (!dados[i][1]) continue;
    try {
      var g = JSON.parse(String(dados[i][1]));
      if (String(g.cartaoId) === cartaoId && g.fatMes === mesAno && !g.pago) {
        total += parseFloat(g.val || 0);
        porCat[g.cat] = (porCat[g.cat] || 0) + parseFloat(g.val || 0);
        gastos.push({ desc: g.desc, val: g.val, cat: g.cat, data: g.data });
      }
    } catch(e) {}
  }
  return { ok: true, total: total, gastos: gastos, porCat: porCat };
}

// ════════════════════════════════════════════════════════════
//  TESTE DE PERMISSÃO DO DRIVE
//  Execute testarDrive() uma vez para autorizar o acesso
// ════════════════════════════════════════════════════════════
function testarDrive() {
  try {
    // Tenta criar/acessar a pasta
    var folders = DriveApp.getFoldersByName('Comprovantes Família');
    var folder = folders.hasNext()
      ? folders.next()
      : DriveApp.createFolder('Comprovantes Família');

    // Cria arquivo de teste
    var blob = Utilities.newBlob('Teste de permissão - Fluxo App', 'text/plain', 'teste-permissao.txt');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log('✅ Drive OK! Pasta: ' + folder.getName());
    Logger.log('✅ Arquivo teste: ' + file.getUrl());
    Logger.log('✅ Agora os comprovantes serão salvos nesta pasta.');

    // Remover arquivo de teste
    file.setTrashed(true);
    return { ok: true, folder: folder.getName(), id: folder.getId() };
  } catch(e) {
    Logger.log('❌ Erro Drive: ' + e.message);
    Logger.log('Solução: Executar > Autorizar acesso no menu do Apps Script');
    return { ok: false, error: e.message };
  }
}

// ════════════════════════════════════════════════════════════
//  DIAGNÓSTICO — execute para ver o estado da planilha
// ════════════════════════════════════════════════════════════
function diagnostico() {
  var abas = ['Contratos','Aluguéis','Pagadores','Tarefas','Transações','Recorrentes','GastosCartao'];
  abas.forEach(function(nome) {
    var sheet = ss().getSheetByName(nome);
    if (!sheet) {
      Logger.log('❌ Aba não encontrada: ' + nome);
    } else {
      Logger.log('✅ ' + nome + ' — ' + (sheet.getLastRow()-1) + ' registros');
    }
  });

  // Listar todas as abas existentes
  var todas = ss().getSheets().map(function(s){ return s.getName(); });
  Logger.log('Todas as abas: ' + todas.join(', '));
}

// ════════════════════════════════════════════════════════════
//  DIAGNÓSTICO — encontrar despesas sem categoria correta
//  Execute para listar recorrências e tarefas com categoria
//  vazia, que por isso caem em "📦 Outros" nos gráficos.
// ════════════════════════════════════════════════════════════
function diagnosticoCategorias() {
  // ── Recorrentes ────────────────────────────────────────
  var rSheet = ss().getSheetByName('Recorrentes');
  if (rSheet) {
    var rDados = rSheet.getDataRange().getValues();
    var rHeaders = rDados[0];
    var idxDesc = rHeaders.indexOf('desc');
    var idxCat  = rHeaders.indexOf('cat');
    var idxVal  = rHeaders.indexOf('value');
    Logger.log('═══ RECORRENTES sem categoria ═══');
    var semCatRecur = [];
    for (var i = 1; i < rDados.length; i++) {
      var cat = String(rDados[i][idxCat] || '').trim();
      if (!cat) {
        semCatRecur.push(rDados[i][idxDesc] + ' (R$ ' + rDados[i][idxVal] + ')');
      }
    }
    if (semCatRecur.length) {
      Logger.log(semCatRecur.length + ' recorrência(s) SEM categoria:');
      semCatRecur.forEach(function(s){ Logger.log('  • ' + s); });
    } else {
      Logger.log('✅ Todas as recorrências têm categoria definida.');
    }
  }

  // ── Tarefas ────────────────────────────────────────────
  var tSheet = ss().getSheetByName('Tarefas');
  if (tSheet) {
    var tDados = tSheet.getDataRange().getValues();
    var tHeaders = tDados[0];
    var idxDescT = tHeaders.indexOf('desc');
    var idxCatT  = tHeaders.indexOf('cat');
    var idxValT  = tHeaders.indexOf('value');
    var idxTypeT = tHeaders.indexOf('type');

    Logger.log('');
    Logger.log('═══ TAREFAS sem categoria (agrupado por descrição) ═══');
    var semCatTarefas = {}; // desc -> {count, total}
    for (var j = 1; j < tDados.length; j++) {
      var tipo = String(tDados[j][idxTypeT] || '');
      if (tipo === 'task') continue; // tarefas sem valor financeiro não importam aqui
      var catT = String(tDados[j][idxCatT] || '').trim();
      if (!catT) {
        var descT = String(tDados[j][idxDescT] || '(sem descrição)');
        if (!semCatTarefas[descT]) semCatTarefas[descT] = { count: 0, total: 0 };
        semCatTarefas[descT].count++;
        semCatTarefas[descT].total += parseFloat(tDados[j][idxValT]) || 0;
      }
    }
    var chaves = Object.keys(semCatTarefas);
    if (chaves.length) {
      Logger.log(chaves.length + ' descrição(ões) diferentes SEM categoria:');
      chaves.sort(function(a,b){ return semCatTarefas[b].total - semCatTarefas[a].total; });
      chaves.forEach(function(desc){
        var info = semCatTarefas[desc];
        Logger.log('  • ' + desc + ' — ' + info.count + 'x — total R$ ' + info.total.toFixed(2));
      });
    } else {
      Logger.log('✅ Todas as tarefas financeiras têm categoria definida.');
    }
  }
}

// ════════════════════════════════════════════════════════════
//  CORREÇÃO EM MASSA — define a categoria de todas as tarefas
//  E recorrências que tenham uma DESCRIÇÃO específica.
//  Edite o mapa abaixo com suas próprias descrições e categorias,
//  depois execute esta função uma vez.
// ════════════════════════════════════════════════════════════
function corrigirCategoriasEmMassa() {
  // ★ EDITE AQUI — descrição EXATA (como aparece na planilha) → categoria
  var mapa = {
    // 'Cartão de Crédito Nubank':  '💳 Cartão',
    // 'Internet Casa':             '🏠 Casa',
    // 'Celular TIM':               '📱 Telefone',
  };

  var chaves = Object.keys(mapa);
  if (!chaves.length) {
    Logger.log('⚠️ Mapa vazio — edite a função corrigirCategoriasEmMassa() com suas descrições e categorias antes de executar.');
    return;
  }

  function normalizar(s) {
    return String(s||'').trim().toLowerCase();
  }
  var mapaNorm = {};
  chaves.forEach(function(k){ mapaNorm[normalizar(k)] = mapa[k]; });

  var totalAtualizado = 0;

  // Recorrentes
  var rSheet = ss().getSheetByName('Recorrentes');
  if (rSheet) {
    var rDados = rSheet.getDataRange().getValues();
    var idxDesc = rDados[0].indexOf('desc');
    var idxCat  = rDados[0].indexOf('cat');
    for (var i = 1; i < rDados.length; i++) {
      var key = normalizar(rDados[i][idxDesc]);
      if (mapaNorm[key]) {
        rSheet.getRange(i+1, idxCat+1).setValue(mapaNorm[key]);
        totalAtualizado++;
      }
    }
  }

  // Tarefas
  var tSheet = ss().getSheetByName('Tarefas');
  if (tSheet) {
    var tDados = tSheet.getDataRange().getValues();
    var idxDescT = tDados[0].indexOf('desc');
    var idxCatT  = tDados[0].indexOf('cat');
    for (var j = 1; j < tDados.length; j++) {
      var keyT = normalizar(tDados[j][idxDescT]);
      if (mapaNorm[keyT]) {
        tSheet.getRange(j+1, idxCatT+1).setValue(mapaNorm[keyT]);
        totalAtualizado++;
      }
    }
  }

  Logger.log('✅ ' + totalAtualizado + ' linha(s) atualizada(s) com nova categoria.');
}

// ════════════════════════════════════════════════════════════
//  TESTE DE SYNC — execute para ver o que cada endpoint retorna
// ════════════════════════════════════════════════════════════
function testarSync() {
  var abas = ['Contratos','Aluguéis','Pagadores'];
  abas.forEach(function(nome) {
    var result = getItemsSheet(nome);
    Logger.log('=== ' + nome + ' ===');
    Logger.log('ok: ' + result.ok + ' | itens: ' + result.data.length);
    if (result.data.length > 0) {
      Logger.log('Primeiro item: ' + JSON.stringify(result.data[0]).substring(0, 200));
    } else {
      // Verificar o conteúdo bruto da aba
      var sheet = ss().getSheetByName(nome);
      if (sheet && sheet.getLastRow() > 1) {
        var raw = sheet.getRange(2, 1, Math.min(3, sheet.getLastRow()-1), 3).getValues();
        Logger.log('Conteúdo bruto (3 linhas): ' + JSON.stringify(raw).substring(0, 300));
      }
    }
  });
}
