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
  var id = Date.now();
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

  // Garantir colunas cat e recurId existem
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('cat') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('cat');
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  if (headers.indexOf('recurId') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('recurId');
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

  sheet.appendRow(row);
  return { ok: true, id: id };
}

function updateTask(body) {
  var r = sheetRows('Tarefas');
  var statusIdx    = r.headers.indexOf('status');
  var comprovIdx   = r.headers.indexOf('comprovUrl');

  // Criar coluna comprovUrl se não existir
  if (comprovIdx === -1 && body.comprovUrl) {
    r.sheet.getRange(1, r.headers.length + 1).setValue('comprovUrl');
    comprovIdx = r.headers.length;
    r.headers.push('comprovUrl');
  }

  for (var i = 0; i < r.rows.length; i++) {
    if (String(r.rows[i][0]) === String(body.id)) {
      r.sheet.getRange(i + 2, statusIdx + 1).setValue(body.status);
      if (body.comprovUrl && comprovIdx > -1) {
        r.sheet.getRange(i + 2, comprovIdx + 1).setValue(body.comprovUrl);
      }
      return { ok: true };
    }
  }
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
    sheet.appendRow(['id','type','desc','value','cat','date','createdAt']);
    sheet.setFrozenRows(1);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#162030').setFontColor('#FFF');
  }
  var id = body.id || Date.now();
  sheet.appendRow([id, body.type, body.desc, parseFloat(body.value), body.cat, body.date, new Date().toISOString()]);
  return { ok: true, id: id };
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
var EMAIL_DESTINO = 'armbr258@gmail.com';

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
  function fmtVal(v) {
    if (!v || v <= 0) return '';
    return ' — <strong>R$ ' + v.toFixed(2).replace('.', ',') + '</strong>';
  }
  function tipoIcon(t) {
    if (t === 'exp')  return '🔴';
    if (t === 'inc')  return '🟢';
    return '📌';
  }
  function fmtBRLTotal(v) {
    return 'R$ ' + v.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }

  var totalHoje     = tarefas.reduce(function(s,t){   return s+(t.type!=='task'?t.value:0);}, 0);
  var totalAtrasado = atrasadas.reduce(function(s,t){ return s+(t.type!=='task'?t.value:0);}, 0);

  function linhaTabela(t, cor) {
    var dl = t.deadline ? t.deadline.split('-').reverse().slice(0,2).join('/') : '';
    return '<tr>'
      + '<td style="padding:10px 12px;border-bottom:1px solid #2a2a2a;font-size:18px;width:30px">'+tipoIcon(t.type)+'</td>'
      + '<td style="padding:10px 12px;border-bottom:1px solid #2a2a2a">'
      +   '<div style="font-weight:700;color:'+(cor||'#eee')+'">'+t.desc+'</div>'
      +   '<div style="font-size:12px;color:#888;margin-top:2px">'+(t.cat||'')+(dl&&cor?' · '+dl:'')+'</div>'
      + '</td>'
      + '<td style="padding:10px 12px;border-bottom:1px solid #2a2a2a;text-align:right;font-weight:800;color:'+(t.value>0?(t.type==='inc'?'#2ecc9a':'#FF6B6B'):'#555')+'">'
      +   (t.value>0?fmtBRLTotal(t.value):'')
      + '</td>'
      + '</tr>';
  }

  var secaoHoje = '';
  if (tarefas.length) {
    secaoHoje = '<div style="padding:20px 24px">'
      + '<div style="font-size:10px;font-weight:800;color:#888;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px;display:flex;justify-content:space-between">'
      +   '<span>📅 Compromissos de hoje</span>'
      +   (totalHoje>0?'<span style="color:#FF6B6B;font-size:13px">'+fmtBRLTotal(totalHoje)+'</span>':'')
      + '</div>'
      + '<table style="width:100%;border-collapse:collapse">'
      + tarefas.map(function(t){ return linhaTabela(t, '#eee'); }).join('')
      + '</table></div>';
  }

  var secaoAtrasado = '';
  if (atrasadas.length) {
    secaoAtrasado = '<div style="padding:20px 24px;border-top:1px solid #222">'
      + '<div style="font-size:10px;font-weight:800;color:#FF6B6B;letter-spacing:1px;text-transform:uppercase;margin-bottom:12px;display:flex;justify-content:space-between">'
      +   '<span>⚠️ Atrasados</span>'
      +   (totalAtrasado>0?'<span style="font-size:13px">'+fmtBRLTotal(totalAtrasado)+'</span>':'')
      + '</div>'
      + '<table style="width:100%;border-collapse:collapse">'
      + atrasadas.map(function(t){ return linhaTabela(t, '#FF6B6B'); }).join('')
      + '</table></div>';
  }

  var html = '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#111;color:#eee;border-radius:12px;overflow:hidden;border:1px solid #222">'
    + '<div style="background:#0f0f1a;padding:22px 24px;border-bottom:1px solid #222">'
    +   '<div style="font-size:20px;font-weight:800">⚡ Fluxo App</div>'
    +   '<div style="color:#888;margin-top:4px;font-size:13px">'+diaSem+', '+diaFmt+'</div>'
    + '</div>'
    + secaoHoje
    + secaoAtrasado
    + '<div style="padding:14px 24px;border-top:1px solid #222;color:#444;font-size:11px;text-align:center">Resumo automático diário · Fluxo App</div>'
    + '</div>';

  GmailApp.sendEmail(EMAIL_DESTINO,
    '⚡ Fluxo — ' + (tarefas.length ? tarefas.length+' compromisso'+(tarefas.length>1?'s':'') : '') +
    (atrasadas.length ? (tarefas.length?' · ':'') + atrasadas.length+' atrasado'+(atrasadas.length>1?'s':'') : '') +
    ' · ' + diaFmt,
    'Abra no Gmail para ver o e-mail formatado.',
    { htmlBody: html, name: 'Fluxo App' }
  );

  Logger.log('✅ E-mail enviado para ' + EMAIL_DESTINO);
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
//  PÁGINA PÚBLICA DO INQUILINO
//  Retorna HTML com histórico de pagamentos do contrato
// ════════════════════════════════════════════════════════════

function paginaInquilino(e) {
  var ctId = e.parameter.ctId || '';

  // Buscar contrato
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
  var sheet = ss().getSheetByName('FCMTokens');
  if (!sheet) {
    sheet = ss().insertSheet('FCMTokens');
    sheet.appendRow(['token', 'updatedAt']);
    sheet.setFrozenRows(1);
  }
  var token = String(body.token || '').trim();
  if (!token) return { ok: false };

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
  var hoje = new Date();
  var diaHoje = hoje.getDate();
  var mesHoje = hoje.getMonth();
  var anoHoje = hoje.getFullYear();

  // Contar tarefas do dia
  var sheetTasks = ss().getSheetByName('Tarefas');
  var total = 0;
  if (sheetTasks && sheetTasks.getLastRow() > 1) {
    var tasks = sheetTasks.getRange(2,1,sheetTasks.getLastRow()-1,sheetTasks.getLastColumn()).getValues();
    var headers = sheetTasks.getRange(1,1,1,sheetTasks.getLastColumn()).getValues()[0];
    var dlIdx = headers.indexOf('deadline');
    var stIdx = headers.indexOf('status');
    var hoje_str = anoHoje + '-' + String(mesHoje+1).padStart(2,'0') + '-' + String(diaHoje).padStart(2,'0');
    tasks.forEach(function(r) {
      if (String(r[dlIdx]).substring(0,10) === hoje_str && r[stIdx] !== 'done') total++;
    });
  }

  var msg = total > 0
    ? total + ' compromisso' + (total>1?'s':'') + ' para hoje'
    : 'Sem compromissos hoje 🎉';

  enviarPush('⚡ Fluxo — Resumo de hoje', msg);
}

// Testar push manualmente
function testarPush() {
  Logger.log('Enviando push de teste...');
  var result = enviarPush('⚡ Fluxo — Teste', 'Push notifications funcionando! 🎉');
  Logger.log('Resultado: ' + JSON.stringify(result));
}
